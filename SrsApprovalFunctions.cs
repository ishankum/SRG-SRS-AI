using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Azure.Storage.Sas;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SrsAi.Functions
{
    public class SrsApprovalFunctions
    {
        private readonly ILogger _logger;
        private readonly IConfiguration _configuration;
        private readonly BlobServiceClient _blobServiceClient;
        private readonly EmailNotificationService _emailService;
        private readonly TeamsNotificationService _teamsService;

        public SrsApprovalFunctions(ILoggerFactory loggerFactory, IConfiguration configuration)
        {
            _logger = loggerFactory.CreateLogger<SrsApprovalFunctions>();
            _configuration = configuration;

            string connectionString = _configuration["AzureWebJobsStorage"] ?? "";
            _blobServiceClient = new BlobServiceClient(connectionString);

            _emailService = new EmailNotificationService(_configuration, _logger);
            _teamsService = new TeamsNotificationService(_configuration, _logger);
        }

        [Function("GetSrsEditor")]
        public async Task<HttpResponseData> GetEditor(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "srs/editor")] HttpRequestData req)
        {
            var query = System.Web.HttpUtility.ParseQueryString(req.Url.Query);
            string draftId = query["draftId"] ?? "";

            if (string.IsNullOrWhiteSpace(draftId))
            {
                var badResp = req.CreateResponse(HttpStatusCode.BadRequest);
                await badResp.WriteStringAsync("draftId query parameter is required.");
                return badResp;
            }

            string jsonBlobName = draftId.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                ? draftId.Substring(0, draftId.Length - 5) + ".json"
                : (draftId.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ? draftId : draftId + ".json");

            var draftContainer = _blobServiceClient.GetBlobContainerClient("srs-drafts");
            var blobClient = draftContainer.GetBlobClient(jsonBlobName);

            if (!await blobClient.ExistsAsync())
            {
                var notFound = req.CreateResponse(HttpStatusCode.NotFound);
                await notFound.WriteStringAsync($"Draft '{jsonBlobName}' was not found in srs-drafts container.");
                return notFound;
            }

            var download = await blobClient.DownloadContentAsync();
            string jsonContent = download.Value.Content.ToString();

            string html = BuildEditorHtml(jsonBlobName, jsonContent);

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "text/html; charset=utf-8");
            await response.WriteStringAsync(html);
            return response;
        }

        [Function("SaveSrsDraft")]
        public async Task<HttpResponseData> SaveDraft(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "srs/save-draft")] HttpRequestData req)
        {
            try
            {
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                using var doc = JsonDocument.Parse(requestBody);
                var root = doc.RootElement;

                string draftId = root.GetProperty("draftId").GetString() ?? "";
                string approverName = root.TryGetProperty("approverName", out var aProp) ? aProp.GetString() ?? "Reviewer" : "Reviewer";
                string jsonPayload = root.GetProperty("srsPayload").GetRawText();

                string jsonBlobName = draftId.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                    ? draftId.Substring(0, draftId.Length - 5) + ".json"
                    : (draftId.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ? draftId : draftId + ".json");

                string docxBlobName = jsonBlobName.Substring(0, jsonBlobName.Length - 5) + ".docx";

                var draftContainer = _blobServiceClient.GetBlobContainerClient("srs-drafts");
                await draftContainer.CreateIfNotExistsAsync();

                // 1. Update JSON in srs-drafts/
                var jsonClient = draftContainer.GetBlobClient(jsonBlobName);
                await jsonClient.UploadAsync(BinaryData.FromString(jsonPayload), overwrite: true);

                // 2. Rebuild Word document (.docx)
                string meetingId = ExtractMeetingId(jsonBlobName);
                byte[] docxBytes = SrsWordDocumentBuilder.BuildWordDocument(jsonPayload, meetingId);

                var docxClient = draftContainer.GetBlobClient(docxBlobName);
                var docxOptions = new Azure.Storage.Blobs.Models.BlobUploadOptions
                {
                    HttpHeaders = new Azure.Storage.Blobs.Models.BlobHttpHeaders
                    {
                        ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    }
                };
                await docxClient.UploadAsync(BinaryData.FromBytes(docxBytes), docxOptions);

                _logger.LogInformation($"Successfully saved draft updates to '{jsonBlobName}' and '{docxBlobName}' by {approverName}.");

                var response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteAsJsonAsync(new { success = true, message = "Draft successfully updated in storage." });
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to save draft: {ex.Message}");
                var errResp = req.CreateResponse(HttpStatusCode.InternalServerError);
                await errResp.WriteAsJsonAsync(new { success = false, error = ex.Message });
                return errResp;
            }
        }

        [Function("SaveAndApproveSrs")]
        public async Task<HttpResponseData> SaveAndApprove(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "srs/save-approve")] HttpRequestData req)
        {
            try
            {
                string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                using var doc = JsonDocument.Parse(requestBody);
                var root = doc.RootElement;

                string draftId = root.GetProperty("draftId").GetString() ?? "";
                string project = root.TryGetProperty("project", out var pProp) ? pProp.GetString() ?? "RSCS" : "RSCS";
                string approverName = root.TryGetProperty("approverName", out var aProp) ? aProp.GetString() ?? "Authorized Reviewer" : "Authorized Reviewer";
                string jsonPayload = root.GetProperty("srsPayload").GetRawText();

                return await ExecuteApprovalAsync(req, draftId, project, approverName, jsonPayload);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to save and approve SRS: {ex.Message}");
                var errResp = req.CreateResponse(HttpStatusCode.InternalServerError);
                await errResp.WriteAsJsonAsync(new { success = false, error = ex.Message });
                return errResp;
            }
        }

        [Function("OneClickApproveSrs")]
        public async Task<HttpResponseData> OneClickApprove(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "srs/approve")] HttpRequestData req)
        {
            var query = System.Web.HttpUtility.ParseQueryString(req.Url.Query);
            string draftId = query["draftId"] ?? "";
            string approverName = query["approver"] ?? "Team Reviewer";

            if (string.IsNullOrWhiteSpace(draftId))
            {
                var badResp = req.CreateResponse(HttpStatusCode.BadRequest);
                await badResp.WriteStringAsync("draftId query parameter is required.");
                return badResp;
            }

            string jsonBlobName = draftId.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                ? draftId.Substring(0, draftId.Length - 5) + ".json"
                : (draftId.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ? draftId : draftId + ".json");

            var draftContainer = _blobServiceClient.GetBlobContainerClient("srs-drafts");
            var jsonClient = draftContainer.GetBlobClient(jsonBlobName);

            if (!await jsonClient.ExistsAsync())
            {
                var notFound = req.CreateResponse(HttpStatusCode.NotFound);
                await notFound.WriteStringAsync($"Draft '{jsonBlobName}' was not found.");
                return notFound;
            }

            var download = await jsonClient.DownloadContentAsync();
            string jsonPayload = download.Value.Content.ToString();

            string project = "RSCS";
            try
            {
                using var pDoc = JsonDocument.Parse(jsonPayload);
                if (pDoc.RootElement.TryGetProperty("project", out var prjProp))
                {
                    project = prjProp.GetString() ?? "RSCS";
                }
            }
            catch { }

            return await ExecuteApprovalAsync(req, draftId, project, approverName, jsonPayload, isBrowserHtml: true);
        }

        private async Task<HttpResponseData> ExecuteApprovalAsync(
            HttpRequestData req,
            string draftId,
            string project,
            string approverName,
            string jsonPayload,
            bool isBrowserHtml = false)
        {
            string jsonBlobName = draftId.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                ? draftId.Substring(0, draftId.Length - 5) + ".json"
                : (draftId.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ? draftId : draftId + ".json");

            string meetingId = ExtractMeetingId(jsonBlobName);
            string approvedPrefix = $"{project.ToLower()}/srs-{meetingId}-{DateTime.UtcNow:yyyyMMddHHmmss}";
            string approvedJsonName = $"{approvedPrefix}.json";
            string approvedDocxName = $"{approvedPrefix}.docx";

            // Rebuild final .docx from the payload
            byte[] docxBytes = SrsWordDocumentBuilder.BuildWordDocument(jsonPayload, meetingId);

            var approvedContainer = _blobServiceClient.GetBlobContainerClient("srs-approved");
            await approvedContainer.CreateIfNotExistsAsync();

            // 1. Upload approved JSON and DOCX to srs-approved/
            var appJsonClient = approvedContainer.GetBlobClient(approvedJsonName);
            await appJsonClient.UploadAsync(BinaryData.FromString(jsonPayload), overwrite: true);

            var appDocxClient = approvedContainer.GetBlobClient(approvedDocxName);
            var docxOptions = new Azure.Storage.Blobs.Models.BlobUploadOptions
            {
                HttpHeaders = new Azure.Storage.Blobs.Models.BlobHttpHeaders
                {
                    ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                }
            };
            await appDocxClient.UploadAsync(BinaryData.FromBytes(docxBytes), docxOptions);

            // 2. Generate SAS link for the approved document
            string finalDocxSasUrl = GenerateReadSasUri(appDocxClient);

            // 3. Dispatch Approval Notification Email with attached .docx
            await _emailService.SendSrsApprovedEmailAsync(
                meetingId,
                project,
                approverName,
                jsonPayload,
                docxBytes,
                $"SRS-Approved-{project}-{meetingId}.docx"
            );

            // 4. Post celebration/approval confirmation card to Teams
            await _teamsService.SendApprovalConfirmationAsync(
                project,
                meetingId,
                approverName,
                approvedDocxName,
                finalDocxSasUrl
            );

            _logger.LogInformation($"Successfully approved and promoted SRS '{draftId}' for project '{project}' by {approverName}.");

            if (isBrowserHtml)
            {
                var htmlResponse = req.CreateResponse(HttpStatusCode.OK);
                htmlResponse.Headers.Add("Content-Type", "text/html; charset=utf-8");
                string successHtml = $@"
<!DOCTYPE html>
<html>
<head>
    <title>SRS Approved</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f0f2f5; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; }}
        .card {{ background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); max-width: 500px; text-align: center; }}
        .icon {{ font-size: 54px; margin-bottom: 16px; }}
        h1 {{ color: #107c41; margin: 0 0 10px 0; font-size: 24px; }}
        p {{ color: #555; line-height: 1.5; margin: 8px 0; }}
        .btn {{ display: inline-block; margin-top: 20px; background: #107c41; color: white; padding: 12px 24px; border-radius: 6px; text-decoration: none; font-weight: bold; }}
    </style>
</head>
<body>
    <div class='card'>
        <div class='icon'>🎉</div>
        <h1>SRS Approved Successfully!</h1>
        <p>The Software Requirements Specification for <strong>{project}</strong> (Meeting {meetingId}) has been finalized and promoted to <code>srs-approved</code>.</p>
        <p>Official notification emails have been dispatched with the attached Word document.</p>
        <a class='btn' href='{finalDocxSasUrl}'>📥 Download Approved .docx</a>
    </div>
</body>
</html>";
                await htmlResponse.WriteStringAsync(successHtml);
                return htmlResponse;
            }
            else
            {
                var jsonResp = req.CreateResponse(HttpStatusCode.OK);
                await jsonResp.WriteAsJsonAsync(new
                {
                    success = true,
                    message = "SRS document approved successfully.",
                    project = project,
                    approvedBlob = approvedDocxName,
                    downloadUrl = finalDocxSasUrl
                });
                return jsonResp;
            }
        }

        private static string ExtractMeetingId(string blobName)
        {
            string name = Path.GetFileNameWithoutExtension(blobName);
            if (name.StartsWith("srs-draft-", StringComparison.OrdinalIgnoreCase))
            {
                string withoutPrefix = name.Substring("srs-draft-".Length);
                int lastHyphen = withoutPrefix.LastIndexOf('-');
                if (lastHyphen > 0)
                {
                    return withoutPrefix.Substring(0, lastHyphen);
                }
                return withoutPrefix;
            }
            return name;
        }

        private static string GenerateReadSasUri(BlobClient blobClient)
        {
            if (blobClient.CanGenerateSasUri)
            {
                BlobSasBuilder sasBuilder = new BlobSasBuilder
                {
                    BlobContainerName = blobClient.BlobContainerName,
                    BlobName = blobClient.Name,
                    Resource = "b",
                    ExpiresOn = DateTimeOffset.UtcNow.AddDays(7)
                };
                sasBuilder.SetPermissions(BlobSasPermissions.Read);
                return blobClient.GenerateSasUri(sasBuilder).ToString();
            }
            return blobClient.Uri.ToString();
        }

        private static string BuildEditorHtml(string draftId, string jsonContent)
        {
            // Escapes JSON for safe JavaScript injection
            string safeJson = jsonContent.Replace("\\", "\\\\").Replace("`", "\\`").Replace("$", "\\$");

            return $@"<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>SRS Review & Direct Document Editor</title>
    <style>
        :root {{
            --primary: #1F4E78;
            --primary-light: #2E75B6;
            --success: #107C41;
            --danger: #C53030;
            --bg: #F4F6F9;
            --card-bg: #FFFFFF;
            --text: #2D3748;
            --border: #CBD5E0;
            --subtle: #718096;
        }}
        * {{ box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg);
            color: var(--text);
            margin: 0;
            padding: 24px;
            line-height: 1.5;
        }}
        .container {{
            max-width: 1050px;
            margin: 0 auto;
            padding-bottom: 90px;
        }}
        .header {{
            background: linear-gradient(135deg, #1F4E78, #2E75B6);
            color: white;
            padding: 24px;
            border-radius: 12px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.12);
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .header p {{ margin: 6px 0 0 0; opacity: 0.9; font-size: 14px; }}
        .meta-bar {{
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            background: white;
            padding: 16px 20px;
            border-radius: 8px;
            margin-bottom: 24px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.06);
            align-items: center;
            border: 1px solid var(--border);
        }}
        .meta-item {{ display: flex; flex-direction: column; }}
        .meta-label {{ font-size: 11px; text-transform: uppercase; font-weight: bold; color: var(--subtle); margin-bottom: 4px; }}
        .meta-val {{ font-size: 13px; font-weight: 600; color: #2D3748; }}
        select, input[type=""text""], textarea {{
            padding: 9px 12px;
            border: 1px solid var(--border);
            border-radius: 6px;
            font-size: 14px;
            font-family: inherit;
            color: var(--text);
        }}
        select:focus, input[type=""text""]:focus, textarea:focus {{
            outline: none;
            border-color: var(--primary-light);
            box-shadow: 0 0 0 3px rgba(46,117,182,0.15);
        }}
        .card {{
            background: var(--card-bg);
            border-radius: 10px;
            padding: 24px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.05);
            margin-bottom: 24px;
            border: 1px solid var(--border);
        }}
        .card h2 {{
            margin-top: 0;
            font-size: 18px;
            color: var(--primary);
            border-bottom: 2px solid #EDF2F7;
            padding-bottom: 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        textarea {{
            width: 100%;
            min-height: 100px;
            resize: vertical;
        }}
        .req-card {{
            background: #F8FAFC;
            border: 1px solid #CBD5E1;
            border-left: 5px solid var(--primary);
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        }}
        .req-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 12px;
            margin-bottom: 14px;
            flex-wrap: wrap;
        }}
        .req-badge {{
            background: #1F4E78;
            color: white;
            font-size: 12px;
            font-weight: bold;
            padding: 4px 10px;
            border-radius: 4px;
            letter-spacing: 0.5px;
        }}
        .field-group {{
            margin-bottom: 14px;
        }}
        .field-group label {{
            display: block;
            font-size: 12px;
            font-weight: 700;
            color: #4A5568;
            margin-bottom: 4px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .impact-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 12px;
            margin-top: 10px;
            background: #EDF2F7;
            padding: 14px;
            border-radius: 6px;
        }}
        .actions-bar {{
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            background: white;
            padding: 16px 30px;
            box-shadow: 0 -4px 20px rgba(0,0,0,0.12);
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-top: 1px solid #CBD5E0;
            z-index: 100;
        }}
        .btn {{
            padding: 11px 22px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            border: none;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }}
        .btn-draft {{ background: #E2E8F0; color: #2D3748; margin-right: 12px; }}
        .btn-draft:hover {{ background: #CBD5E0; }}
        .btn-approve {{ background: var(--success); color: white; }}
        .btn-approve:hover {{ background: #0E6B37; }}
        .btn-add {{ background: #EDF2F7; color: var(--primary); border: 1px dashed var(--primary); width: 100%; justify-content: center; padding: 12px; }}
        .btn-add:hover {{ background: #E2E8F0; }}
        .btn-del {{ background: #FED7D7; color: #9B2C2C; padding: 6px 12px; font-size: 12px; }}
        .btn-del:hover {{ background: #FEB2B2; }}
        .toast {{
            display: none;
            position: fixed;
            top: 24px;
            right: 24px;
            background: #2D3748;
            color: white;
            padding: 14px 24px;
            border-radius: 8px;
            box-shadow: 0 4px 14px rgba(0,0,0,0.2);
            z-index: 1000;
            font-size: 14px;
            font-weight: 500;
        }}
    </style>
</head>
<body>
    <div class=""container"">
        <div class=""header"">
            <div>
                <h1>SRS Direct Document Editor</h1>
                <p>Modify requirements, acceptance criteria, database mappings, or open questions directly.</p>
            </div>
            <div id=""badgeProject"" style=""background: rgba(255,255,255,0.25); padding: 8px 16px; border-radius: 6px; font-weight: bold; font-size: 15px;"">RSCS</div>
        </div>

        <div class=""meta-bar"">
            <div class=""meta-item"" style=""min-width: 200px;"">
                <span class=""meta-label"">Draft File</span>
                <span class=""meta-val"" style=""word-break: break-all;"">{draftId}</span>
            </div>
            <div class=""meta-item"">
                <span class=""meta-label"">Target Project</span>
                <select id=""projectSelect"">
                    <option value=""RSCS"">RSCS</option>
                    <option value=""CBRE"">CBRE</option>
                    <option value=""AstuteDFM"">AstuteDFM</option>
                </select>
            </div>
            <div class=""meta-item"" style=""flex-grow: 1;"">
                <span class=""meta-label"">Meeting Title</span>
                <input type=""text"" id=""meetingTitleInput"" placeholder=""Meeting Title / Topic"" />
            </div>
            <div class=""meta-item"" style=""min-width: 160px;"">
                <span class=""meta-label"">Reviewer Name</span>
                <input type=""text"" id=""reviewerInput"" value=""Lead Reviewer"" placeholder=""Your name"" />
            </div>
        </div>

        <!-- Section 1: Executive Summary -->
        <div class=""card"">
            <h2>1. Executive Summary & Overview</h2>
            <p style=""font-size: 13px; color: var(--subtle); margin-top: -6px;"">High-level summary of the business problem, discussion context, and technical objectives.</p>
            <textarea id=""execSummary"" style=""min-height: 120px;"" placeholder=""Enter executive summary...""></textarea>
        </div>

        <!-- Section 2: Functional Requirements -->
        <div class=""card"">
            <h2>
                <span>2. Functional Requirements & System Impacts</span>
                <span id=""reqCountBadge"" style=""font-size: 13px; background: #E2E8F0; padding: 4px 10px; border-radius: 12px; color: #4A5568;"">0 Requirements</span>
            </h2>
            <div id=""requirementsContainer""></div>
            <button class=""btn btn-add"" onclick=""addRequirement()"">➕ Add New Requirement Card</button>
        </div>

        <!-- Section 3: Open Questions -->
        <div class=""card"">
            <h2>3. Open Questions & Clarifications Needed</h2>
            <p style=""font-size: 13px; color: var(--subtle); margin-top: -6px;"">Enter one open question or decision needed per line. These will be highlighted in yellow callout boxes in the Word document.</p>
            <textarea id=""openQuestions"" style=""min-height: 110px;"" placeholder=""Enter questions (one per line)...""></textarea>
        </div>

        <!-- Section 4: Technical Assumptions -->
        <div class=""card"">
            <h2>4. Technical Assumptions</h2>
            <p style=""font-size: 13px; color: var(--subtle); margin-top: -6px;"">Enter one technical assumption made during the discussion per line.</p>
            <textarea id=""assumptions"" style=""min-height: 110px;"" placeholder=""Enter assumptions (one per line)...""></textarea>
        </div>

        <!-- Sticky Actions Bar -->
        <div class=""actions-bar"">
            <span style=""color: #718096; font-size: 13px;"">💾 Changes update the Word (<code>.docx</code>) document and Azure storage immediately.</span>
            <div>
                <button class=""btn btn-draft"" onclick=""saveDraft()"">💾 Save as Draft</button>
                <button class=""btn btn-approve"" onclick=""saveAndApprove()"">✅ Save & Approve SRS</button>
            </div>
        </div>
    </div>

    <div id=""toast"" class=""toast""></div>

    <script>
        let currentPayload = {{}};
        try {{
            currentPayload = JSON.parse(`{safeJson}`);
        }} catch(e) {{
            currentPayload = {{}};
        }}
        const draftId = ""{draftId}"";

        function init() {{
            // Project
            const proj = currentPayload.project || ""RSCS"";
            document.getElementById('projectSelect').value = proj;
            document.getElementById('badgeProject').innerText = proj;

            document.getElementById('projectSelect').addEventListener('change', function() {{
                document.getElementById('badgeProject').innerText = this.value;
            }});

            // Meeting Title
            document.getElementById('meetingTitleInput').value = currentPayload.meetingTitle || """";

            // Executive Summary
            document.getElementById('execSummary').value = currentPayload.executiveSummary || currentPayload.overview || """";

            // Open Questions
            const oq = currentPayload.openQuestions || [];
            document.getElementById('openQuestions').value = Array.isArray(oq) 
                ? oq.map(x => typeof x === 'object' && x !== null ? (x.question || JSON.stringify(x)) : String(x)).join('\n') 
                : (oq || '');

            // Assumptions
            const assump = currentPayload.assumptions || [];
            document.getElementById('assumptions').value = Array.isArray(assump) 
                ? assump.map(x => typeof x === 'object' && x !== null ? (x.assumption || JSON.stringify(x)) : String(x)).join('\n') 
                : (assump || '');

            // Requirements array fallback (support requirements or epics/userStories if present)
            if (!Array.isArray(currentPayload.requirements) || currentPayload.requirements.length === 0) {{
                if (Array.isArray(currentPayload.epics) && currentPayload.epics.length > 0) {{
                    currentPayload.requirements = [];
                    let rNum = 1;
                    currentPayload.epics.forEach((e) => {{
                        if (Array.isArray(e.userStories) && e.userStories.length > 0) {{
                            e.userStories.forEach((s) => {{
                                currentPayload.requirements.push({{
                                    id: `REQ-${{String(rNum++).padStart(3, '0')}}`,
                                    title: s.title || `${{e.title || 'Requirement'}} - Story`,
                                    desiredBehavior: s.description || e.description || '',
                                    classification: 'NEW',
                                    sourceMeeting: '',
                                    existingSystemComparison: '',
                                    acceptanceCriteria: s.acceptanceCriteria || [],
                                    impacts: {{ database: '', backend: '', frontend: '', mobile: '' }}
                                }});
                            }});
                        }} else {{
                            currentPayload.requirements.push({{
                                id: `REQ-${{String(rNum++).padStart(3, '0')}}`,
                                title: e.title || `Requirement ${{rNum}}`,
                                desiredBehavior: e.description || '',
                                classification: 'NEW',
                                sourceMeeting: '',
                                existingSystemComparison: '',
                                acceptanceCriteria: e.acceptanceCriteria || [],
                                impacts: {{ database: '', backend: '', frontend: '', mobile: '' }}
                            }});
                        }}
                    }});
                }} else {{
                    currentPayload.requirements = [];
                }}
            }}

            renderRequirements();
        }}

        function renderRequirements() {{
            const container = document.getElementById('requirementsContainer');
            container.innerHTML = '';
            const reqs = currentPayload.requirements || [];
            document.getElementById('reqCountBadge').innerText = `${{reqs.length}} Requirement${{reqs.length === 1 ? '' : 's'}}`;

            reqs.forEach((req, idx) => {{
                const acText = Array.isArray(req.acceptanceCriteria) ? req.acceptanceCriteria.join('\n') : (req.acceptanceCriteria || '');
                const impacts = req.impacts || {{}};

                const card = document.createElement('div');
                card.className = 'req-card';
                card.innerHTML = `
                    <div class=""req-header"">
                        <div style=""display: flex; align-items: center; gap: 8px; flex-grow: 1;"">
                            <input type=""text"" class=""req-id-val"" style=""width: 100px; font-weight: bold; background: #1F4E78; color: white; border-radius: 4px; padding: 6px 10px; text-align: center;"" value=""${{escapeHtml(req.id || `REQ-${{String(idx + 1).padStart(3, '0')}}`)}}\"" placeholder=""REQ-001"" title=""Editable Requirement ID"" />
                            <input type=""text"" class=""req-title-val"" style=""flex-grow: 1; font-weight: bold;"" value=""${{escapeHtml(req.title || '')}}\"" placeholder=""Requirement Title"" />
                        </div>
                        <div style=""display: flex; align-items: center; gap: 8px;"">
                            <select class=""req-class-val"">
                                <option value=""NEW"" ${{req.classification === 'NEW' ? 'selected' : ''}}>NEW</option>
                                <option value=""CHANGE"" ${{req.classification === 'CHANGE' ? 'selected' : ''}}>CHANGE</option>
                                <option value=""CONFLICT"" ${{req.classification === 'CONFLICT' ? 'selected' : ''}}>CONFLICT</option>
                            </select>
                            <button class=""btn btn-del"" onclick=""deleteRequirement(${{idx}})"">🗑️ Delete</button>
                        </div>
                    </div>

                    <div class=""field-group"">
                        <label>Desired Behavior / Description:</label>
                        <textarea class=""req-behavior-val"" style=""min-height: 80px;"" placeholder=""Detailed functional requirement and expected system behavior..."">${{escapeHtml(req.desiredBehavior || '')}}</textarea>
                    </div>

                    <div class=""field-group"">
                        <label>Transcript Quote / Source Meeting Evidence:</label>
                        <input type=""text"" class=""req-source-val"" style=""width: 100%;"" value=""${{escapeHtml(req.sourceMeeting || '')}}\"" placeholder=""Quote or discussion timestamp reference"" />
                    </div>

                    <div class=""field-group"">
                        <label>Existing System Impact & Database Comparison:</label>
                        <textarea class=""req-comparison-val"" style=""min-height: 60px;"" placeholder=""How does this impact existing DB schema, SPs, or API architecture?"">${{escapeHtml(req.existingSystemComparison || '')}}</textarea>
                    </div>

                    <div class=""field-group"">
                        <label>Acceptance Criteria (one per line):</label>
                        <textarea class=""req-ac-val"" style=""min-height: 75px;"" placeholder=""User can view invoice details...&#10;System validates tax ID before saving..."">${{escapeHtml(acText)}}</textarea>
                    </div>

                    <div class=""field-group"">
                        <label>Component Impacts Breakdown:</label>
                        <div class=""impact-grid"">
                            <div>
                                <label style=""font-size: 11px;"">Database (Tables, Columns, SPs):</label>
                                <textarea class=""req-db-val"" style=""min-height: 65px; font-size: 13px;"" placeholder=""Impacted tables, columns, SPs..."">${{escapeHtml(impacts.database || '')}}</textarea>
                            </div>
                            <div>
                                <label style=""font-size: 11px;"">Backend (APIs, Controllers, Services):</label>
                                <textarea class=""req-backend-val"" style=""min-height: 65px; font-size: 13px;"" placeholder=""Impacted controllers, services, APIs..."">${{escapeHtml(impacts.backend || '')}}</textarea>
                            </div>
                            <div>
                                <label style=""font-size: 11px;"">Frontend (Pages, UI Components):</label>
                                <textarea class=""req-frontend-val"" style=""min-height: 65px; font-size: 13px;"" placeholder=""Impacted pages, forms, UI components..."">${{escapeHtml(impacts.frontend || '')}}</textarea>
                            </div>
                            <div>
                                <label style=""font-size: 11px;"">Mobile API / Compatibility:</label>
                                <textarea class=""req-mobile-val"" style=""min-height: 65px; font-size: 13px;"" placeholder=""Mobile API compatibility notes..."">${{escapeHtml(impacts.mobile || '')}}</textarea>
                            </div>
                        </div>
                    </div>
                `;
                container.appendChild(card);
            }});
        }}

        function addRequirement() {{
            collectPayload(); // Preserve any in-flight edits across existing cards
            const newIndex = (currentPayload.requirements || []).length + 1;
            const newReq = {{
                id: `REQ-${{String(newIndex).padStart(3, '0')}}`,
                title: ""New Functional Requirement"",
                desiredBehavior: """",
                sourceMeeting: """",
                classification: ""NEW"",
                existingSystemComparison: """",
                acceptanceCriteria: [],
                impacts: {{ database: """", backend: """", frontend: """", mobile: """" }}
            }};
            currentPayload.requirements.push(newReq);
            renderRequirements();
            showToast(""Added new requirement card."");
        }}

        function deleteRequirement(idx) {{
            collectPayload(); // Preserve current edits
            const reqId = currentPayload.requirements[idx]?.id || `#${{idx + 1}}`;
            if (!confirm(`Are you sure you want to delete requirement ${{reqId}}?`)) return;
            currentPayload.requirements.splice(idx, 1);
            renderRequirements();
            showToast(""Requirement deleted."");
        }}

        function collectPayload() {{
            currentPayload.project = document.getElementById('projectSelect').value;
            currentPayload.meetingTitle = document.getElementById('meetingTitleInput').value;
            currentPayload.executiveSummary = document.getElementById('execSummary').value;

            // Open questions
            currentPayload.openQuestions = document.getElementById('openQuestions').value
                .split('\n')
                .map(x => x.trim())
                .filter(x => x.length > 0);

            // Assumptions
            currentPayload.assumptions = document.getElementById('assumptions').value
                .split('\n')
                .map(x => x.trim())
                .filter(x => x.length > 0);

            // Direct extraction from DOM for all requirement cards
            const cards = document.querySelectorAll('.req-card');
            const gatheredReqs = [];
            cards.forEach((card, i) => {{
                const idInput = card.querySelector('.req-id-val');
                const titleInput = card.querySelector('.req-title-val');
                const classSelect = card.querySelector('.req-class-val');
                const behaviorInput = card.querySelector('.req-behavior-val');
                const sourceInput = card.querySelector('.req-source-val');
                const compInput = card.querySelector('.req-comparison-val');
                const acInput = card.querySelector('.req-ac-val');
                const dbInput = card.querySelector('.req-db-val');
                const backInput = card.querySelector('.req-backend-val');
                const frontInput = card.querySelector('.req-frontend-val');
                const mobInput = card.querySelector('.req-mobile-val');

                gatheredReqs.push({{
                    id: (idInput && idInput.value.trim()) ? idInput.value.trim() : `REQ-${{String(i+1).padStart(3, '0')}}`,
                    title: titleInput ? titleInput.value : '',
                    classification: classSelect ? classSelect.value : 'NEW',
                    desiredBehavior: behaviorInput ? behaviorInput.value : '',
                    sourceMeeting: sourceInput ? sourceInput.value : '',
                    existingSystemComparison: compInput ? compInput.value : '',
                    acceptanceCriteria: acInput ? acInput.value.split('\n').map(x => x.trim()).filter(x => x.length > 0) : [],
                    impacts: {{
                        database: dbInput ? dbInput.value : '',
                        backend: backInput ? backInput.value : '',
                        frontend: frontInput ? frontInput.value : '',
                        mobile: mobInput ? mobInput.value : ''
                    }}
                }});
            }});
            currentPayload.requirements = gatheredReqs;

            return currentPayload;
        }}

        async function saveDraft() {{
            showToast(""Saving draft updates..."");
            const payload = collectPayload();
            const approver = document.getElementById('reviewerInput').value || ""Lead Reviewer"";

            try {{
                const res = await fetch('/api/srs/save-draft', {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify({{ draftId: draftId, approverName: approver, srsPayload: payload }})
                }});
                const data = await res.json();
                if (res.ok) {{
                    showToast(""✅ Draft & Word (.docx) successfully updated!"");
                }} else {{
                    showToast(""❌ Error saving draft: "" + (data.error || ""Server error""));
                }}
            }} catch(err) {{
                showToast(""❌ Network error: "" + err.message);
            }}
        }}

        async function saveAndApprove() {{
            if (!confirm(""Are you sure you want to officially approve and finalize this SRS document?"")) return;
            showToast(""Promoting to srs-approved and dispatching emails..."");
            const payload = collectPayload();
            const approver = document.getElementById('reviewerInput').value || ""Lead Reviewer"";

            try {{
                const res = await fetch('/api/srs/save-approve', {{
                    method: 'POST',
                    headers: {{ 'Content-Type': 'application/json' }},
                    body: JSON.stringify({{ draftId: draftId, project: payload.project, approverName: approver, srsPayload: payload }})
                }});
                const data = await res.json();
                if (res.ok) {{
                    alert(""🎉 SRS Approved Successfully!\n\nThe finalized Word (.docx) document has been promoted to 'srs-approved' and confirmation emails have been dispatched."");
                    window.location.reload();
                }} else {{
                    alert(""Error approving document: "" + (data.error || ""Unknown error""));
                }}
            }} catch(err) {{
                alert(""Network error: "" + err.message);
            }}
        }}

        function escapeHtml(str) {{
            if (!str) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/""/g, '&quot;')
                .replace(/'/g, '&#039;');
        }}

        function showToast(msg) {{
            const t = document.getElementById('toast');
            t.innerText = msg;
            t.style.display = 'block';
            setTimeout(() => {{ t.style.display = 'none'; }}, 3500);
        }}

        init();
    </script>
</body>
</html>";
        }
    }
}
