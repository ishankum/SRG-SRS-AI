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
                await docxClient.UploadAsync(new BinaryData(docxBytes), overwrite: true);

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
            await appDocxClient.UploadAsync(new BinaryData(docxBytes), overwrite: true);

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
            // Format: srs-draft-{meetingId}-{timestamp}.json
            string name = Path.GetFileNameWithoutExtension(blobName);
            var parts = name.Split('-');
            if (parts.Length >= 3 && parts[0] == "srs" && parts[1] == "draft")
            {
                return parts[2];
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
            --success: #107C41;
            --bg: #F4F6F9;
            --card-bg: #FFFFFF;
            --text: #2D3748;
            --border: #E2E8F0;
        }}
        * {{ box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg);
            color: var(--text);
            margin: 0;
            padding: 24px;
        }}
        .container {{
            max-width: 1000px;
            margin: 0 auto;
        }}
        .header {{
            background: linear-gradient(135deg, #1F4E78, #2E75B6);
            color: white;
            padding: 24px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            margin-bottom: 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .header p {{ margin: 6px 0 0 0; opacity: 0.9; font-size: 14px; }}
        .meta-bar {{
            display: flex;
            gap: 16px;
            background: white;
            padding: 16px 20px;
            border-radius: 8px;
            margin-bottom: 24px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            align-items: center;
        }}
        .meta-item {{ display: flex; flex-direction: column; }}
        .meta-label {{ font-size: 11px; text-transform: uppercase; font-weight: bold; color: #718096; }}
        .meta-val {{ font-size: 14px; font-weight: 600; color: #2D3748; }}
        select, input[type=""text""] {{
            padding: 8px 12px;
            border: 1px solid var(--border);
            border-radius: 6px;
            font-size: 14px;
        }}
        .card {{
            background: var(--card-bg);
            border-radius: 10px;
            padding: 24px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.06);
            margin-bottom: 20px;
            border: 1px solid var(--border);
        }}
        .card h2 {{
            margin-top: 0;
            font-size: 18px;
            color: var(--primary);
            border-bottom: 2px solid #EDF2F7;
            padding-bottom: 10px;
        }}
        textarea {{
            width: 100%;
            min-height: 120px;
            padding: 12px;
            border: 1px solid var(--border);
            border-radius: 6px;
            font-family: inherit;
            font-size: 14px;
            line-height: 1.5;
            resize: vertical;
        }}
        .epic-item {{
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 16px;
        }}
        .actions-bar {{
            position: sticky;
            bottom: 20px;
            background: white;
            padding: 16px 24px;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.15);
            display: flex;
            justify-content: space-between;
            align-items: center;
            border: 1px solid #CBD5E0;
        }}
        .btn {{
            padding: 12px 24px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            border: none;
            transition: all 0.2s;
        }}
        .btn-draft {{
            background: #E2E8F0;
            color: #2D3748;
            margin-right: 12px;
        }}
        .btn-draft:hover {{ background: #CBD5E0; }}
        .btn-approve {{
            background: var(--success);
            color: white;
        }}
        .btn-approve:hover {{ background: #0E6B37; }}
        .toast {{
            display: none;
            position: fixed;
            top: 24px;
            right: 24px;
            background: #2D3748;
            color: white;
            padding: 14px 24px;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 1000;
        }}
    </style>
</head>
<body>
    <div class=""container"">
        <div class=""header"">
            <div>
                <h1>SRS Direct Review & Editor</h1>
                <p>Edit requirements, user stories, or database mappings directly before saving or approving.</p>
            </div>
            <div id=""badgeProject"" style=""background: rgba(255,255,255,0.2); padding: 8px 16px; border-radius: 6px; font-weight: bold;"">RSCS</div>
        </div>

        <div class=""meta-bar"">
            <div class=""meta-item"">
                <span class=""meta-label"">Draft File</span>
                <span class=""meta-val"">{draftId}</span>
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
                <span class=""meta-label"">Reviewer Name</span>
                <input type=""text"" id=""reviewerInput"" value=""Lead Reviewer"" placeholder=""Your name"" />
            </div>
        </div>

        <!-- Section 1: Executive Summary -->
        <div class=""card"">
            <h2>1. Executive Summary & Overview</h2>
            <textarea id=""execSummary""></textarea>
        </div>

        <!-- Section 2: Epics and User Stories -->
        <div class=""card"">
            <h2>2. Epics & User Stories</h2>
            <div id=""epicsContainer""></div>
        </div>

        <!-- Section 3: Database & Schema Mappings -->
        <div class=""card"">
            <h2>3. Database Schema Mappings</h2>
            <textarea id=""dbMappings"" style=""min-height: 140px;""></textarea>
        </div>

        <!-- Section 4: Non-Functional & Security Constraints -->
        <div class=""card"">
            <h2>4. Non-Functional & Compliance Constraints</h2>
            <textarea id=""nfrSummary"" style=""min-height: 100px;""></textarea>
        </div>

        <!-- Sticky Actions Bar -->
        <div class=""actions-bar"">
            <span style=""color: #718096; font-size: 13px;"">Changes made here update the Word (.docx) document automatically.</span>
            <div>
                <button class=""btn btn-draft"" onclick=""saveDraft()"">💾 Save as Draft</button>
                <button class=""btn btn-approve"" onclick=""saveAndApprove()"">✅ Save & Approve SRS</button>
            </div>
        </div>
    </div>

    <div id=""toast"" class=""toast""></div>

    <script>
        let currentPayload = JSON.parse(`{safeJson}`);
        const draftId = ""{draftId}"";

        function init() {{
            document.getElementById('execSummary').value = currentPayload.executiveSummary || '';
            document.getElementById('dbMappings').value = JSON.stringify(currentPayload.databaseMapping || [], null, 2);
            document.getElementById('nfrSummary').value = (currentPayload.nonFunctionalRequirements || []).join('\n');

            if (currentPayload.project) {{
                document.getElementById('projectSelect').value = currentPayload.project;
                document.getElementById('badgeProject').innerText = currentPayload.project;
            }}

            renderEpics();
        }}

        function renderEpics() {{
            const container = document.getElementById('epicsContainer');
            container.innerHTML = '';
            const epics = currentPayload.epics || [];

            epics.forEach((epic, idx) => {{
                const div = document.createElement('div');
                div.className = 'epic-item';
                div.innerHTML = `
                    <div style=""display: flex; justify-content: space-between; margin-bottom: 8px;"">
                        <strong>Epic ${{idx + 1}}: <input type=""text"" style=""width: 80%;"" value=""${{epic.title || ''}}"" onchange=""updateEpicTitle(${{idx}}, this.value)"" /></strong>
                    </div>
                    <textarea style=""min-height: 70px; margin-bottom: 8px;"" onchange=""updateEpicDesc(${{idx}}, this.value)"">${{epic.description || ''}}</textarea>
                    <label style=""font-size: 12px; font-weight: bold; color: #4A5568;"">Acceptance Criteria (one per line):</label>
                    <textarea style=""min-height: 80px;"" onchange=""updateEpicAc(${{idx}}, this.value)"">${{(epic.acceptanceCriteria || []).join('\n')}}</textarea>
                `;
                container.appendChild(div);
            }});
        }}

        function updateEpicTitle(idx, val) {{ currentPayload.epics[idx].title = val; }}
        function updateEpicDesc(idx, val) {{ currentPayload.epics[idx].description = val; }}
        function updateEpicAc(idx, val) {{ currentPayload.epics[idx].acceptanceCriteria = val.split('\n').filter(x => x.trim().length > 0); }}

        function collectPayload() {{
            currentPayload.executiveSummary = document.getElementById('execSummary').value;
            currentPayload.project = document.getElementById('projectSelect').value;
            try {{
                currentPayload.databaseMapping = JSON.parse(document.getElementById('dbMappings').value);
            }} catch(e) {{}}
            currentPayload.nonFunctionalRequirements = document.getElementById('nfrSummary').value.split('\n').filter(x => x.trim().length > 0);
            return currentPayload;
        }}

        async function saveDraft() {{
            showToast(""Saving draft updates..."");
            const payload = collectPayload();
            const approver = document.getElementById('reviewerInput').value;

            const res = await fetch('/api/srs/save-draft', {{
                method: 'POST',
                headers: {{ 'Content-Type': 'application/json' }},
                body: JSON.stringify({{ draftId: draftId, approverName: approver, srsPayload: payload }})
            }});
            if (res.ok) {{
                showToast(""✅ Draft successfully saved!"");
            }} else {{
                showToast(""❌ Error saving draft."");
            }}
        }}

        async function saveAndApprove() {{
            if (!confirm(""Are you sure you want to officially approve and finalize this SRS document?"")) return;
            showToast(""Promoting to srs-approved and dispatching emails..."");
            const payload = collectPayload();
            const approver = document.getElementById('reviewerInput').value;

            const res = await fetch('/api/srs/save-approve', {{
                method: 'POST',
                headers: {{ 'Content-Type': 'application/json' }},
                body: JSON.stringify({{ draftId: draftId, project: payload.project, approverName: approver, srsPayload: payload }})
            }});
            const data = await res.json();
            if (res.ok) {{
                alert(""🎉 SRS Approved Successfully! Confirmation emails have been dispatched."");
                window.location.reload();
            }} else {{
                alert(""Error approving document: "" + (data.error || ""Unknown error""));
            }}
        }}

        function showToast(msg) {{
            const t = document.getElementById('toast');
            t.innerText = msg;
            t.style.display = 'block';
            setTimeout(() => {{ t.style.display = 'none'; }}, 3000);
        }}

        init();
    </script>
</body>
</html>";
        }
    }
}
