using Azure.Storage.Blobs;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace SrsAi.Functions
{
    public class SrsAgentWorker
    {
        private readonly ILogger _logger;
        private readonly IConfiguration _configuration;
        private readonly BlobServiceClient _blobServiceClient;
        private readonly ProjectContextFetcher _contextFetcher;
        private readonly EmailNotificationService _emailService;
        private readonly TeamsNotificationService _teamsService;
        private readonly ProjectClassifier _classifier;

        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(5)
        };

        public SrsAgentWorker(ILoggerFactory loggerFactory, IConfiguration configuration)
        {
            _logger = loggerFactory.CreateLogger<SrsAgentWorker>();
            _configuration = configuration;

            string connectionString = _configuration["AzureWebJobsStorage"];
            _blobServiceClient = new BlobServiceClient(connectionString);
            _contextFetcher = new ProjectContextFetcher(configuration, _logger);
            _emailService = new EmailNotificationService(configuration, _logger);
            _teamsService = new TeamsNotificationService(configuration, _logger);
            _classifier = new ProjectClassifier(configuration, _logger);
        }

        [Function("SrsAgentWorker")]
        public async Task Run([QueueTrigger("srs-generation", Connection = "AzureWebJobsStorage")] string queueMessage)
        {
            _logger.LogInformation($"SRS Agent started for message: {queueMessage}");

            try
            {
                using JsonDocument queueDoc = JsonDocument.Parse(queueMessage);
                string meetingId = queueDoc.RootElement.GetProperty("MeetingId").GetString() ?? "UNKNOWN";
                string transcriptBlobName = queueDoc.RootElement.GetProperty("TranscriptBlobName").GetString() ?? "";

                // 1. Read Transcript from Blob
                var transcriptContainer = _blobServiceClient.GetBlobContainerClient("transcripts");
                var transcriptBlob = transcriptContainer.GetBlobClient(transcriptBlobName);

                var downloadResult = await transcriptBlob.DownloadContentAsync();
                string transcriptData = downloadResult.Value.Content.ToString();

                // Dynamic Project Classification from Transcript Content
                var classification = _classifier.Classify(transcriptData);
                string projectName = classification.ProjectName;
                _logger.LogInformation($"[SrsAgentWorker] Classified project as '{projectName}' (Source: {classification.MatchedSource}, Confidence: {classification.Confidence:P0}). Fetching Snapshots...");

                // Read snapshot files (with candidate extensions for db-schema.json / db-schema.txt)
                string architectureDoc = await _contextFetcher.GetProjectSnapshotAsync(projectName, "architecture.md", "architecture.txt");
                string dbSchemaSnapshot = await _contextFetcher.GetProjectSnapshotAsync(projectName, "db-schema.json", "db-schema.txt", "db_schema.json");

                string combinedContext = $@"
                === ARCHITECTURE RULES ({projectName}) ===
                {architectureDoc}

                === DATABASE SCHEMA SNAPSHOT ({projectName}) ===
                {dbSchemaSnapshot}
                ";

                // 2. Invoke AI Model
                string generatedSrsJson = await GenerateSrsFromAiAsync(transcriptData, combinedContext, projectName);

                // 3. Generate Word (.docx) SRS Document and JSON Draft
                byte[] docxBytes = SrsWordDocumentBuilder.BuildWordDocument(generatedSrsJson, meetingId);

                var draftContainer = _blobServiceClient.GetBlobContainerClient("srs-drafts");
                await draftContainer.CreateIfNotExistsAsync();

                string timestampStr = DateTime.UtcNow.ToString("yyyyMMddHHmmss");

                // Save Editable Word Document (.docx)
                string docxBlobName = $"srs-draft-{meetingId}-{timestampStr}.docx";
                var docxBlobClient = draftContainer.GetBlobClient(docxBlobName);

                var docxOptions = new Azure.Storage.Blobs.Models.BlobUploadOptions
                {
                    HttpHeaders = new Azure.Storage.Blobs.Models.BlobHttpHeaders
                    {
                        ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    }
                };

                await docxBlobClient.UploadAsync(BinaryData.FromBytes(docxBytes), docxOptions);
                _logger.LogInformation($"Draft SRS Word Document (.docx) successfully saved to Blob: {docxBlobName}");

                // Save Structured JSON version as well
                string jsonBlobName = $"srs-draft-{meetingId}-{timestampStr}.json";
                var jsonBlobClient = draftContainer.GetBlobClient(jsonBlobName);
                await jsonBlobClient.UploadAsync(BinaryData.FromString(generatedSrsJson), overwrite: true);
                _logger.LogInformation($"Draft SRS JSON successfully saved to Blob: {jsonBlobName}");

                // 4. Generate 24-hr SAS URL for Word document (.docx) download
                string docxSasUrl = string.Empty;
                if (docxBlobClient.CanGenerateSasUri)
                {
                    var sasBuilder = new Azure.Storage.Sas.BlobSasBuilder
                    {
                        BlobContainerName = "srs-drafts",
                        BlobName = docxBlobName,
                        Resource = "b",
                        ExpiresOn = DateTimeOffset.UtcNow.AddHours(24)
                    };
                    sasBuilder.SetPermissions(Azure.Storage.Sas.BlobSasPermissions.Read);
                    docxSasUrl = docxBlobClient.GenerateSasUri(sasBuilder).ToString();
                }
                else
                {
                    docxSasUrl = docxBlobClient.Uri.ToString();
                }

                // 5. Parse summary and metrics for Teams Adaptive Card
                string summary = "SRS proposal generated based on meeting transcript.";
                int epicsCount = 0;
                int storiesCount = 0;
                int tablesCount = 0;
                string meetingTitle = $"Meeting {meetingId}";

                try
                {
                    using var srsDoc = JsonDocument.Parse(generatedSrsJson);
                    var root = srsDoc.RootElement;
                    if (root.TryGetProperty("executiveSummary", out var es)) summary = es.GetString() ?? summary;
                    if (root.TryGetProperty("meetingTitle", out var mt)) meetingTitle = mt.GetString() ?? meetingTitle;
                    if (root.TryGetProperty("epics", out var epics) && epics.ValueKind == JsonValueKind.Array)
                    {
                        epicsCount = epics.GetArrayLength();
                        foreach (var epic in epics.EnumerateArray())
                        {
                            if (epic.TryGetProperty("userStories", out var stories) && stories.ValueKind == JsonValueKind.Array)
                            {
                                storiesCount += stories.GetArrayLength();
                            }
                        }
                    }
                    else if (root.TryGetProperty("requirements", out var reqs) && reqs.ValueKind == JsonValueKind.Array)
                    {
                        storiesCount = reqs.GetArrayLength();
                    }

                    if (root.TryGetProperty("databaseMapping", out var dbm) && dbm.ValueKind == JsonValueKind.Array)
                    {
                        tablesCount = dbm.GetArrayLength();
                    }
                }
                catch { }

                // 6. Post Interactive Adaptive Card to Microsoft Teams
                await _teamsService.SendSrsDraftCardAsync(
                    projectName: projectName,
                    meetingId: meetingId,
                    meetingTitle: meetingTitle,
                    executiveSummary: summary,
                    epicsCount: epicsCount,
                    storiesCount: storiesCount,
                    tablesCount: tablesCount,
                    docxSasUrl: docxSasUrl,
                    draftBlobName: jsonBlobName
                );

                // 7. Send email notification to configured reviewers with attached Word document
                await _emailService.SendSrsDraftEmailAsync(
                    meetingId: meetingId,
                    projectName: projectName,
                    jsonPayload: generatedSrsJson,
                    docxBytes: docxBytes,
                    docxFileName: docxBlobName
                );
            }
            catch (Exception ex)
            {
                _logger.LogError($"SRS Generation failed: {ex.Message}");
                throw;
            }
        }

        private async Task<string> GenerateSrsFromAiAsync(string transcriptData, string combinedContext, string projectName)
        {
            string baseUrl = _configuration["AiModelBaseUrl"];
            string apiKey = _configuration["AiModelApiKey"];

            string prompt = $@"
        Analyze the following meeting transcript to generate a Software Requirements Specification (SRS).
        You are also provided with the current project architecture and database schema.
        Compare the transcript requests with the existing system context. Flag any duplicates, changes, or conflicts. 
        Do not invent deadlines or API behaviors.
        
        === EXISTING SYSTEM CONTEXT ===
        
        ARCHITECTURE RULES AND DATABASE SCHEMA:
        {combinedContext}
        
        ===============================

        === MEETING TRANSCRIPT ===
        {transcriptData}
        ==========================

        Format required (Return ONLY valid JSON):
        {{
          ""project"": ""Extracted Project Name"",
          ""meetingTitle"": ""Meeting Title"",
          ""requirements"": [
            {{
              ""id"": ""REQ-001"",
              ""title"": ""Short Descriptive Title"",
              ""desiredBehavior"": ""Detailed requirement description..."",
              ""sourceMeeting"": ""Transcript quote or timestamp reference"",
              ""classification"": ""NEW | CHANGE | CONFLICT"",
              ""existingSystemComparison"": ""How does this impact existing DB schema, SPs, or API Architecture?"",
              ""acceptanceCriteria"": [""Criteria 1"", ""Criteria 2""],
              ""impacts"": {{
                ""backend"": ""Impacted controllers, services, and APIs"",
                ""database"": ""Impacted tables, columns, and stored procedures"",
                ""frontend"": ""UI pages, components, or buttons required"",
                ""mobile"": ""Mobile API impact or compatibility recommendation""
              }}
            }}
          ],
          ""openQuestions"": [""Question requiring clarification""],
          ""assumptions"": [""Technical assumption made""]
        }}
    ";

            var payload = new
            {
                contents = new[]
                {
                    new
                    {
                        parts = new[]
                        {
                            new { text = prompt }
                        }
                    }
                }
            };

            string requestUrl = $"{baseUrl}?key={apiKey}";

            int maxRetries = 3;
            HttpResponseMessage? response = null;
            string responseBody = string.Empty;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                var requestContent = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
                _logger.LogInformation($"Calling AI API URL (Attempt {attempt}/{maxRetries}): {baseUrl}");

                response = await _httpClient.PostAsync(requestUrl, requestContent);

                if (response.IsSuccessStatusCode)
                {
                    responseBody = await response.Content.ReadAsStringAsync();
                    break;
                }

                string errorDetails = await response.Content.ReadAsStringAsync();
                _logger.LogWarning($"AI API call attempt {attempt} failed with status code {(int)response.StatusCode} ({response.StatusCode}): {errorDetails}");

                if (attempt < maxRetries && (response.StatusCode == (System.Net.HttpStatusCode)503 || response.StatusCode == (System.Net.HttpStatusCode)429 || response.StatusCode == (System.Net.HttpStatusCode)500))
                {
                    int delayMs = attempt * 3000;
                    _logger.LogInformation($"Transient AI service error ({response.StatusCode}). Retrying in {delayMs}ms...");
                    await Task.Delay(delayMs);
                }
                else
                {
                    throw new HttpRequestException($"AI Model API call failed with status code {(int)response.StatusCode} ({response.StatusCode}). Error details: {errorDetails}");
                }
            }

            // Extract text from Gemini response
            using JsonDocument doc = JsonDocument.Parse(responseBody);
            var textElement = doc.RootElement
                .GetProperty("candidates")[0]
                .GetProperty("content")
                .GetProperty("parts")[0]
                .GetProperty("text");

            // Extract raw text from Gemini response
            string rawText = textElement.GetString()?.Trim() ?? "";
            return ExtractJsonPayload(rawText);
        }

        private static string ExtractJsonPayload(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "{}";

            int firstBrace = input.IndexOf('{');
            int lastBrace = input.LastIndexOf('}');

            if (firstBrace >= 0 && lastBrace > firstBrace)
            {
                return input.Substring(firstBrace, lastBrace - firstBrace + 1).Trim();
            }

            return input.Trim();
        }
    }
}