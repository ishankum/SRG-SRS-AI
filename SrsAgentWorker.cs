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

                _logger.LogInformation($"Transcript read successfully. Fetching Blob Snapshots for RSCS...");

                string projectName = "RSCS";

                // Read snapshot files (with candidate extensions for db-schema.json / db-schema.txt)
                string architectureDoc = await _contextFetcher.GetProjectSnapshotAsync(projectName, "architecture.md", "architecture.txt");
                string dbSchemaSnapshot = await _contextFetcher.GetProjectSnapshotAsync(projectName, "db-schema.json", "db-schema.txt", "db_schema.json");

                string combinedContext = $@"
                === ARCHITECTURE RULES ===
                {architectureDoc}

                === DATABASE SCHEMA SNAPSHOT ===
                {dbSchemaSnapshot}
                ";

                // 2. Invoke AI Model (with automatic retry for 503/429 transient errors)
                string generatedSrsJson = await GenerateSrsFromAiAsync(transcriptData, combinedContext);

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

                // 4. Send email notification to configured reviewers with attached Word document
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

        private async Task<string> GenerateSrsFromAiAsync(string transcriptData, string combinedContext)
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