using Azure.Storage.Blobs;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;

namespace SrsAi.Functions
{
    public class ContextSyncFunctions
    {
        private readonly ILogger _logger;
        private readonly IConfiguration _configuration;
        private readonly BlobServiceClient _blobServiceClient;
        private readonly DatabaseSchemaReader _schemaReader;
        private readonly GitRepositoryFetcher _gitFetcher;

        public ContextSyncFunctions(ILoggerFactory loggerFactory, IConfiguration configuration)
        {
            _logger = loggerFactory.CreateLogger<ContextSyncFunctions>();
            _configuration = configuration;

            string connectionString = _configuration["AzureWebJobsStorage"] ?? "";
            _blobServiceClient = new BlobServiceClient(connectionString);

            _schemaReader = new DatabaseSchemaReader(_logger);
            _gitFetcher = new GitRepositoryFetcher(_configuration, _logger);
        }

        [Function("SyncContextManual")]
        public async Task<HttpResponseData> RunHttpSync(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", "get", Route = "sync/context")] HttpRequestData req)
        {
            _logger.LogInformation("Received HTTP request to sync project database schema and architecture context.");

            // Extract project query parameter or default to RSCS
            string projectName = "RSCS";
            var query = System.Web.HttpUtility.ParseQueryString(req.Url.Query);
            if (!string.IsNullOrWhiteSpace(query["project"]))
            {
                projectName = query["project"]!.Trim().ToUpper();
            }

            var syncResult = await SyncProjectContextAsync(projectName);

            var response = req.CreateResponse(syncResult.Success ? HttpStatusCode.OK : HttpStatusCode.InternalServerError);
            await response.WriteAsJsonAsync(syncResult);
            return response;
        }

        [Function("SyncContextTimer")]
        public async Task RunTimerSync([TimerTrigger("0 0 */6 * * *")] TimerInfo timerInfo)
        {
            _logger.LogInformation($"Scheduled Timer Trigger executed context sync at {DateTime.UtcNow}.");

            string activeProjectsStr = _configuration["ActiveProjects"] ?? "RSCS";
            string[] activeProjects = activeProjectsStr.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string proj in activeProjects)
            {
                string projectName = proj.Trim().ToUpper();
                if (!string.IsNullOrWhiteSpace(projectName))
                {
                    _logger.LogInformation($"Starting scheduled context sync for project: {projectName}");
                    await SyncProjectContextAsync(projectName);
                }
            }
        }

        private async Task<SyncResult> SyncProjectContextAsync(string projectName)
        {
            try
            {
                var containerClient = _blobServiceClient.GetBlobContainerClient("project-knowledge");
                await containerClient.CreateIfNotExistsAsync();

                bool dbSynced = false;
                bool gitSynced = false;

                // 1. Resolve DB Connection Strings for Target Project
                var connectionStrings = new List<string>();

                string mainConn = _configuration[$"DbConnection_{projectName}_Main"] ?? _configuration["DbConnection_Main"] ?? "";
                string clientConn = _configuration[$"DbConnection_{projectName}_Client"] ?? _configuration["DbConnection_Client"] ?? "";

                if (!string.IsNullOrWhiteSpace(mainConn)) connectionStrings.Add(mainConn);
                if (!string.IsNullOrWhiteSpace(clientConn)) connectionStrings.Add(clientConn);

                if (connectionStrings.Count > 0)
                {
                    string schemaJson = await _schemaReader.ExtractSchemaJsonAsync(connectionStrings);
                    if (!string.IsNullOrWhiteSpace(schemaJson) && schemaJson != "{}")
                    {
                        string schemaBlobPath = $"{projectName.ToLower()}/db-schema.json";
                        var schemaBlobClient = containerClient.GetBlobClient(schemaBlobPath);
                        await schemaBlobClient.UploadAsync(BinaryData.FromString(schemaJson), overwrite: true);

                        _logger.LogInformation($"Successfully updated database schema snapshot at Blob: {schemaBlobPath}");
                        dbSynced = true;
                    }
                }
                else
                {
                    _logger.LogWarning($"No DB connection strings configured for project '{projectName}'.");
                }

                // 2. Resolve Git Repository and PAT Token for Target Project
                string gitRepo = _configuration[$"GitRepo_{projectName}"] ?? _configuration["GitRepoBaseUrl"] ?? "";
                string patToken = _configuration["GitPersonalAccessToken"] ?? "";

                if (!string.IsNullOrWhiteSpace(gitRepo))
                {
                    string? markdownContent = await _gitFetcher.FetchArchitectureMarkdownAsync(gitRepo, patToken, "architecture.md");
                    if (!string.IsNullOrWhiteSpace(markdownContent))
                    {
                        string archBlobPath = $"{projectName.ToLower()}/architecture.md";
                        var archBlobClient = containerClient.GetBlobClient(archBlobPath);
                        await archBlobClient.UploadAsync(BinaryData.FromString(markdownContent), overwrite: true);

                        _logger.LogInformation($"Successfully updated architecture.md snapshot at Blob: {archBlobPath}");
                        gitSynced = true;
                    }
                }
                else
                {
                    _logger.LogWarning($"No Git repo configured for project '{projectName}'.");
                }

                return new SyncResult
                {
                    Success = dbSynced || gitSynced,
                    Project = projectName,
                    DatabaseSchemaSynced = dbSynced,
                    ArchitectureDocSynced = gitSynced,
                    TimestampUtc = DateTime.UtcNow
                };
            }
            catch (Exception ex)
            {
                _logger.LogError($"Context sync failed for project '{projectName}': {ex.Message}");
                return new SyncResult
                {
                    Success = false,
                    Project = projectName,
                    Error = ex.Message,
                    TimestampUtc = DateTime.UtcNow
                };
            }
        }
    }

    public class SyncResult
    {
        public bool Success { get; set; }
        public string Project { get; set; } = "";
        public bool DatabaseSchemaSynced { get; set; }
        public bool ArchitectureDocSynced { get; set; }
        public string? Error { get; set; }
        public DateTime TimestampUtc { get; set; }
    }
}
