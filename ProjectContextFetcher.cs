using System;
using System.Linq;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SrsAi.Functions
{
    public class ProjectContextFetcher
    {
        private readonly ILogger _logger;
        private readonly BlobContainerClient _knowledgeContainer;

        public ProjectContextFetcher(IConfiguration configuration, ILogger logger)
        {
            _logger = logger;
            string connectionString = configuration["AzureWebJobsStorage"];
            var blobServiceClient = new BlobServiceClient(connectionString);

            // Container storing pre-compiled project knowledge snapshots
            _knowledgeContainer = blobServiceClient.GetBlobContainerClient("project-knowledge");
        }

        public async Task<string> GetProjectSnapshotAsync(string projectName, params string[] fileNameCandidates)
        {
            // Support case variations for project folder names (e.g. "RSCS", "rscs", "Rscs")
            string[] projectFolderCandidates = new[]
            {
                projectName,
                projectName.ToLowerInvariant(),
                projectName.ToUpperInvariant()
            };

            foreach (var folder in projectFolderCandidates.Distinct())
            {
                foreach (var fileName in fileNameCandidates)
                {
                    string blobName = $"{folder}/{fileName}";
                    _logger.LogInformation($"Fetching snapshot context from Blob: {blobName}");

                    try
                    {
                        var blobClient = _knowledgeContainer.GetBlobClient(blobName);

                        if (await blobClient.ExistsAsync())
                        {
                            var downloadResult = await blobClient.DownloadContentAsync();
                            _logger.LogInformation($"Successfully retrieved snapshot: {blobName}");
                            return downloadResult.Value.Content.ToString();
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error reading snapshot {blobName}: {ex.Message}");
                    }
                }
            }

            string primaryName = fileNameCandidates.Length > 0 ? fileNameCandidates[0] : "unknown";
            _logger.LogWarning($"Snapshot files [{string.Join(", ", fileNameCandidates)}] for project {projectName} (or rscs) not found in Blob Storage container 'project-knowledge'.");
            return $"[Context Missing: {primaryName} for project {projectName} is not available.]";
        }
    }
}