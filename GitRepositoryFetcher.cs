using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SrsAi.Functions
{
    public class GitRepositoryFetcher
    {
        private readonly ILogger _logger;
        private readonly IConfiguration _configuration;
        private static readonly HttpClient _httpClient = new HttpClient();

        public GitRepositoryFetcher(IConfiguration configuration, ILogger logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public async Task<string?> FetchArchitectureMarkdownAsync(string repoOwnerAndName, string patToken, string filePath = "architecture.md")
        {
            if (string.IsNullOrWhiteSpace(repoOwnerAndName))
            {
                _logger.LogWarning("Git repository owner/name is empty. Skipping GitHub fetch.");
                return null;
            }

            // Clean up repo string if full URL was provided (e.g., https://github.com/StudioRG/RSCSLive.git -> StudioRG/RSCSLive)
            string cleanRepo = repoOwnerAndName
                .Replace("https://github.com/", "")
                .Replace(".git", "")
                .Trim('/');

            string requestUrl = $"https://api.github.com/repos/{cleanRepo}/contents/{filePath}";

            _logger.LogInformation($"Fetching '{filePath}' from GitHub repository: {cleanRepo} via API: {requestUrl}");

            using var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
            request.Headers.UserAgent.Add(new ProductInfoHeaderValue("SrsAiPilot", "1.0"));
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github.v3.raw"));

            if (!string.IsNullOrWhiteSpace(patToken) && patToken != "tumhara_git_pat")
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", patToken);
            }

            try
            {
                var response = await _httpClient.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    string markdownContent = await response.Content.ReadAsStringAsync();
                    _logger.LogInformation($"Successfully fetched '{filePath}' ({markdownContent.Length} bytes) from GitHub repo '{cleanRepo}'.");
                    return markdownContent;
                }

                string errorMsg = await response.Content.ReadAsStringAsync();
                _logger.LogWarning($"GitHub API fetch failed with status code {response.StatusCode}: {errorMsg}");
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Exception occurred while fetching '{filePath}' from GitHub: {ex.Message}");
                return null;
            }
        }
    }
}
