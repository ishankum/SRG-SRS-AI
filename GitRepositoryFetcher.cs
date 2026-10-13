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

        public async Task<string?> FetchArchitectureMarkdownAsync(string repoOwnerAndName, string patToken, string filePath = "Azmbl_StudioRG/Azmbl_StudioRG/Content/architecture.md", string? branch = null)
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

            string cleanPath = (filePath ?? "Azmbl_StudioRG/Azmbl_StudioRG/Content/architecture.md").TrimStart('/');

            // 1. Try initial requested path and branch
            string initialUrl = BuildContentsApiUrl(cleanRepo, cleanPath, branch);
            bool hasToken = !string.IsNullOrWhiteSpace(patToken) && patToken != "tumhara_git_pat";
            _logger.LogInformation($"Fetching '{cleanPath}' from GitHub repository: {cleanRepo} via API: {initialUrl} (Auth Token Configured: {hasToken})");

            string? content = await FetchRawContentAsync(initialUrl, patToken);
            if (!string.IsNullOrWhiteSpace(content))
            {
                _logger.LogInformation($"Successfully fetched '{cleanPath}' ({content.Length} bytes) from GitHub repo '{cleanRepo}'.");
                return content;
            }

            // 2. If not found, probe common branches (main, master, dev, development) and path variations
            var candidateBranches = new List<string>();
            if (!string.IsNullOrWhiteSpace(branch))
            {
                candidateBranches.Add(branch);
            }
            foreach (var b in new[] { "main", "master", "dev", "development", "staging" })
            {
                if (!candidateBranches.Contains(b, StringComparer.OrdinalIgnoreCase))
                {
                    candidateBranches.Add(b);
                }
            }

            string repoName = cleanRepo.Contains('/') ? cleanRepo.Split('/')[1] : cleanRepo;
            var candidatePaths = new List<string> { cleanPath };

            void AddPathVariation(string p)
            {
                if (!candidatePaths.Contains(p, StringComparer.OrdinalIgnoreCase))
                {
                    candidatePaths.Add(p);
                }
            }

            AddPathVariation($"Azmbl_{repoName}/Azmbl_{repoName}/Content/architecture.md");
            AddPathVariation($"Azmbl_{repoName}/Content/architecture.md");
            AddPathVariation("Azmbl_StudioRG/Azmbl_StudioRG/Content/architecture.md");
            AddPathVariation("Azmbl_StudioRG/Content/architecture.md");
            AddPathVariation("Content/architecture.md");
            AddPathVariation("architecture.md");

            foreach (var candidateBranch in candidateBranches)
            {
                foreach (var pathVariation in candidatePaths)
                {
                    if (pathVariation == cleanPath && candidateBranch == branch)
                        continue;

                    string testUrl = BuildContentsApiUrl(cleanRepo, pathVariation, candidateBranch);
                    string? candidateContent = await FetchRawContentAsync(testUrl, patToken);
                    if (!string.IsNullOrWhiteSpace(candidateContent))
                    {
                        _logger.LogInformation($"Successfully found and fetched architecture file at '{pathVariation}' on branch '{candidateBranch}' from '{cleanRepo}'.");
                        return candidateContent;
                    }
                }
            }

            // 3. Diagnostic check: probe repository endpoint to detect PAT permission issues
            await RunRepositoryDiagnosticsAsync(cleanRepo, patToken, cleanPath);

            return null;
        }

        private static string BuildContentsApiUrl(string repo, string path, string? branch)
        {
            string url = $"https://api.github.com/repos/{repo}/contents/{path}";
            if (!string.IsNullOrWhiteSpace(branch))
            {
                url += $"?ref={Uri.EscapeDataString(branch)}";
            }
            return url;
        }

        private async Task<string?> FetchRawContentAsync(string requestUrl, string patToken)
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
            request.Headers.UserAgent.Add(new ProductInfoHeaderValue("SrsAiPilot", "1.0"));
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github.v3.raw"));
            request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");

            SetAuthorizationHeader(request, patToken);

            try
            {
                var response = await _httpClient.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    string markdownContent = await response.Content.ReadAsStringAsync();

                    // If GitHub returned JSON metadata with base64 encoded content instead of raw text
                    if (markdownContent.TrimStart().StartsWith("{") && markdownContent.Contains("\"content\"") && markdownContent.Contains("\"encoding\""))
                    {
                        try
                        {
                            using var jsonDoc = System.Text.Json.JsonDocument.Parse(markdownContent);
                            if (jsonDoc.RootElement.TryGetProperty("content", out var base64Prop))
                            {
                                string rawBase64 = base64Prop.GetString()?.Replace("\n", "").Replace("\r", "") ?? "";
                                byte[] bytes = Convert.FromBase64String(rawBase64);
                                markdownContent = System.Text.Encoding.UTF8.GetString(bytes);
                            }
                        }
                        catch (Exception parseEx)
                        {
                            _logger.LogWarning($"Failed to decode base64 GitHub content, using raw payload: {parseEx.Message}");
                        }
                    }

                    return markdownContent;
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Request failed for {requestUrl}: {ex.Message}");
                return null;
            }
        }

        private async Task RunRepositoryDiagnosticsAsync(string cleanRepo, string patToken, string cleanPath)
        {
            try
            {
                string repoInfoUrl = $"https://api.github.com/repos/{cleanRepo}";
                using var request = new HttpRequestMessage(HttpMethod.Get, repoInfoUrl);
                request.Headers.UserAgent.Add(new ProductInfoHeaderValue("SrsAiPilot", "1.0"));
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
                request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");

                SetAuthorizationHeader(request, patToken);

                var response = await _httpClient.SendAsync(request);
                string responseBody = await response.Content.ReadAsStringAsync();

                if (response.StatusCode == System.Net.HttpStatusCode.NotFound || response.StatusCode == System.Net.HttpStatusCode.Unauthorized || response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                {
                    _logger.LogError($"[PAT Diagnostic] GitHub API returned {response.StatusCode} when querying repository info at '{repoInfoUrl}'. This indicates the configured PAT token does not have access to repository '{cleanRepo}' or the repository is private and the token lacks 'Contents: Read-only' permission. Error: {responseBody}");
                }
                else if (response.IsSuccessStatusCode)
                {
                    try
                    {
                        using var doc = System.Text.Json.JsonDocument.Parse(responseBody);
                        string defaultBranch = doc.RootElement.TryGetProperty("default_branch", out var b) ? b.GetString() ?? "unknown" : "unknown";
                        _logger.LogWarning($"[PAT Diagnostic] Repository '{cleanRepo}' is accessible! Default branch is '{defaultBranch}'. However, '{cleanPath}' was not found in the tested branches. Please verify the exact branch name and file path in the repository.");
                    }
                    catch
                    {
                        _logger.LogWarning($"[PAT Diagnostic] Repository '{cleanRepo}' is accessible, but file '{cleanPath}' was not found.");
                    }
                }
            }
            catch (Exception diagEx)
            {
                _logger.LogWarning($"Could not complete GitHub repository diagnostic: {diagEx.Message}");
            }
        }

        private static void SetAuthorizationHeader(HttpRequestMessage request, string patToken)
        {
            if (!string.IsNullOrWhiteSpace(patToken) && patToken != "tumhara_git_pat")
            {
                string scheme = patToken.StartsWith("ghp_") ? "token" : "Bearer";
                request.Headers.Authorization = new AuthenticationHeaderValue(scheme, patToken);
            }
        }
    }
}
