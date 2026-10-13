using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SrsAi.Functions
{
    public class TeamsNotificationService
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger _logger;
        private static readonly HttpClient _httpClient = new HttpClient();

        public TeamsNotificationService(IConfiguration configuration, ILogger logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public async Task SendSrsDraftCardAsync(
            string projectName,
            string meetingId,
            string meetingTitle,
            string executiveSummary,
            int epicsCount,
            int storiesCount,
            int tablesCount,
            string docxSasUrl,
            string draftBlobName)
        {
            string webhookUrl = _configuration["TeamsWebhookUrl"] ?? "";
            if (string.IsNullOrWhiteSpace(webhookUrl))
            {
                _logger.LogWarning("TeamsWebhookUrl is not configured. Skipping Teams Adaptive Card notification.");
                return;
            }

            string publicApiBaseUrl = _configuration["PublicApiBaseUrl"]?.TrimEnd('/') 
                ?? "https://srs-ai-pilot-fn-dmdyc4f7cradawa0.westus3-01.azurewebsites.net";

            string editorUrl = $"{publicApiBaseUrl}/api/srs/editor?draftId={Uri.EscapeDataString(draftBlobName)}";
            string approveUrl = $"{publicApiBaseUrl}/api/srs/approve?draftId={Uri.EscapeDataString(draftBlobName)}";

            string truncatedSummary = executiveSummary;
            if (truncatedSummary.Length > 300)
            {
                truncatedSummary = truncatedSummary.Substring(0, 297) + "...";
            }

            var adaptiveCardPayload = new
            {
                type = "message",
                attachments = new[]
                {
                    new
                    {
                        contentType = "application/vnd.microsoft.card.adaptive",
                        content = new
                        {
                            type = "AdaptiveCard",
                            version = "1.4",
                            schema = "http://adaptivecards.io/schemas/adaptive-card.json",
                            body = new object[]
                            {
                                new
                                {
                                    type = "Container",
                                    style = "emphasis",
                                    items = new object[]
                                    {
                                        new
                                        {
                                            type = "ColumnSet",
                                            columns = new object[]
                                            {
                                                new
                                                {
                                                    type = "Column",
                                                    width = "stretch",
                                                    items = new object[]
                                                    {
                                                        new
                                                        {
                                                            type = "TextBlock",
                                                            text = $"📄 SRS DRAFT READY FOR REVIEW",
                                                            weight = "Bolder",
                                                            size = "Medium",
                                                            color = "Accent"
                                                        },
                                                        new
                                                        {
                                                            type = "TextBlock",
                                                            text = $"Project: **{projectName}** | Meeting: {meetingId}",
                                                            spacing = "None",
                                                            isSubtle = true
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                },
                                new
                                {
                                    type = "TextBlock",
                                    text = string.IsNullOrWhiteSpace(meetingTitle) ? $"Meeting {meetingId}" : meetingTitle,
                                    weight = "Bolder",
                                    size = "Large",
                                    wrap = true
                                },
                                new
                                {
                                    type = "TextBlock",
                                    text = "**Executive Summary:**",
                                    weight = "Bolder",
                                    spacing = "Medium"
                                },
                                new
                                {
                                    type = "TextBlock",
                                    text = truncatedSummary,
                                    wrap = true
                                },
                                new
                                {
                                    type = "FactSet",
                                    spacing = "Medium",
                                    facts = new[]
                                    {
                                        new { title = "Target Project:", value = projectName },
                                        new { title = "Epics Identified:", value = epicsCount.ToString() },
                                        new { title = "User Stories:", value = storiesCount.ToString() },
                                        new { title = "Database Tables:", value = tablesCount.ToString() },
                                        new { title = "Status:", value = "Pending Human Review" }
                                    }
                                }
                            },
                            actions = new object[]
                            {
                                new
                                {
                                    type = "Action.OpenUrl",
                                    title = "📥 Download Word (.docx)",
                                    url = docxSasUrl
                                },
                                new
                                {
                                    type = "Action.OpenUrl",
                                    title = "✏️ Review & Edit Document",
                                    url = editorUrl
                                },
                                new
                                {
                                    type = "Action.OpenUrl",
                                    title = "✅ Approve SRS",
                                    url = approveUrl
                                }
                            }
                        }
                    }
                }
            };

            await PostToWebhookAsync(webhookUrl, adaptiveCardPayload, "SRS Draft Card");
        }

        public async Task SendApprovalConfirmationAsync(
            string projectName,
            string meetingId,
            string approverName,
            string approvedBlobName,
            string finalDocxSasUrl)
        {
            string webhookUrl = _configuration["TeamsWebhookUrl"] ?? "";
            if (string.IsNullOrWhiteSpace(webhookUrl)) return;

            var approvalCardPayload = new
            {
                type = "message",
                attachments = new[]
                {
                    new
                    {
                        contentType = "application/vnd.microsoft.card.adaptive",
                        content = new
                        {
                            type = "AdaptiveCard",
                            version = "1.4",
                            schema = "http://adaptivecards.io/schemas/adaptive-card.json",
                            body = new object[]
                            {
                                new
                                {
                                    type = "Container",
                                    style = "good",
                                    items = new object[]
                                    {
                                        new
                                        {
                                            type = "TextBlock",
                                            text = $"🎉 SRS DOCUMENT OFFICIALLY APPROVED",
                                            weight = "Bolder",
                                            size = "Medium",
                                            color = "Good"
                                        },
                                        new
                                        {
                                            type = "TextBlock",
                                            text = $"Project: **{projectName}** | Meeting: {meetingId}",
                                            isSubtle = true,
                                            spacing = "None"
                                        }
                                    }
                                },
                                new
                                {
                                    type = "FactSet",
                                    spacing = "Medium",
                                    facts = new[]
                                    {
                                        new { title = "Approved By:", value = approverName },
                                        new { title = "Timestamp (UTC):", value = DateTime.UtcNow.ToString("u") },
                                        new { title = "Storage Container:", value = "srs-approved" },
                                        new { title = "Document File:", value = approvedBlobName }
                                    }
                                }
                            },
                            actions = new object[]
                            {
                                new
                                {
                                    type = "Action.OpenUrl",
                                    title = "📥 Download Final Approved (.docx)",
                                    url = finalDocxSasUrl
                                }
                            }
                        }
                    }
                }
            };

            await PostToWebhookAsync(webhookUrl, approvalCardPayload, "Approval Confirmation Card");
        }

        private async Task PostToWebhookAsync(string webhookUrl, object payload, string cardType)
        {
            try
            {
                string jsonBody = JsonSerializer.Serialize(payload);
                using var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(webhookUrl, content);
                if (response.IsSuccessStatusCode)
                {
                    _logger.LogInformation($"Successfully posted Teams {cardType} to incoming webhook.");
                }
                else
                {
                    string err = await response.Content.ReadAsStringAsync();
                    _logger.LogWarning($"Teams webhook call returned {response.StatusCode}: {err}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Exception posting Teams {cardType}: {ex.Message}");
            }
        }
    }
}
