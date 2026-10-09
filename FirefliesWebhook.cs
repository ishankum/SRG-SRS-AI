using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Azure.Storage.Queues;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SrsAi.Functions
{
    public class FirefliesWebhook
    {
        private readonly ILogger _logger;
        private readonly IConfiguration _configuration;

        public FirefliesWebhook(ILoggerFactory loggerFactory, IConfiguration configuration)
        {
            _logger = loggerFactory.CreateLogger<FirefliesWebhook>();
            _configuration = configuration;
        }

        [Function("FirefliesWebhook")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "integrations/fireflies/webhook")] HttpRequestData req)
        {
            _logger.LogInformation("Received Fireflies webhook request.");

            // 1. Read raw body for HMAC signature validation
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            if (string.IsNullOrWhiteSpace(requestBody))
            {
                _logger.LogWarning("Webhook request body is empty.");
                var badRequest = req.CreateResponse(HttpStatusCode.BadRequest);
                await badRequest.WriteStringAsync("Error: Request body is empty. Please provide a valid JSON payload.");
                return badRequest;
            }

            if (!ValidateSignature(req, requestBody))
            {
                _logger.LogWarning("Invalid X-Hub-Signature validation failed.");
                var unauthorizedResponse = req.CreateResponse(HttpStatusCode.Unauthorized);
                await unauthorizedResponse.WriteStringAsync("Error: Unauthorized. X-Hub-Signature verification failed.");
                return unauthorizedResponse;
            }

            try
            {
                // 2. Safely initialize QueueClient
                string connectionString = _configuration["AzureWebJobsStorage"];
                if (string.IsNullOrWhiteSpace(connectionString) || connectionString.Contains("UseDevelopmentStorage=true", StringComparison.OrdinalIgnoreCase))
                {
                    _logger.LogError("AzureWebJobsStorage is missing or set to UseDevelopmentStorage=true in Azure Portal environment variables.");
                    var configErrorResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
                    await configErrorResponse.WriteStringAsync("Configuration Error: 'AzureWebJobsStorage' is not configured in Azure Portal.");
                    return configErrorResponse;
                }

                var queueClient = new QueueClient(connectionString, "srs-processing", new QueueClientOptions
                {
                    MessageEncoding = QueueMessageEncoding.Base64
                });

                // 3. Parse event, meeting_id, and timestamp safely (handling String & Number types)
                using JsonDocument doc = JsonDocument.Parse(requestBody);
                JsonElement root = doc.RootElement;

                string meetingId = GetSafeString(root, "meeting_id", "meetingId", "id");
                if (string.IsNullOrWhiteSpace(meetingId))
                {
                    meetingId = $"TEST-{DateTime.UtcNow:yyyyMMddHHmmss}";
                    _logger.LogWarning($"meeting_id not found in payload. Generated fallback: {meetingId}");
                }

                string eventType = GetSafeString(root, "event", "eventType");
                if (string.IsNullOrWhiteSpace(eventType))
                {
                    eventType = "meeting.transcribed";
                }

                string timestamp = GetSafeString(root, "timestamp");
                if (string.IsNullOrWhiteSpace(timestamp))
                {
                    timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString();
                }

                var queueMessage = new
                {
                    MeetingId = meetingId,
                    Event = eventType,
                    ReceivedTimestamp = timestamp,
                    Status = "TRANSCRIPT_AVAILABLE"
                };

                // 4. Durably enqueue the event before returning 2xx
                await queueClient.CreateIfNotExistsAsync();
                await queueClient.SendMessageAsync(JsonSerializer.Serialize(queueMessage));

                _logger.LogInformation($"Successfully enqueued meeting_id: {meetingId}");

                var successResponse = req.CreateResponse(HttpStatusCode.OK);
                successResponse.Headers.Add("Content-Type", "application/json");
                await successResponse.WriteStringAsync(JsonSerializer.Serialize(new
                {
                    status = "success",
                    message = "Webhook accepted and enqueued successfully",
                    meetingId = meetingId
                }));
                return successResponse;
            }
            catch (JsonException jsonEx)
            {
                _logger.LogError($"Invalid JSON payload received: {jsonEx.Message}");
                var badRequestResponse = req.CreateResponse(HttpStatusCode.BadRequest);
                await badRequestResponse.WriteStringAsync($"Error: Payload is not valid JSON. {jsonEx.Message}");
                return badRequestResponse;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to process webhook: {ex.Message}");
                var errorResponse = req.CreateResponse(HttpStatusCode.InternalServerError);
                await errorResponse.WriteStringAsync($"Internal server error: {ex.Message}");
                return errorResponse;
            }
        }

        private static string GetSafeString(JsonElement root, params string[] propertyNames)
        {
            foreach (var name in propertyNames)
            {
                if (root.TryGetProperty(name, out JsonElement element))
                {
                    if (element.ValueKind == JsonValueKind.String)
                    {
                        return element.GetString() ?? "";
                    }
                    if (element.ValueKind == JsonValueKind.Number)
                    {
                        return element.GetRawText();
                    }
                }
            }
            return "";
        }

        private bool ValidateSignature(HttpRequestData req, string body)
        {
            string secret = _configuration["FirefliesWebhookSecret"] ?? "";

            // If secret is default test secret or unconfigured, log warning and allow request
            if (string.IsNullOrEmpty(secret) || secret == "test_secret_for_local_development")
            {
                _logger.LogWarning("FirefliesWebhookSecret is unconfigured or in test mode. Bypassing signature validation for testing.");
                return true;
            }

            // Check signature headers
            string? signatureHeader = null;
            if (req.Headers.TryGetValues("X-Hub-Signature", out var h1)) signatureHeader = string.Join(",", h1);
            else if (req.Headers.TryGetValues("X-Hub-Signature-256", out var h2)) signatureHeader = string.Join(",", h2);
            else if (req.Headers.TryGetValues("X-Fireflies-Signature", out var h3)) signatureHeader = string.Join(",", h3);

            if (string.IsNullOrEmpty(signatureHeader))
            {
                _logger.LogWarning("Signature header missing from request.");
                return false;
            }

            // Strip "sha256=" prefix if Fireflies prepends it
            if (signatureHeader.StartsWith("sha256=", StringComparison.OrdinalIgnoreCase))
            {
                signatureHeader = signatureHeader.Substring(7).Trim();
            }

            using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(secret));
            byte[] hashBytes = hmac.ComputeHash(Encoding.UTF8.GetBytes(body));
            string computedSignature = BitConverter.ToString(hashBytes).Replace("-", "").ToLower();

            bool isValid = computedSignature.Equals(signatureHeader, StringComparison.OrdinalIgnoreCase);
            if (!isValid)
            {
                _logger.LogWarning($"Signature mismatch! Computed: {computedSignature}, Received: {signatureHeader}");
            }

            return isValid;
        }
    }
}