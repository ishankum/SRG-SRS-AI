using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SrsAi.Functions
{
    public class EmailNotificationService
    {
        private readonly ILogger _logger;
        private readonly IConfiguration _configuration;

        public EmailNotificationService(IConfiguration configuration, ILogger logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public async Task SendSrsDraftEmailAsync(string meetingId, string projectName, string jsonPayload, byte[] docxBytes, string docxFileName)
        {
            string reviewerEmailsStr = _configuration["ReviewerEmails"] ?? "";
            string smtpHost = _configuration["SmtpHost"] ?? "";

            if (string.IsNullOrWhiteSpace(reviewerEmailsStr))
            {
                _logger.LogWarning("ReviewerEmails is not configured in environment variables. Skipping email notification.");
                return;
            }

            if (string.IsNullOrWhiteSpace(smtpHost))
            {
                _logger.LogWarning("SmtpHost is not configured in environment variables. Skipping email delivery.");
                return;
            }

            try
            {
                string senderEmail = _configuration["SenderEmail"] ?? _configuration["SmtpUsername"] ?? "noreply@srs-ai.com";
                int smtpPort = int.TryParse(_configuration["SmtpPort"], out int port) ? port : 587;
                string smtpUsername = _configuration["SmtpUsername"] ?? "";
                string smtpPassword = _configuration["SmtpPassword"] ?? "";
                bool enableSsl = bool.TryParse(_configuration["EnableSsl"], out bool ssl) ? ssl : true;

                using MailMessage message = new MailMessage();
                message.From = new MailAddress(senderEmail, "SRS AI Automation Agent");

                // Add reviewer recipients (supports comma-separated emails)
                string[] emails = reviewerEmailsStr.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string email in emails)
                {
                    string trimmed = email.Trim();
                    if (!string.IsNullOrWhiteSpace(trimmed))
                    {
                        message.To.Add(trimmed);
                    }
                }

                if (message.To.Count == 0)
                {
                    _logger.LogWarning("No valid email recipients found in ReviewerEmails.");
                    return;
                }

                message.Subject = $"[SRS Draft Review Required] Project: {projectName} - Meeting: {meetingId}";
                message.IsBodyHtml = true;
                message.Body = BuildHtmlEmailBody(projectName, meetingId, jsonPayload);

                // Attach the editable Word document (.docx)
                if (docxBytes != null && docxBytes.Length > 0)
                {
                    using MemoryStream attachmentStream = new MemoryStream(docxBytes);
                    Attachment attachment = new Attachment(
                        attachmentStream,
                        docxFileName,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    );
                    message.Attachments.Add(attachment);

                    await SendEmailWithSmtpAsync(smtpHost, smtpPort, smtpUsername, smtpPassword, enableSsl, message);
                }
                else
                {
                    await SendEmailWithSmtpAsync(smtpHost, smtpPort, smtpUsername, smtpPassword, enableSsl, message);
                }

                _logger.LogInformation($"Successfully sent SRS draft notification email to {reviewerEmailsStr} for meeting {meetingId}.");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to send SRS draft email notification: {ex.Message}");
                // Log and continue so email errors do not fail the SRS Blob creation worker
            }
        }

        private async Task SendEmailWithSmtpAsync(string host, int port, string username, string password, bool enableSsl, MailMessage message)
        {
            using SmtpClient smtp = new SmtpClient(host, port)
            {
                EnableSsl = enableSsl,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false
            };

            if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
            {
                smtp.Credentials = new NetworkCredential(username, password);
            }

            await smtp.SendMailAsync(message);
        }

        private static string BuildHtmlEmailBody(string project, string meetingId, string jsonPayload)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(@"
<html>
<body style='font-family: Arial, sans-serif; color: #333333; line-height: 1.6; max-width: 650px; margin: 0 auto; padding: 20px;'>
    <div style='background-color: #1F4E78; color: #ffffff; padding: 15px 20px; border-radius: 6px 6px 0 0;'>
        <h2 style='margin: 0; font-size: 20px;'>Software Requirements Specification (SRS) Draft</h2>
        <p style='margin: 5px 0 0 0; font-size: 13px; color: #D9E1F2;'>Action Required: Human Review & Approval</p>
    </div>

    <div style='border: 1px solid #D9D9D9; border-top: none; padding: 20px; border-radius: 0 0 6px 6px; background-color: #FAFAFA;'>
        <table style='width: 100%; border-collapse: collapse; margin-bottom: 20px;'>
            <tr>
                <td style='padding: 8px; font-weight: bold; width: 140px; background-color: #F2F4F7; border-bottom: 1px solid #E5E5E5;'>Project Name:</td>
                <td style='padding: 8px; background-color: #F2F4F7; border-bottom: 1px solid #E5E5E5;'>").Append(WebUtility.HtmlEncode(project)).Append(@"</td>
            </tr>
            <tr>
                <td style='padding: 8px; font-weight: bold; background-color: #FFFFFF; border-bottom: 1px solid #E5E5E5;'>Meeting ID:</td>
                <td style='padding: 8px; background-color: #FFFFFF; border-bottom: 1px solid #E5E5E5;'>").Append(WebUtility.HtmlEncode(meetingId)).Append(@"</td>
            </tr>
            <tr>
                <td style='padding: 8px; font-weight: bold; background-color: #F2F4F7; border-bottom: 1px solid #E5E5E5;'>Generated Date:</td>
                <td style='padding: 8px; background-color: #F2F4F7; border-bottom: 1px solid #E5E5E5;'>").Append(DateTime.UtcNow.ToString("dd-MMM-yyyy HH:mm UTC")).Append(@"</td>
            </tr>
        </table>

        <h3 style='color: #1F4E78; font-size: 16px; border-bottom: 2px solid #1F4E78; padding-bottom: 5px;'>Summary of Extracted Requirements</h3>
");

            try
            {
                using JsonDocument doc = JsonDocument.Parse(jsonPayload);
                JsonElement root = doc.RootElement;

                if (root.TryGetProperty("requirements", out var reqs) && reqs.ValueKind == JsonValueKind.Array)
                {
                    sb.Append("<ul style='padding-left: 20px;'>");
                    foreach (var req in reqs.EnumerateArray())
                    {
                        string id = req.TryGetProperty("id", out var idElem) ? idElem.GetString() ?? "REQ" : "REQ";
                        string title = req.TryGetProperty("title", out var titleElem) ? titleElem.GetString() ?? "" : "";
                        string classification = req.TryGetProperty("classification", out var cElem) ? cElem.GetString() ?? "NEW" : "NEW";

                        sb.Append("<li style='margin-bottom: 8px;'>")
                          .Append("<strong>").Append(WebUtility.HtmlEncode(id)).Append(": ").Append(WebUtility.HtmlEncode(title)).Append("</strong> ")
                          .Append("<span style='background-color: #2F5597; color: #ffffff; padding: 2px 6px; border-radius: 4px; font-size: 11px;'>")
                          .Append(WebUtility.HtmlEncode(classification.ToUpper()))
                          .Append("</span></li>");
                    }
                    sb.Append("</ul>");
                }

                if (root.TryGetProperty("openQuestions", out var questions) && questions.ValueKind == JsonValueKind.Array && questions.GetArrayLength() > 0)
                {
                    sb.Append("<div style='background-color: #FFF2CC; border-left: 4px solid #FFC000; padding: 12px; margin-top: 15px;'>")
                      .Append("<h4 style='margin: 0 0 5px 0; color: #7F6000; font-size: 14px;'>Open Questions / Review Decisions Needed:</h4>")
                      .Append("<ul style='margin: 0; padding-left: 20px;'>");
                    foreach (var q in questions.EnumerateArray())
                    {
                        sb.Append("<li>").Append(WebUtility.HtmlEncode(q.GetString() ?? "")).Append("</li>");
                    }
                    sb.Append("</ul></div>");
                }
            }
            catch
            {
                sb.Append("<p>Extracted requirements payload generated successfully.</p>");
            }

            sb.Append(@"
        <div style='margin-top: 25px; padding-top: 15px; border-top: 1px solid #D9D9D9; font-size: 12px; color: #595959;'>
            <p style='margin: 0;'>📎 <strong>Attachment:</strong> The complete editable Microsoft Word SRS document (<code>.docx</code>) is attached to this email.</p>
            <p style='margin: 5px 0 0 0;'>Please review and submit approvals or modifications.</p>
        </div>
    </div>
</body>
</html>");

            return sb.ToString();
        }
    }
}
