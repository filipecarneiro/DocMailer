using DocMailer.Models;
using System.Net;
using System.Net.Mail;
using Markdig;

namespace DocMailer.Services
{
    /// <summary>
    /// Service for sending emails
    /// </summary>
    public class EmailService
    {
        private readonly EmailConfig _config;
        private readonly MarkdownPipeline _markdownPipeline;

        public EmailService(EmailConfig config)
        {
            _config = config;
            _markdownPipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .Build();
        }

        public async Task SendEmailAsync(string subject, string markdownContent, string recipientEmail, string recipientName, string fromEmail, string fromName, string? attachmentPath = null)
        {
            // Convert Markdown to HTML for personal, clean emails without custom styling
            var htmlContent = Markdown.ToHtml(markdownContent, _markdownPipeline);

            using var client = new SmtpClient(_config.SmtpServer, _config.SmtpPort)
            {
                EnableSsl = _config.EnableSsl,
                Credentials = new NetworkCredential(_config.Username, _config.Password)
            };

            using var message = new MailMessage
            {
                From = new MailAddress(fromEmail, fromName),
                Subject = subject,
                Body = htmlContent,
                IsBodyHtml = true
            };

            message.To.Add(new MailAddress(recipientEmail, recipientName));

            // Add attachment if provided
            if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
            {
                message.Attachments.Add(new Attachment(attachmentPath));
            }

            await client.SendMailAsync(message);
        }
    }
}
