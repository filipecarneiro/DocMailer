using DocMailer.Models;
using System.Text.RegularExpressions;

namespace DocMailer.Services
{
    /// <summary>
    /// Service for processing Markdown templates
    /// </summary>
    public class TemplateService
    {
        public string ProcessDocumentTemplate(DocumentTemplate template, Recipient recipient)
        {
            var content = RemoveYamlMetadata(template.Content);
            
            // Replace basic placeholders
            content = ReplacePlaceholders(content, recipient);
            
            return content;
        }

        public string ProcessEmailTemplate(EmailTemplate template, Recipient recipient)
        {
            var content = RemoveYamlMetadata(template.Content);
            
            // Replace basic placeholders
            content = ReplacePlaceholders(content, recipient);
            
            // Replace sender placeholders from template metadata
            content = ReplaceSenderPlaceholders(content, template);
            
            return content;
        }

        public string ProcessEmailSubject(string subject, Recipient recipient)
        {
            return ReplacePlaceholders(subject, recipient);
        }

        private string ReplacePlaceholders(string content, Recipient recipient)
        {
            // Basic placeholders - trim whitespace to prevent Markdown formatting issues
            content = content.Replace("{{DisplayName}}", recipient.DisplayName?.Trim() ?? "");
            content = content.Replace("{{Name}}", recipient.DisplayName?.Trim() ?? ""); // Keep for backward compatibility
            content = content.Replace("{{FirstName}}", recipient.FirstName?.Trim() ?? "");
            content = content.Replace("{{FullName}}", recipient.GetFirstName());
            content = content.Replace("{{Email}}", recipient.Email?.Trim() ?? "");
            content = content.Replace("{{Company}}", recipient.Company?.Trim() ?? "");
            content = content.Replace("{{Position}}", recipient.Position?.Trim() ?? "");
            content = content.Replace("{{SubscriptionDate}}", recipient.SubscriptionDate.ToString("dd/MM/yyyy"));
            content = content.Replace("{{CurrentDate}}", DateTime.Now.ToString("dd/MM/yyyy"));
            content = content.Replace("{{CurrentDateLong}}", FormatDateLongPortuguese(DateTime.Now));

            // Custom field placeholders - trim whitespace from custom fields too
            foreach (var field in recipient.CustomFields)
            {
                var fieldValue = field.Value?.ToString()?.Trim() ?? "";
                content = content.Replace($"{{{{{field.Key}}}}}", fieldValue);
            }

            return content;
        }

        public DocumentTemplate LoadDocumentTemplate(string filePath)
        {
            var content = File.ReadAllText(filePath);
            var metadata = ExtractMetadata(content);
            
            return new DocumentTemplate
            {
                Name = Path.GetFileNameWithoutExtension(filePath),
                FilePath = filePath,
                Content = content,
                Metadata = metadata,
                LastModified = File.GetLastWriteTime(filePath)
            };
        }

        public EmailTemplate LoadEmailTemplate(string filePath)
        {
            var content = File.ReadAllText(filePath);
            var metadata = ExtractMetadata(content);
            
            return new EmailTemplate
            {
                Name = Path.GetFileNameWithoutExtension(filePath),
                FilePath = filePath,
                Content = content,
                Subject = metadata.ContainsKey("subject") ? metadata["subject"].ToString() ?? "" : "",
                Metadata = metadata,
                LastModified = File.GetLastWriteTime(filePath)
            };
        }

        private Dictionary<string, object> ExtractMetadata(string content)
        {
            var metadata = new Dictionary<string, object>();
            
            // Extract YAML metadata from the beginning of the file
            var yamlMatch = Regex.Match(content, @"^---\s*\n(.*?)\n---\s*\n", RegexOptions.Singleline);
            
            if (yamlMatch.Success)
            {
                var yamlContent = yamlMatch.Groups[1].Value;
                var lines = yamlContent.Split('\n');
                
                foreach (var line in lines)
                {
                    var parts = line.Split(':', 2);
                    if (parts.Length == 2)
                    {
                        var key = parts[0].Trim();
                        var value = parts[1].Trim().Trim('"', '\'');
                        metadata[key] = value;
                    }
                }
            }
            
            return metadata;
        }

        private string RemoveYamlMetadata(string content)
        {
            // Remove YAML metadata from the beginning of the content
            var yamlMatch = Regex.Match(content, @"^---\s*\n(.*?)\n---\s*\n", RegexOptions.Singleline);
            
            if (yamlMatch.Success)
            {
                return content.Substring(yamlMatch.Length);
            }
            
            return content;
        }

        private string FormatDateLongPortuguese(DateTime date)
        {
            var months = new[]
            {
                "", "janeiro", "fevereiro", "março", "abril", "maio", "junho",
                "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
            };

            return $"{date.Day} de {months[date.Month]} de {date.Year}";
        }

        private string ReplaceSenderPlaceholders(string content, EmailTemplate template)
        {
            // Replace sender placeholders from template metadata
            var fromName = template.Metadata.ContainsKey("fromName") ? template.Metadata["fromName"].ToString() ?? "" : "";
            var fromEmail = template.Metadata.ContainsKey("fromEmail") ? template.Metadata["fromEmail"].ToString() ?? "" : "";
            
            content = content.Replace("{{FromName}}", fromName.Trim());
            content = content.Replace("{{FromEmail}}", fromEmail.Trim());
            
            return content;
        }
    }
}
