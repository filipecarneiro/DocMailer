namespace DocMailer.Models
{
    /// <summary>
    /// Application configuration settings
    /// </summary>
    public class AppConfig
    {
        public EmailConfig Email { get; set; } = new EmailConfig();
        public string ExcelFilePath { get; set; } = string.Empty;
        public string OutputDirectory { get; set; } = "Output";
        public string DocumentTemplatePath { get; set; } = "Templates/document.md";
        public string EmailTemplatePath { get; set; } = "Templates/email.md";
    }

    public class EmailConfig
    {
        public string SmtpServer { get; set; } = string.Empty;
        public int SmtpPort { get; set; } = 587;
        public bool EnableSsl { get; set; } = true;
        public string Username { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
    }
}
