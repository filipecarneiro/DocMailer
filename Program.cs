using DocMailer.Models;
using DocMailer.Services;
using DocMailer.Utils;

namespace DocMailer
{
    /// <summary>
    /// Main DocMailer application
    /// </summary>
    public class Program
    {
        private static AppConfig _config = new();
        private static ExcelReaderService _excelReader = new();
        private static TemplateService _templateService = new();
        private static PdfGeneratorService _pdfGenerator = new();
        private static EmailService? _emailService;

        public static async Task Main(string[] args)
        {
            try
            {
                Logger.LogInfo("Starting DocMailer...");

                // Load configuration
                _config = ConfigHelper.LoadConfig<AppConfig>("config.json");
                _emailService = new EmailService(_config.Email);

                // Ensure required directories exist
                EnsureDirectoriesExist();

                // Parse command line arguments
                bool isDryRun = args.Contains("--dry-run", StringComparer.OrdinalIgnoreCase);
                var mainArgs = args.Where(arg => !arg.Equals("--dry-run", StringComparison.OrdinalIgnoreCase)).ToArray();

                if (isDryRun)
                {
                    Logger.LogInfo("DRY RUN MODE - No emails will be sent, no Excel files will be updated");
                    Console.WriteLine("🔍 DRY RUN MODE - Simulating operations without sending emails or updating files");
                    Console.WriteLine();
                }

                // Process command line arguments
                if (mainArgs.Length > 0)
                {
                    switch (mainArgs[0].ToLower())
                    {
                        case "send-all":
                            await ProcessRecipients(SendMode.All, isDryRun);
                            break;
                        case "send-not-sent":
                            await ProcessRecipients(SendMode.NotSent, isDryRun);
                            break;
                        case "send-not-responded":
                            await ProcessRecipients(SendMode.NotResponded, isDryRun);
                            break;
                        case "send-test":
                            await ProcessRecipients(SendMode.Test, isDryRun);
                            break;
                        case "test":
                            await TestConfiguration(isDryRun);
                            break;
                        case "help":
                        default:
                            ShowHelp();
                            break;
                    }
                }
                else
                {
                    ShowHelp();
                }

                Logger.LogInfo("DocMailer completed successfully.");
            }
            catch (Exception ex)
            {
                Logger.LogError("Critical application error", ex);
                Console.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
            }
        }

        private static async Task ProcessRecipients(SendMode mode = SendMode.All, bool isDryRun = false)
        {
            var modeText = isDryRun ? $"{mode} mode (DRY RUN)" : $"{mode} mode";
            Logger.LogInfo($"Starting recipient processing in {modeText}...");

            // Check if Excel file exists
            if (!File.Exists(_config.ExcelFilePath))
            {
                Logger.LogError($"Excel file not found: {_config.ExcelFilePath}");
                return;
            }

            // Read recipients from Excel
            var allRecipients = _excelReader.ReadRecipients(_config.ExcelFilePath);
            Logger.LogInfo($"Found {allRecipients.Count} total recipients.");

            // Filter recipients based on mode
            var recipients = FilterRecipients(allRecipients, mode);
            Logger.LogInfo($"Selected {recipients.Count} recipients for processing based on {mode} mode.");

            if (recipients.Count == 0)
            {
                Logger.LogInfo("No recipients to process.");
                return;
            }

            // Load templates
            var documentTemplate = _templateService.LoadDocumentTemplate(_config.DocumentTemplatePath);
            var emailTemplate = _templateService.LoadEmailTemplate(_config.EmailTemplatePath);

            // Process each recipient
            var successCount = 0;
            var errorCount = 0;

            foreach (var recipient in recipients)
            {
                try
                {
                    Logger.LogInfo($"Processing: {recipient.Name} ({recipient.Email})");

                    // Generate PDF document
                    var documentContent = _templateService.ProcessDocumentTemplate(documentTemplate, recipient);
                    
                    string pdfPath;
                    if (isDryRun)
                    {
                        // In dry run mode, don't actually generate the PDF, just show what would be generated
                        var documentTitle = documentTemplate.Metadata.ContainsKey("title") ? 
                            documentTemplate.Metadata["title"].ToString() ?? "Document" : "Document";
                        var safeTitle = documentTitle.Replace(" ", "_");
                        var safeName = recipient.Name.Replace(" ", "_");
                        pdfPath = Path.Combine(_config.OutputDirectory, $"{safeTitle}-{safeName}.pdf");
                        Logger.LogInfo($"[DRY RUN] Would generate PDF: {pdfPath}");
                        Console.WriteLine($"  📄 Would generate: {Path.GetFileName(pdfPath)}");
                    }
                    else
                    {
                        pdfPath = _pdfGenerator.GeneratePdf(documentContent, _config.OutputDirectory, recipient, documentTemplate);
                        Logger.LogInfo($"PDF generated: {pdfPath}");
                    }

                    // Process email template
                    var emailContent = _templateService.ProcessEmailTemplate(emailTemplate, recipient);
                    var emailSubject = _templateService.ProcessEmailSubject(emailTemplate.Subject, recipient);
                    
                    // Get sender info from template metadata
                    var fromEmail = emailTemplate.Metadata.ContainsKey("fromEmail") ? emailTemplate.Metadata["fromEmail"].ToString() ?? "" : "";
                    var fromName = emailTemplate.Metadata.ContainsKey("fromName") ? emailTemplate.Metadata["fromName"].ToString() ?? "" : "";

                    // Send email with attachment (or simulate in dry run)
                    if (_emailService != null)
                    {
                        if (isDryRun)
                        {
                            Logger.LogInfo($"[DRY RUN] Would send email to: {recipient.Email}");
                            Logger.LogInfo($"[DRY RUN] Subject: {emailSubject}");
                            Logger.LogInfo($"[DRY RUN] From: {fromName} <{fromEmail}>");
                            Console.WriteLine($"  📧 Would send to: {recipient.Email}");
                            Console.WriteLine($"     Subject: {emailSubject}");
                            Console.WriteLine($"     From: {fromName} <{fromEmail}>");
                        }
                        else
                        {
                            await _emailService.SendEmailAsync(emailSubject, emailContent, recipient.Email, recipient.Name, fromEmail, fromName, pdfPath);
                            Logger.LogInfo($"Email sent to: {recipient.Email}");
                            
                            // Update Excel with success
                            _excelReader.UpdateRecipientStatus(_config.ExcelFilePath, recipient, true);
                        }
                    }

                    successCount++;
                }
                catch (Exception ex)
                {
                    Logger.LogError($"Error processing {recipient.Name}: {ex.Message}", ex);
                    
                    if (!isDryRun)
                    {
                        // Update Excel with error (only in real mode)
                        _excelReader.UpdateRecipientStatus(_config.ExcelFilePath, recipient, false, ex.Message);
                    }
                    else
                    {
                        Console.WriteLine($"  ❌ Would record error: {ex.Message}");
                    }
                    
                    errorCount++;
                }
            }

            var resultText = isDryRun ? "Dry run completed" : "Processing completed";
            Logger.LogInfo($"{resultText}. Successes: {successCount}, Errors: {errorCount}");
            
            if (isDryRun)
            {
                Console.WriteLine();
                Console.WriteLine($"🔍 Dry run summary: {successCount} would succeed, {errorCount} would have errors");
                Console.WriteLine("   No files were created or modified.");
            }
        }

        private static List<Recipient> FilterRecipients(List<Recipient> allRecipients, SendMode mode)
        {
            return mode switch
            {
                SendMode.All => allRecipients,
                SendMode.NotSent => allRecipients.Where(r => !r.LastSent.HasValue).ToList(),
                SendMode.NotResponded => allRecipients.Where(r => r.LastSent.HasValue && (!r.Responded.HasValue || !r.Responded.Value)).ToList(),
                SendMode.Test => allRecipients.Where(r => r.Email.Contains("test", StringComparison.OrdinalIgnoreCase) || 
                                                         r.Name.Contains("test", StringComparison.OrdinalIgnoreCase)).ToList(),
                _ => allRecipients
            };
        }

        private static async Task TestConfiguration(bool isDryRun = false)
        {
            var testModeText = isDryRun ? "Testing configuration (DRY RUN)..." : "Testing configuration...";
            Logger.LogInfo(testModeText);

            try
            {
                // Test email connection (send test email)
                var testRecipient = new Recipient
                {
                    Name = "Test User",
                    Email = _config.Email.Username, // Use the configured email as test recipient
                    Company = "Test Company",
                    Position = "Developer"
                };

                var testContent = "# Test Email\n\nThis is a test email from DocMailer.\n\n**Name:** {{Name}}\n**Email:** {{Email}}";
                var testTemplate = new EmailTemplate 
                { 
                    Content = testContent, 
                    Subject = "DocMailer Test - {{Name}}",
                    Metadata = new Dictionary<string, object>
                    {
                        { "fromEmail", _config.Email.Username },
                        { "fromName", "DocMailer Test" }
                    }
                };
                
                var processedContent = _templateService.ProcessEmailTemplate(testTemplate, testRecipient);
                var processedSubject = _templateService.ProcessEmailSubject(testTemplate.Subject, testRecipient);

                if (_emailService != null)
                {
                    if (isDryRun)
                    {
                        Logger.LogInfo("[DRY RUN] Would send test email");
                        Logger.LogInfo($"[DRY RUN] To: {testRecipient.Email}");
                        Logger.LogInfo($"[DRY RUN] Subject: {processedSubject}");
                        Logger.LogInfo($"[DRY RUN] From: DocMailer Test <{_config.Email.Username}>");
                        Console.WriteLine("🔍 Configuration test (DRY RUN):");
                        Console.WriteLine($"  📧 Would send test email to: {testRecipient.Email}");
                        Console.WriteLine($"     Subject: {processedSubject}");
                        Console.WriteLine($"     SMTP Server: {_config.Email.SmtpServer}:{_config.Email.SmtpPort}");
                        Console.WriteLine($"     SSL Enabled: {_config.Email.EnableSsl}");
                        Console.WriteLine("   Test email would be sent successfully!");
                    }
                    else
                    {
                        await _emailService.SendEmailAsync(processedSubject, processedContent, testRecipient.Email, testRecipient.Name, 
                            _config.Email.Username, "DocMailer Test");
                        Logger.LogInfo("Test email sent successfully!");
                    }
                }
            }
            catch (Exception ex)
            {
                if (isDryRun)
                {
                    Logger.LogError("[DRY RUN] Configuration test would fail", ex);
                    Console.WriteLine($"❌ Configuration test would fail: {ex.Message}");
                }
                else
                {
                    Logger.LogError("Error in configuration test", ex);
                }
            }
        }

        private static void EnsureDirectoriesExist()
        {
            Directory.CreateDirectory(_config.OutputDirectory);
            Directory.CreateDirectory(Path.GetDirectoryName(_config.DocumentTemplatePath) ?? "Templates");
            Directory.CreateDirectory(Path.GetDirectoryName(_config.EmailTemplatePath) ?? "Templates");
            Directory.CreateDirectory("Data");
        }

        private static void ShowHelp()
        {
            Console.WriteLine("=== DocMailer - Document and Email Generator ===");
            Console.WriteLine();
            Console.WriteLine("Usage: DocMailer.exe [command]");
            Console.WriteLine();
            Console.WriteLine("Available commands:");
            Console.WriteLine("  send-all           - Send emails to all recipients");
            Console.WriteLine("  send-not-sent      - Send emails only to recipients not previously sent");
            Console.WriteLine("  send-not-responded - Send emails only to recipients who haven't responded");
            Console.WriteLine("  send-test          - Send emails only to test recipients (name/email contains 'test')");
            Console.WriteLine("  test              - Test configuration by sending a test email");
            Console.WriteLine("  help              - Show this help");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --dry-run          - Simulate operations without sending emails or updating files");
            Console.WriteLine("                       Can be used with any send command or test command");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  dotnet run send-all                    # Send to all recipients");
            Console.WriteLine("  dotnet run send-all --dry-run          # Preview what would be sent");
            Console.WriteLine("  dotnet run send-not-sent --dry-run     # Preview unsent recipients");
            Console.WriteLine("  dotnet run test --dry-run              # Test configuration without sending");
            Console.WriteLine();
            Console.WriteLine("Excel File Format:");
            Console.WriteLine("  Required columns: Name, Email");
            Console.WriteLine("  Optional columns: Company, Position, LastSent, Responded");
            Console.WriteLine("  - LastSent: Updated automatically with timestamp or error message");
            Console.WriteLine("  - Responded: Manual update (true/false) to track responses");
            Console.WriteLine();
            Console.WriteLine("Configuration:");
            Console.WriteLine("  - Edit config.json file with your settings");
            Console.WriteLine("  - Place Excel file in Data/ folder");
            Console.WriteLine("  - Update template paths in config.json");
            Console.WriteLine("  - Generated files go to Output/ folder");
        }
    }
}
