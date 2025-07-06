using DocMailer.Models;
using DocMailer.Services;
using DocMailer.Utils;
using System.Text;

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
                // Configure console encoding for proper UTF-8 character display
                Console.OutputEncoding = Encoding.UTF8;
                Console.InputEncoding = Encoding.UTF8;

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
                        case "send-to":
                            if (mainArgs.Length > 1)
                            {
                                await ProcessRecipients(SendMode.Specific, isDryRun, mainArgs[1]);
                            }
                            else
                            {
                                Console.WriteLine("Error: send-to command requires an email address.");
                                Console.WriteLine("Usage: dotnet run send-to email@example.com [--dry-run]");
                                return;
                            }
                            break;
                        case "send-thankyou":
                            await ProcessRecipients(SendMode.Thankyou, isDryRun);
                            break;
                        case "send-thankyou-to":
                            if (mainArgs.Length > 1)
                            {
                                await ProcessRecipientsWithThankyou(mainArgs[1], isDryRun);
                            }
                            else
                            {
                                Console.WriteLine("Error: send-thankyou-to command requires an email address.");
                                Console.WriteLine("Usage: dotnet run send-thankyou-to email@example.com [--dry-run]");
                                return;
                            }
                            break;
                        case "test":
                            await TestConfiguration(isDryRun);
                            break;
                        case "stats":
                            ShowStats();
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

        private static async Task ProcessRecipients(SendMode mode = SendMode.All, bool isDryRun = false, string? specificEmail = null)
        {
            var modeText = mode == SendMode.Specific && !string.IsNullOrEmpty(specificEmail) 
                ? $"Specific recipient ({specificEmail})" + (isDryRun ? " (DRY RUN)" : "")
                : (isDryRun ? $"{mode} mode (DRY RUN)" : $"{mode} mode");
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
            var recipients = FilterRecipients(allRecipients, mode, specificEmail);
            Logger.LogInfo($"Selected {recipients.Count} recipients for processing based on {mode} mode.");

            if (recipients.Count == 0)
            {
                if (mode == SendMode.Specific && !string.IsNullOrEmpty(specificEmail))
                {
                    Logger.LogInfo($"No recipient found with email: {specificEmail}");
                    Console.WriteLine($"❌ No recipient found with email: {specificEmail}");
                    Console.WriteLine("   Please check the email address and ensure it exists in your Excel file.");
                }
                else
                {
                    Logger.LogInfo("No recipients to process.");
                }
                return;
            }

            // Load templates
            var documentTemplate = _templateService.LoadDocumentTemplate(_config.DocumentTemplatePath);
            var emailTemplatePath = mode == SendMode.Thankyou ? _config.ThankyouTemplatePath : _config.EmailTemplatePath;
            var emailTemplate = _templateService.LoadEmailTemplate(emailTemplatePath);

            // Process each recipient
            var successCount = 0;
            var errorCount = 0;

            foreach (var recipient in recipients)
            {
                try
                {
                    Logger.LogInfo($"Processing: {recipient.Name} ({recipient.Email})");

                    string pdfPath = "";
                    
                    // Generate PDF document only for non-thank-you emails
                    if (mode != SendMode.Thankyou)
                    {
                        var documentContent = _templateService.ProcessDocumentTemplate(documentTemplate, recipient);
                        
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
                    }
                    else if (isDryRun)
                    {
                        Console.WriteLine($"  📧 Thank you email (no attachment)");
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
                            if (mode == SendMode.Thankyou)
                            {
                                Console.WriteLine($"     Attachment: None (thank you email)");
                            }
                            else if (!string.IsNullOrEmpty(pdfPath))
                            {
                                Console.WriteLine($"     Attachment: {Path.GetFileName(pdfPath)}");
                            }
                        }
                        else
                        {
                            // For thank you emails, don't send any attachment
                            var attachmentPath = mode == SendMode.Thankyou ? null : pdfPath;
                            await _emailService.SendEmailAsync(emailSubject, emailContent, recipient.Email, recipient.Name, fromEmail, fromName, attachmentPath);
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

        private static List<Recipient> FilterRecipients(List<Recipient> allRecipients, SendMode mode, string? specificEmail = null)
        {
            // First filter out canceled recipients from all operations
            var activeRecipients = allRecipients.Where(r => !r.IsCanceled).ToList();
            
            return mode switch
            {
                SendMode.All => activeRecipients,
                SendMode.NotSent => activeRecipients.Where(r => !r.LastSent.HasValue).ToList(),
                SendMode.NotResponded => activeRecipients.Where(r => r.LastSent.HasValue && (!r.Responded.HasValue || !r.Responded.Value)).ToList(),
                SendMode.Test => activeRecipients.Where(r => r.Email.Contains("test", StringComparison.OrdinalIgnoreCase) || 
                                                         r.Name.Contains("test", StringComparison.OrdinalIgnoreCase)).ToList(),
                SendMode.Thankyou => activeRecipients.Where(r => r.Responded.HasValue && r.Responded.Value).ToList(),
                SendMode.Specific when !string.IsNullOrEmpty(specificEmail) => activeRecipients.Where(r => 
                    r.Email.Equals(specificEmail, StringComparison.OrdinalIgnoreCase)).ToList(),
                _ => activeRecipients
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

        private static void ShowStats()
        {
            Logger.LogInfo("Generating campaign statistics...");

            try
            {
                // Check if Excel file exists
                if (!File.Exists(_config.ExcelFilePath))
                {
                    Console.WriteLine($"❌ Excel file not found: {_config.ExcelFilePath}");
                    Logger.LogError($"Excel file not found: {_config.ExcelFilePath}");
                    return;
                }

                // Read recipients from Excel
                var allRecipients = _excelReader.ReadRecipients(_config.ExcelFilePath);

                // Calculate statistics
                var totalRecipients = allRecipients.Count;
                var canceledRecipients = allRecipients.Where(r => r.IsCanceled).ToList();
                var activeRecipients = allRecipients.Where(r => !r.IsCanceled).ToList();
                var sentRecipients = activeRecipients.Where(r => r.LastSent.HasValue).ToList();
                var respondedRecipients = activeRecipients.Where(r => r.Responded.HasValue && r.Responded.Value).ToList();
                var notSentRecipients = activeRecipients.Where(r => !r.LastSent.HasValue).ToList();
                var sentButNotRespondedRecipients = activeRecipients.Where(r => r.LastSent.HasValue && (!r.Responded.HasValue || !r.Responded.Value)).ToList();

                // Calculate percentages (based on active recipients, excluding canceled)
                var activeRecipientsCount = activeRecipients.Count;
                var sentPercentage = activeRecipientsCount > 0 ? (sentRecipients.Count * 100.0) / activeRecipientsCount : 0;
                var respondedPercentage = activeRecipientsCount > 0 ? (respondedRecipients.Count * 100.0) / activeRecipientsCount : 0;
                var responseRateAmongSent = sentRecipients.Count > 0 ? (respondedRecipients.Count * 100.0) / sentRecipients.Count : 0;

                // Display statistics
                Console.WriteLine();
                Console.WriteLine("📊 ═══════════════════════════════════════════════════════════════");
                Console.WriteLine("📊                    DOCMAILER CAMPAIGN STATISTICS");
                Console.WriteLine("📊 ═══════════════════════════════════════════════════════════════");
                Console.WriteLine();

                // Overall Statistics
                Console.WriteLine("📈 OVERALL STATISTICS");
                Console.WriteLine("─────────────────────────────────────────────────────────────────");
                Console.WriteLine($"👥 Total Recipients:           {totalRecipients:N0}");
                Console.WriteLine($"✅ Active Recipients:          {activeRecipientsCount:N0}");
                if (canceledRecipients.Count > 0)
                {
                    Console.WriteLine($"❌ Canceled Recipients:       {canceledRecipients.Count:N0}");
                }
                Console.WriteLine();

                // Email Sending Statistics
                Console.WriteLine("📧 EMAIL SENDING STATUS");
                Console.WriteLine("─────────────────────────────────────────────────────────────────");
                Console.WriteLine($"✅ Emails Sent:               {sentRecipients.Count:N0} ({sentPercentage:F1}%)");
                Console.WriteLine($"⏳ Not Sent Yet:              {notSentRecipients.Count:N0} ({(100 - sentPercentage):F1}%)");
                Console.WriteLine();

                // Response Statistics
                Console.WriteLine("💬 RESPONSE STATISTICS (Active Recipients Only)");
                Console.WriteLine("─────────────────────────────────────────────────────────────────");
                Console.WriteLine($"🎯 Total Responses:           {respondedRecipients.Count:N0} ({respondedPercentage:F1}% of active)");
                Console.WriteLine($"📨 Response Rate (of sent):   {responseRateAmongSent:F1}%");
                Console.WriteLine($"🔄 Sent but No Response:      {sentButNotRespondedRecipients.Count:N0}");
                Console.WriteLine();

                // Progress Bar for Sending
                Console.WriteLine("📊 SENDING PROGRESS");
                Console.WriteLine("─────────────────────────────────────────────────────────────────");
                var progressBar = CreateProgressBar(sentPercentage, 50);
                Console.WriteLine($"[{progressBar}] {sentPercentage:F1}%");
                Console.WriteLine();

                // Progress Bar for Responses
                Console.WriteLine("📊 RESPONSE PROGRESS");
                Console.WriteLine("─────────────────────────────────────────────────────────────────");
                var responseProgressBar = CreateProgressBar(respondedPercentage, 50);
                Console.WriteLine($"[{responseProgressBar}] {respondedPercentage:F1}%");
                Console.WriteLine();

                // Recommendations
                Console.WriteLine("💡 RECOMMENDATIONS");
                Console.WriteLine("─────────────────────────────────────────────────────────────────");
                
                if (notSentRecipients.Count > 0)
                {
                    Console.WriteLine($"• Consider running: dotnet run send-not-sent");
                    Console.WriteLine($"  └─ This will send to {notSentRecipients.Count} recipients who haven't received emails yet");
                }

                if (sentButNotRespondedRecipients.Count > 0)
                {
                    Console.WriteLine($"• Consider running: dotnet run send-not-responded");
                    Console.WriteLine($"  └─ This will follow up with {sentButNotRespondedRecipients.Count} recipients who haven't responded");
                }

                if (responseRateAmongSent < 20 && sentRecipients.Count > 5)
                {
                    Console.WriteLine($"• Response rate is low ({responseRateAmongSent:F1}%). Consider:");
                    Console.WriteLine($"  └─ Reviewing email content for clarity");
                    Console.WriteLine($"  └─ Adding a clearer call-to-action");
                    Console.WriteLine($"  └─ Following up with non-responders");
                }

                if (totalRecipients == 0)
                {
                    Console.WriteLine($"• No recipients found. Check your Excel file format and data.");
                }

                Console.WriteLine();
                Console.WriteLine("📊 ═══════════════════════════════════════════════════════════════");

                Logger.LogInfo($"Campaign statistics: {totalRecipients} total ({activeRecipientsCount} active, {canceledRecipients.Count} canceled), {sentRecipients.Count} sent ({sentPercentage:F1}%), {respondedRecipients.Count} responded ({respondedPercentage:F1}%)");
            }
            catch (Exception ex)
            {
                Logger.LogError("Error generating statistics", ex);
                Console.WriteLine($"❌ Error generating statistics: {ex.Message}");
            }
        }

        private static string CreateProgressBar(double percentage, int width)
        {
            var filledWidth = (int)((percentage / 100.0) * width);
            var filled = new string('█', filledWidth);
            var empty = new string('░', width - filledWidth);
            return filled + empty;
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
            Console.WriteLine("  send-to <email>    - Send email to a specific recipient by email address");
            Console.WriteLine("  send-thankyou      - Send thank you emails to recipients who have responded");
            Console.WriteLine("  send-thankyou-to <email> - Send thank you email to a specific recipient who has responded");
            Console.WriteLine("  test              - Test configuration by sending a test email");
            Console.WriteLine("  stats             - Show campaign statistics and progress");
            Console.WriteLine("  help              - Show this help");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --dry-run          - Simulate operations without sending emails or updating files");
            Console.WriteLine("                       Can be used with any send command or test command");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  dotnet run send-all                         # Send to all recipients");
            Console.WriteLine("  dotnet run send-all --dry-run               # Preview what would be sent");
            Console.WriteLine("  dotnet run send-not-sent --dry-run          # Preview unsent recipients");
            Console.WriteLine("  dotnet run send-to john@example.com         # Send to specific recipient");
            Console.WriteLine("  dotnet run send-to john@example.com --dry-run # Preview send to specific recipient");
            Console.WriteLine("  dotnet run send-thankyou                    # Send thank you to responders");
            Console.WriteLine("  dotnet run send-thankyou --dry-run          # Preview thank you emails");
            Console.WriteLine("  dotnet run send-thankyou-to john@example.com # Send thank you to specific responder");
            Console.WriteLine("  dotnet run send-thankyou-to john@example.com --dry-run # Preview thank you to specific responder");
            Console.WriteLine("  dotnet run test --dry-run                   # Test configuration without sending");
            Console.WriteLine("  dotnet run stats                            # Show campaign statistics");
            Console.WriteLine();
            Console.WriteLine("Excel File Format:");
            Console.WriteLine("  Required columns: Name, Email");
            Console.WriteLine("  Optional columns: Company, Position, LastSent, Responded");
            Console.WriteLine("  - LastSent: Updated automatically with timestamp or error message");
            Console.WriteLine("  - Responded: Manual update (true/false/CANCELED/-1) to track responses");
            Console.WriteLine("    * TRUE/true/1/YES/Y = Responded");
            Console.WriteLine("    * FALSE/false/0/NO/N = Not responded");
            Console.WriteLine("    * CANCELED/-1 = Exclude from all operations (shown in stats only)");
            Console.WriteLine();
            Console.WriteLine("Configuration:");
            Console.WriteLine("  - Edit config.json file with your settings");
            Console.WriteLine("  - Place Excel file in Data/ folder");
            Console.WriteLine("  - Update template paths in config.json");
            Console.WriteLine("  - Generated files go to Output/ folder");
        }

        private static async Task ProcessRecipientsWithThankyou(string specificEmail, bool isDryRun = false)
        {
            var modeText = $"Send thank you to specific recipient ({specificEmail})" + (isDryRun ? " (DRY RUN)" : "");
            Logger.LogInfo($"Starting thank you email processing for {modeText}...");

            // Check if Excel file exists
            if (!File.Exists(_config.ExcelFilePath))
            {
                Logger.LogError($"Excel file not found: {_config.ExcelFilePath}");
                return;
            }

            // Read recipients from Excel
            var allRecipients = _excelReader.ReadRecipients(_config.ExcelFilePath);
            Logger.LogInfo($"Found {allRecipients.Count} total recipients.");

            // Find the specific recipient who has responded
            var recipient = allRecipients.FirstOrDefault(r => 
                r.Email.Equals(specificEmail, StringComparison.OrdinalIgnoreCase) && 
                !r.IsCanceled && 
                r.Responded.HasValue && 
                r.Responded.Value);

            if (recipient == null)
            {
                // Check if the recipient exists but hasn't responded
                var existingRecipient = allRecipients.FirstOrDefault(r => 
                    r.Email.Equals(specificEmail, StringComparison.OrdinalIgnoreCase) && 
                    !r.IsCanceled);

                if (existingRecipient == null)
                {
                    Logger.LogInfo($"No recipient found with email: {specificEmail}");
                    Console.WriteLine($"❌ No recipient found with email: {specificEmail}");
                    Console.WriteLine("   Please check the email address and ensure it exists in your Excel file.");
                }
                else if (!existingRecipient.Responded.HasValue || !existingRecipient.Responded.Value)
                {
                    Logger.LogInfo($"Recipient {specificEmail} has not responded yet.");
                    Console.WriteLine($"❌ Recipient {specificEmail} has not responded yet.");
                    Console.WriteLine("   Thank you emails are only sent to recipients who have responded (Responded = TRUE).");
                    Console.WriteLine("   Please update the 'Responded' column in your Excel file first.");
                }
                else if (existingRecipient.IsCanceled)
                {
                    Logger.LogInfo($"Recipient {specificEmail} is canceled.");
                    Console.WriteLine($"❌ Recipient {specificEmail} is canceled.");
                    Console.WriteLine("   Thank you emails are not sent to canceled recipients.");
                }
                return;
            }

            Logger.LogInfo($"Found recipient for thank you email: {recipient.Name} ({recipient.Email})");

            // Load thank you email template
            var emailTemplate = _templateService.LoadEmailTemplate(_config.ThankyouTemplatePath);

            try
            {
                Logger.LogInfo($"Processing thank you email for: {recipient.Name} ({recipient.Email})");

                // Process email template
                var emailContent = _templateService.ProcessEmailTemplate(emailTemplate, recipient);
                var emailSubject = _templateService.ProcessEmailSubject(emailTemplate.Subject, recipient);
                
                // Get sender info from template metadata
                var fromEmail = emailTemplate.Metadata.ContainsKey("fromEmail") ? emailTemplate.Metadata["fromEmail"].ToString() ?? "" : "";
                var fromName = emailTemplate.Metadata.ContainsKey("fromName") ? emailTemplate.Metadata["fromName"].ToString() ?? "" : "";

                // Send email (or simulate in dry run)
                if (_emailService != null)
                {
                    if (isDryRun)
                    {
                        Logger.LogInfo($"[DRY RUN] Would send thank you email to: {recipient.Email}");
                        Logger.LogInfo($"[DRY RUN] Subject: {emailSubject}");
                        Logger.LogInfo($"[DRY RUN] From: {fromName} <{fromEmail}>");
                        Console.WriteLine($"🔍 DRY RUN - Would send thank you email:");
                        Console.WriteLine($"  📧 To: {recipient.Email} ({recipient.Name})");
                        Console.WriteLine($"     Subject: {emailSubject}");
                        Console.WriteLine($"     From: {fromName} <{fromEmail}>");
                        Console.WriteLine($"     Template: {Path.GetFileName(_config.ThankyouTemplatePath)}");
                        Console.WriteLine("   Thank you email would be sent successfully!");
                    }
                    else
                    {
                        await _emailService.SendEmailAsync(emailSubject, emailContent, recipient.Email, recipient.Name, fromEmail, fromName);
                        Logger.LogInfo($"Thank you email sent to: {recipient.Email}");
                        Console.WriteLine($"✅ Thank you email sent to: {recipient.Email} ({recipient.Name})");
                        Console.WriteLine($"   Subject: {emailSubject}");
                        Console.WriteLine($"   From: {fromName} <{fromEmail}>");
                        
                        // Note: We don't update the Excel file for thank you emails as they don't affect LastSent status
                    }
                }

                Logger.LogInfo("Thank you email processing completed successfully.");
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error processing thank you email for {recipient.Name}: {ex.Message}", ex);
                Console.WriteLine($"❌ Error sending thank you email to {recipient.Email}: {ex.Message}");
            }
        }
    }
}
