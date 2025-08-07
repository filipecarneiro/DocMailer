using DocMailer.Models;
using DocMailer.Utils;
using OfficeOpenXml;

namespace DocMailer.Services
{
    /// <summary>
    /// Service for reading data from Excel files
    /// </summary>
    public class ExcelReaderService
    {
        public List<Recipient> ReadRecipients(string filePath)
        {
            var recipients = new List<Recipient>();

            // Configure EPPlus license context (for non-commercial use)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];
            
            // Assume first row contains headers
            var rowCount = worksheet.Dimension.Rows;
            var columnCount = worksheet.Dimension.Columns;
            
            // Find column indices
            var nameCol = FindColumnIndex(worksheet, "DisplayName") ?? FindColumnIndex(worksheet, "Name");
            var emailCol = FindColumnIndex(worksheet, "Email") ?? FindColumnIndex(worksheet, "EmailAddress");
            var companyCol = FindColumnIndex(worksheet, "Company");
            var positionCol = FindColumnIndex(worksheet, "Position");
            var fullNameCol = FindColumnIndex(worksheet, "FullName");
            var lastSentCol = FindColumnIndex(worksheet, "LastSent");
            var respondedCol = FindColumnIndex(worksheet, "Responded");
            
            // Log warning if no email column found
            if (!emailCol.HasValue)
            {
                Logger.LogInfo("Warning: No 'Email' or 'EmailAddress' column found in Excel file. All recipients will be skipped.");
            }
            
            for (int row = 2; row <= rowCount; row++)
            {
                // Check if email column exists
                if (!emailCol.HasValue)
                {
                    Logger.LogInfo($"Row {row}: Skipping recipient - no Email or EmailAddress column found in Excel file.");
                    continue;
                }
                
                var emailText = worksheet.Cells[row, emailCol.Value].Text?.Trim() ?? string.Empty;
                
                // Skip recipients without email addresses
                if (string.IsNullOrEmpty(emailText))
                {
                    Logger.LogInfo($"Row {row}: Skipping recipient - no email address provided.");
                    continue;
                }
                
                var nameText = string.Empty;
                
                // Try to get name from Name or DisplayName columns, fallback to email
                if (nameCol.HasValue)
                {
                    nameText = worksheet.Cells[row, nameCol.Value].Text?.Trim() ?? string.Empty;
                }
                
                // If name is empty or no name column found, use email as name
                if (string.IsNullOrEmpty(nameText))
                {
                    nameText = emailText;
                }
                
                var recipient = new Recipient
                {
                    DisplayName = nameText,
                    Email = emailText,
                    Company = companyCol.HasValue ? worksheet.Cells[row, companyCol.Value].Text?.Trim() ?? string.Empty : string.Empty,
                    Position = positionCol.HasValue ? worksheet.Cells[row, positionCol.Value].Text?.Trim() ?? string.Empty : string.Empty,
                    FullName = fullNameCol.HasValue ? worksheet.Cells[row, fullNameCol.Value].Text?.Trim() ?? string.Empty : string.Empty,
                    RowNumber = row
                };

                // Parse LastSent if column exists
                if (lastSentCol.HasValue && DateTime.TryParse(worksheet.Cells[row, lastSentCol.Value].Text, out var lastSent))
                {
                    recipient.LastSent = lastSent;
                }

                // Parse Responded if column exists - handle various true/false representations
                if (respondedCol.HasValue)
                {
                    var respondedText = worksheet.Cells[row, respondedCol.Value].Text?.Trim().ToUpperInvariant();
                    if (!string.IsNullOrEmpty(respondedText))
                    {
                        // Handle CANCELED status first
                        if (respondedText == "CANCELED" || respondedText == "-1")
                        {
                            recipient.IsCanceled = true;
                        }
                        // Handle various representations of true/false
                        else if (respondedText == "TRUE" || respondedText == "1" || respondedText == "YES" || 
                            respondedText == "Y")
                        {
                            recipient.Responded = true;
                        }
                        else if (respondedText == "FALSE" || respondedText == "0" || respondedText == "NO" || 
                                 respondedText == "N")
                        {
                            recipient.Responded = false;
                        }
                        // If none of the above, leave as null (unknown)
                    }
                }

                // Add custom fields from all columns, skipping known standard columns
                for (int col = 1; col <= columnCount; col++)
                {
                    var header = worksheet.Cells[1, col].Text?.Trim();
                    var value = worksheet.Cells[row, col].Text?.Trim();
                    
                    // Skip known standard columns
                    if ((nameCol.HasValue && col == nameCol.Value) || 
                        (emailCol.HasValue && col == emailCol.Value) || 
                        (companyCol.HasValue && col == companyCol.Value) || 
                        (positionCol.HasValue && col == positionCol.Value) || 
                        (fullNameCol.HasValue && col == fullNameCol.Value) || 
                        (lastSentCol.HasValue && col == lastSentCol.Value) || 
                        (respondedCol.HasValue && col == respondedCol.Value))
                        continue;
                    
                    // Add any other column as custom field
                    if (!string.IsNullOrEmpty(header) && !string.IsNullOrEmpty(value))
                    {
                        recipient.CustomFields[header] = value;
                    }
                }

                // Log custom fields count
                if (recipient.CustomFields.Count > 0)
                {
                    Logger.LogInfo($"Recipient {recipient.DisplayName} has {recipient.CustomFields.Count} custom fields.");
                }

                recipients.Add(recipient);
            }

            return recipients;
        }

        public void UpdateRecipientStatus(string filePath, Recipient recipient, bool success, string? errorMessage = null)
        {
            // Configure EPPlus license context (for non-commercial use)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];
            
            // Find or create LastSent column
            var lastSentCol = FindColumnIndex(worksheet, "LastSent");
            if (!lastSentCol.HasValue)
            {
                lastSentCol = worksheet.Dimension.Columns + 1;
                worksheet.Cells[1, lastSentCol.Value].Value = "LastSent";
            }

            // Update the cell based on success/failure
            if (success)
            {
                worksheet.Cells[recipient.RowNumber, lastSentCol.Value].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            }
            else
            {
                worksheet.Cells[recipient.RowNumber, lastSentCol.Value].Value = $"ERROR: {errorMessage}";
            }

            package.Save();
        }

        private int? FindColumnIndex(ExcelWorksheet worksheet, string columnName)
        {
            var columnCount = worksheet.Dimension.Columns;
            for (int col = 1; col <= columnCount; col++)
            {
                if (string.Equals(worksheet.Cells[1, col].Text, columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }
            return null;
        }
    }
}
