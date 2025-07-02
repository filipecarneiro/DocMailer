# DocMailer

A C# application that processes Excel data to generate personalized PDF documents and send automated emails with attachments.

[![.NET](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com/download)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![GitHub](https://img.shields.io/badge/GitHub-filipecarneiro/DocMailer-blue.svg)](https://github.com/filipecarneiro/DocMailer)

## 🚀 Features

- **Excel Integration**: Reads recipient data from Excel files
- **Markdown Templates**: Uses Markdown templates with YAML metadata for documents and emails
- **PDF Generation**: Converts Markdown to styled PDF documents
- **Email Automation**: Sends personalized emails with PDF attachments
- **Flexible Configuration**: JSON-based configuration system
- **Logging System**: Comprehensive logging for monitoring and debugging
## 📁 Project Structure

```
DocMailer/
├── Models/                    # Data models
│   ├── Recipient.cs          # Recipient data model
│   ├── DocumentTemplate.cs   # Document template model
│   ├── EmailTemplate.cs      # Email template model
│   ├── AppConfig.cs          # Application configuration
│   └── SendMode.cs           # Email sending modes
├── Services/                  # Core services
│   ├── ExcelReaderService.cs # Excel file processing
│   ├── TemplateService.cs    # Template processing
│   ├── PdfGeneratorService.cs# PDF generation
│   └── EmailService.cs       # Email sending
├── Utils/                     # Utility classes
│   ├── ConfigHelper.cs       # Configuration helper
│   └── Logger.cs             # Logging system
├── Templates/                 # Template files
│   ├── document.example.md   # Example document template
│   ├── email.example.md      # Example email template
│   ├── document.example.css  # Example CSS stylesheet
│   ├── document.md           # Your document template (created from example)
│   ├── email.md              # Your email template (created from example)
│   └── document.css          # Your CSS stylesheet (created from example)
├── Data/                      # Data files
│   ├── recipients.example.xlsx # Example Excel file
│   └── recipients.xlsx       # Your Excel file (created from example)
├── Output/                    # Generated files
│   ├── *.pdf                 # Generated PDF documents
│   └── *.log                 # Log files
├── config.example.json        # Example configuration
└── config.json               # Your configuration (created from example)
```

## 📋 Prerequisites

- .NET 8.0 or later
- SMTP email account (Gmail, Outlook, etc.)
- Excel file with recipient data

## 🛠️ Installation

### Option 1: Clone from GitHub
```bash
git clone https://github.com/filipecarneiro/DocMailer.git
cd DocMailer
```

### Option 2: Download ZIP
1. Download the latest release from [GitHub](https://github.com/filipecarneiro/DocMailer)
2. Extract the ZIP file
3. Navigate to the project directory

### Build the Project
```bash
# Restore dependencies and build
dotnet restore
dotnet build

# Or run setup script
# Windows: setup.bat
# Linux/Mac: ./setup.sh
```

## 🚀 Quick Start

### 1. Setup Configuration
```bash
# Copy and customize configuration
cp config.example.json config.json
# Edit config.json with your email settings
```

### 2. Setup Excel Data
```bash
# Copy and customize recipient data
cp Data/recipients.example.xlsx Data/recipients.xlsx
# Edit recipients.xlsx with your recipient data
```

### 3. Setup Templates
```bash
# Copy and customize templates
cp Templates/document.example.md Templates/document.md
cp Templates/email.example.md Templates/email.md
cp Templates/document.example.css Templates/document.css
# Edit the templates with your content
```

### 4. Test Your Setup
```bash
# Test configuration and send a test email
dotnet run test
```

### 5. Send to All Recipients
```bash
# Process all recipients
dotnet run send-all
```

## ⚙️ Configuration

1. **Copy the example configuration file:**
   ```bash
   cp config.example.json config.json
   ```

2. **Edit `config.json` with your settings:**

```json
{
  "email": {
    "smtpServer": "smtp.gmail.com",
    "smtpPort": 587,
    "enableSsl": true,
    "username": "your_email@gmail.com",
    "password": "your_app_password"
  },
  "excelFilePath": "Data/recipients.xlsx",
  "outputDirectory": "Output",
  "documentTemplatePath": "Templates/document.md",
  "emailTemplatePath": "Templates/email.md"
}
```

### Gmail Setup
For Gmail, use an App Password instead of your regular password:
1. Enable 2-factor authentication
2. Generate an App Password in Google Account settings
3. Use the App Password in the config file

## 📊 Excel File Format

### Setup Excel Data

1. **Copy the example Excel file:**
   ```bash
   cp Data/recipients.example.xlsx Data/recipients.xlsx
   ```

2. **Edit `recipients.xlsx` with your data** or use the example structure:

| Name | Email | Company | Position | ClientID | ProjectTitle | StartDate | Duration | TotalValue |
|------|-------|---------|----------|----------|--------------|-----------|----------|------------|
| John Smith | john.smith@acmecorp.com | ACME Corporation | Project Manager | CLI001 | Website Development | 2025-07-15 | 3 months | $15,000 |
| Sarah Johnson | sarah.j@techstartup.io | TechStartup Inc | CEO | CLI002 | Mobile App Development | 2025-08-01 | 6 months | $45,000 |

**Required columns:** Name, Email  
**Optional columns:** Company, Position, and any custom fields you need

**System columns (automatically managed):**
- **LastSent** - Last email sent timestamp (updated automatically)
- **Responded** - Manual update to track responses (TRUE/true/1/YES/Y for responded, FALSE/false/0/NO/N for not responded)

**Note**: The "Responded" column is used by the `stats` and `send-not-responded` commands to track which recipients have responded to your emails.

## 📝 Templates

### Setup Templates

1. **Copy the example template files:**
   ```bash
   cp Templates/document.example.md Templates/document.md
   cp Templates/email.example.md Templates/email.md
   cp Templates/document.example.css Templates/document.css
   ```

2. **Customize your templates** by editing the copied files with your content.

### Document Template (`Templates/document.md`)
```markdown
---
title: "Service Agreement"
type: "service_agreement"
version: "1.0"
author: "Your Company Name"
date: "{{CurrentDate}}"
pageSize: "A4"
headerCenter: "Templates/logo.png"
styleSheet: "Templates/document.css"
---

# Service Agreement

**Client:** {{Name}}
**Company:** {{Company}}
**Date:** {{CurrentDate}}

Your document content here...
```

### Email Template (`Templates/email.md`)
```markdown
---
subject: "Service Agreement - {{Company}} - Please Review and Sign"
type: "service_agreement_email"
fromEmail: "your-email@company.com"
fromName: "Your Name"
---

Dear {{FirstName}},

Your personalized email content here...

Best regards,
{{FromName}}
```

### Available Placeholders
- `{{Name}}` - Full recipient name
- `{{FirstName}}` - First name extracted from full name
- `{{Email}}` - Recipient email address
- `{{Company}}` - Recipient company
- `{{Position}}` - Recipient position/title
- `{{CurrentDate}}` - Current date (dd/MM/yyyy format)
- `{{CurrentDateLong}}` - Current date in long format
- `{{ClientID}}` - Client identification number
- `{{ProjectTitle}}` - Project title
- `{{StartDate}}` - Project start date
- `{{Duration}}` - Project duration
- `{{TotalValue}}` - Total project value
- `{{PaymentSchedule}}` - Payment schedule
- `{{PaymentDueDate}}` - Payment due date
- `{{ServiceDescription1/2/3}}` - Service descriptions
- `{{CustomFieldName}}` - Any custom field from Excel

## 🏃‍♂️ Usage

### Available Commands

| Command | Description | Dry Run Support |
|---------|-------------|-----------------|
| `test` | Test configuration by sending a test email | ✅ |
| `send-all` | Send emails to all recipients | ✅ |
| `send-not-sent` | Send emails only to recipients not previously sent | ✅ |
| `send-not-responded` | Send emails only to recipients who haven't responded | ✅ |
| `send-test` | Send emails only to test recipients (name/email contains 'test') | ✅ |
| `send-to <email>` | Send email to a specific recipient by email address | ✅ |
| `stats` | Show comprehensive campaign statistics and progress | ❌ |
| `help` | Show help information | ❌ |

### Test Configuration
```bash
dotnet run test
```
Sends a test email to verify your configuration.

### Dry Run Mode (Preview)
Add `--dry-run` to any command to simulate operations without sending emails or updating files:

```bash
# Preview what would be sent to all recipients
dotnet run send-all --dry-run

# Preview configuration test
dotnet run test --dry-run

# Preview what would be sent to unsent recipients
dotnet run send-not-sent --dry-run

# Preview sending to a specific recipient
dotnet run send-to recipient@example.com --dry-run
```

### Send to All Recipients
```bash
dotnet run send-all
```
Processes all recipients from the Excel file.

### Send to Unsent Recipients Only
```bash
dotnet run send-not-sent
```
Only sends to recipients who haven't received emails yet.

### Send to Non-Responders
```bash
dotnet run send-not-responded
```
Sends to recipients who received emails but haven't responded.

### Send Test Email
```bash
dotnet run send-test
```
Sends a test email to the first recipient for testing purposes.

### Send to Specific Recipient
```bash
dotnet run send-to recipient@example.com
```
Sends email to a specific recipient by email address. Useful for:
- Resending to recipients with corrected data
- Testing with specific individuals
- Handling special cases or follow-ups

```bash
# Preview before sending to specific recipient
dotnet run send-to recipient@example.com --dry-run

# Send to specific recipient (real mode)
dotnet run send-to recipient@example.com
```

### Campaign Statistics
```bash
dotnet run stats
```
Displays comprehensive campaign statistics including:
- **Total Recipients**: Complete count of recipients in your Excel file
- **Email Sending Status**: Number and percentage of emails sent vs. not sent
- **Response Statistics**: Number and percentage of recipients who have responded
- **Progress Bars**: Visual representation of sending and response progress
- **Recommendations**: Suggested next actions based on current campaign status

Example output:
```
📊 DOCMAILER CAMPAIGN STATISTICS
📊 ═══════════════════════════════════════════════════════════════

📈 OVERALL STATISTICS
─────────────────────────────────────────────────────────────────
👥 Total Recipients:           29

📧 EMAIL SENDING STATUS
─────────────────────────────────────────────────────────────────
✅ Emails Sent:               29 (100.0%)
⏳ Not Sent Yet:              0 (0.0%)

💬 RESPONSE STATISTICS
─────────────────────────────────────────────────────────────────
🎯 Total Responses:           4 (13.8% of all)
📨 Response Rate (of sent):   13.8%
🔄 Sent but No Response:      25

📊 SENDING PROGRESS
─────────────────────────────────────────────────────────────────
[██████████████████████████████████████████████████] 100.0%

📊 RESPONSE PROGRESS
─────────────────────────────────────────────────────────────────
[██████░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░] 13.8%

💡 RECOMMENDATIONS
─────────────────────────────────────────────────────────────────
• Consider running: dotnet run send-not-responded
  └─ This will follow up with 25 recipients who haven't responded
```

**Note**: To track responses, manually update the "Responded" column in your Excel file with TRUE/true/1/YES/Y for recipients who have responded.

### Command Options
- `--dry-run`: Preview mode - shows what would be done without actually sending emails or updating files

### Process Flow
Each command will:
1. Read recipient data from Excel
2. Generate personalized PDF for each recipient *(skipped in dry-run)*
3. Send email with PDF attachment *(simulated in dry-run)*
4. Update Excel with sending status and timestamp *(skipped in dry-run)*
5. Log all activities to console and file

## 📦 Dependencies

- **EPPlus** - Excel file processing
- **Markdig** - Markdown processing
- **iText7** - PDF generation
- **System.Text.Json** - JSON configuration

## 🔧 Customization

### Adding Custom Fields
1. Add columns to your Excel file
2. Use `{{ColumnName}}` in your templates
3. The system automatically processes custom fields

### Creating New Templates
1. Create new `.md` files in the `Templates/` folder
2. Add YAML metadata at the top
3. Use placeholder syntax for dynamic content
4. Update the code to load your new templates

### Styling PDFs
Modify the CSS in `PdfGeneratorService.cs` to change PDF appearance.

## 📋 Logging

All activities are logged to:
- **Console** - Real-time output
- **File** - `Output/docmailer.log`

Log levels: INFO, WARNING, ERROR

## ⚠️ Troubleshooting

### 🔍 Using Dry Run for Debugging
Before sending actual emails, always test with `--dry-run`:

```bash
# Test your setup without sending emails
dotnet run send-all --dry-run

# Verify configuration without sending test email
dotnet run test --dry-run
```

This will show you:
- Which recipients would be processed
- What PDFs would be generated
- What email subjects and content would be sent
- Any potential errors without affecting your data

### Common Issues

**Email not sending:**
- Check SMTP settings
- Verify email credentials
- Ensure "Less secure app access" is enabled (if using basic auth)
- Use App Passwords for Gmail/Outlook

**Excel file not found:**
- Verify file path in config.json
- Ensure file exists in Data/ folder
- Check file permissions

**PDF generation fails:**
- Check template syntax
- Verify Markdown formatting
- Review error logs

**Template placeholders not working:**
- Ensure placeholder names match Excel column headers
- Check for typos in placeholder syntax
- Verify YAML metadata format

**Markdown formatting not working in PDFs:**
- The system automatically trims whitespace from all fields
- If you have old data with trailing spaces, they will be cleaned automatically
- Bold formatting `**Name**` should work correctly in generated PDFs

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### How to Contribute
1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Setup
1. Clone your fork
2. Install dependencies: `dotnet restore`
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

For support and questions:

- **GitHub Issues**: [Create an issue](https://github.com/filipecarneiro/DocMailer/issues)
- **Documentation**: Check this README and inline code comments
- **Logs**: Review logs in `Output/docmailer.log` for troubleshooting

### Reporting Issues
When reporting issues, please include:
1. Operating system and .NET version
2. Error messages from logs
3. Steps to reproduce the problem
4. Sample data (anonymized) if relevant

## 🙋‍♂️ Author

**Filipe Carneiro**
- GitHub: [@filipecarneiro](https://github.com/filipecarneiro)
- Email: Available in commit history

---

**DocMailer** - Streamlining document generation and email automation.
