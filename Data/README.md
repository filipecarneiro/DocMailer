# Sample Excel File

To test the system, create an Excel file named `recipients.xlsx` in the `Data/` folder with the following structure:

| Name | Email | Company | Position | LastSent | Responded | Extra_Field |
|------|-------|---------|----------|----------|-----------|-------------|
| John Smith | john@company.com | ABC Company | Director |  |  | Value1 |
| Mary Johnson | mary@company.com | XYZ Company | Manager |  |  | Value2 |
| Peter Brown | peter@company.com | 123 Company | Developer |  |  | Value3 |
| Test User | test@company.com | Test Corp | Tester |  |  | TestValue |

## Instructions:
1. Open Excel
2. Create a new spreadsheet
3. Add the headers in the first row (A1:G1)
4. Add the recipient data in the following rows
5. Save as `recipients.xlsx` in the `Data/` folder

## Required Fields:
- **Name** (column A): Full name of the recipient
- **Email** (column B): Valid email address

## Optional Fields:
- **Company** (column C): Company name
- **Position** (column D): Position/role in the company
- **LastSent** (column E): Auto-updated with timestamp when email is sent successfully, or error message if failed
- **Responded** (column F): Manual update (true/false) to track if recipient responded
- **Extra Fields** (column G+): Any additional field you want to include in the templates

## Command Line Usage:
- `dotnet run send-all` - Send to all recipients
- `dotnet run send-not-sent` - Send only to recipients without LastSent timestamp
- `dotnet run send-not-responded` - Send only to recipients who received email but haven't responded
- `dotnet run send-test` - Send only to recipients with "test" in name or email

## Notes:
- The **LastSent** column is automatically managed by the system
- You need to manually update the **Responded** column (true/false) based on actual responses
- Test recipients should have "test" in their name or email for the send-test command
