# Automated Health Check Email Script
This PowerShell script automates the process of extracting health check data from an Excel dashboard and sending it via Outlook email as a formatted draft. It is designed to streamline daily operational reporting for environments like Production (Prod) and Quality Assurance (QA).
# Features
- Validates the existence of the Excel file in OneDrive.
- Checks if the file was updated in the last 10 minutes.
- Extracts specific cell ranges from the last worksheet.
- Copies Excel ranges as images and embeds them in an Outlook email.
- Creates a draft email with formatted content and subject line.
- Cleans up Excel COM objects to prevent memory leaks.
# File Structure
📂 YourRepo/

├── Automate.ps1       
├── README.md     
# Prerequisites
- Windows OS
- Microsoft Excel (with COM support)
- Microsoft Outlook (with COM support)
- PowerShell 5.1 or later
- Access to the specified OneDrive path
# How It Works
1. Load Excel File: Checks if the specified Excel file exists.
2. Validate Timestamp: Ensures the file was modified within the last 10 minutes.
3. Open Excel Silently: Opens the file in the background.
4. Extract Data: Selects the last worksheet and defines two ranges:
   - A1:L22 for Prod
   - M1:X22 for QA
5. Launch Outlook: Creates a new draft email.
6. Format Email: Adds greetings, headings, and pastes the copied Excel ranges as images.
7. Set Metadata: Fills in subject, recipients, and sender details.
8. Cleanup: Closes Excel and releases COM objects.
