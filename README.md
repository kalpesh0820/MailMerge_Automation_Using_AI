# MailMerge_Automation
# Excel VBA Customer Email Automation

## Project Overview
This project automates sending personalized customer emails using Microsoft Excel VBA and the classic Microsoft Outlook desktop application.

The automation reads customer details from Excel, checks required information and attachment paths, creates personalized emails, attaches the relevant PDF file, and submits the emails to Outlook for sending. It also records the processing status for each customer in Excel.

## Objectives
- Reduce repetitive manual email work.
- Send personalized emails to multiple customers.
- Attach customer-specific PDF files automatically.
- Validate customer information before processing.
- Track successful submission, skipped rows, and errors in Excel.

## Tools and Technologies
- **Microsoft Excel:** Stores customer details and attachment paths.
- **VBA (Visual Basic for Applications):** Automates the workflow.
- **Microsoft Outlook Desktop:** Creates and sends emails.
- **Outlook COM Automation:** Connects Excel VBA to the Outlook application.

## Excel Dataset Structure

| Column | Description |
|---|---|
| B | Customer Name |
| C | Customer Email ID |
| D | Pending Amount |
| E | PDF Attachment Path |
| F | Email Processing Status |

The column headers are in Row 1, and customer records start from Row 2.

## How the Automation Works

1. Reads customer records from the active Excel worksheet.
2. Asks the user to confirm before starting the email process.
3. Connects to the configured classic Outlook desktop application.
4. Validates the customer name, email address, pending amount, and attachment path.
5. Checks whether the attachment file exists.
6. Creates a personalized email for each valid customer.
7. Adds the customer's PDF attachment.
8. Submits the email to Outlook using VBA.
9. Updates Column F with the processing status.
10. Displays a summary of submitted emails, skipped rows, and errors.

## Sample Email

**Subject:** Pending Amount Details

Hello Customer Name,

Your Amount is Pending Rs. 5000.

Please check the Attachment File for more details.

The actual customer name and pending amount are populated dynamically from Excel.

## Validation and Error Handling
- Skips rows with missing required information.
- Skips rows when the attachment file cannot be found.
- Records row-level errors in Column F.
- Displays a summary after processing.
- Requires user confirmation before beginning the sending process.

## Key Learning Outcomes
- Practical use of Excel VBA for business process automation.
- Reading and validating Excel data programmatically.
- Automating Outlook email creation and attachment handling.
- Implementing error handling and status tracking.
- Reducing repetitive administrative tasks.

## Important Requirements
- Microsoft Excel with VBA support.
- Classic Microsoft Outlook desktop installed and configured.
- Valid customer email addresses.
- Correct local or accessible file paths for PDF attachments.

## Limitations
The macro submits emails to Outlook; the Excel status does not independently verify delivery to the recipient. Outlook connectivity, account configuration, and sending restrictions can affect actual sending.

## Future Enhancements
- Add a preview and approval step for the first email.
- Generate a detailed execution log with timestamps.
- Add duplicate-send prevention.
- Create a summary dashboard for email processing.
- Add configurable email templates and improved validation.

## Project Summary
This project demonstrates how Excel VBA and Outlook automation can be combined to automate a repetitive business communication process while improving consistency, validation, and traceability.
