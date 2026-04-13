# Automated Excel Report Email Dispatcher

## Overview

Automated Excel Report Email Dispatcher is a portfolio-grade Python automation project designed to simulate a real freelance client workflow where Excel or CSV recipient instructions are used to distribute report files by email.

The script reads a structured recipient list, cleans and validates the data, matches each recipient to the correct report file, prepares email messages with attachments, and produces a professional dispatch register in Excel format.

This project is built to reflect practical business automation work for clients who need recurring report distribution processes for internal teams, customers, departments, branch managers, finance contacts, or operational stakeholders.

The solution includes structured logging, validation controls, dry-run protection, output reporting, and formatted Excel delivery logs so it looks like a real client-ready automation system rather than a beginner exercise.

---

## Features

- reads recipient instructions from CSV or Excel files
- supports flexible input naming for recipient files
- scans a reports folder and matches report files automatically
- standardizes and cleans messy input column names
- normalizes recipient names, email fields, and text fields
- validates required columns before processing
- validates email addresses for To, CC, and BCC fields
- filters inactive recipients
- removes fully empty rows
- removes duplicate distribution rows
- supports custom subject lines and custom email messages
- automatically generates default subject lines when none are provided
- automatically generates default email body text when none is provided
- attaches the correct report file for each recipient
- supports dry-run mode for safe testing
- supports real SMTP-based sending through environment variables
- logs all major workflow events to a log file
- creates a formatted Excel dispatch register
- creates a text summary file for each run
- uses a professional folder structure suitable for portfolio presentation

---

## Project Structure

```
scripts/
├── excel_report_email_dispatcher/
│   ├── input/
│   │   ├── recipient_distribution.xlsx
│   │   └── reports/
│   │       ├── north_region_sales.xlsx
│   │       ├── finance_summary.xlsx
│   │       └── operations_report.xlsx
│   ├── output/
│   │   ├── dispatch_register_YYYYMMDD_HHMMSS.xlsx
│   │   └── dispatch_summary_YYYYMMDD_HHMMSS.txt
│   ├── log/
│   │   └── report_dispatch_YYYYMMDD_HHMMSS.log
│   ├── README.md
│   └── report_email_dispatcher.py
```

---

## Business Use Case

A business often needs to send recurring Excel reports to multiple stakeholders.

Examples include:

- monthly sales reports sent to regional managers
- finance summaries sent to accounting contacts
- customer-specific reports sent to client email addresses
- operations reports sent to internal teams
- department-based Excel files distributed on a schedule

Instead of manually attaching files and composing emails one by one, this project automates the workflow and produces a formal register of what was prepared or sent.

---

## Input Requirements

The script expects a recipient instruction file inside the input folder.

Supported file names include:

- recipient_distribution.xlsx
- recipient_distribution.xls
- recipient_distribution.csv
- recipients.xlsx
- recipients.xls
- recipients.csv

The script also expects report files to exist inside:

input/reports/

---

## Required Columns

The recipient file must contain these required columns:

- recipient_name
- email_to
- report_file

These columns are optional but supported:

- company
- email_cc
- email_bcc
- subject
- message
- is_active

Column names are cleaned automatically by the script, so minor differences in capitalization or spacing are handled.

Examples:

- Recipient Name
- recipient name
- RECIPIENT_NAME

These will all be standardized during processing.

---

## Example Recipient File Layout

recipient_name,email_to,report_file,company,email_cc,email_bcc,subject,message,is_active
Alice Smith,alice@example.com,north_region_sales.xlsx,North Division,,,Monthly Sales Report,"Hi Alice, please find your monthly report attached.",yes
Bob Jones,bob@example.com,finance_summary.xlsx,Finance Team,manager@example.com,,,"Please review the attached finance summary.",yes
Carol White,carol@example.com,operations_report.xlsx,Operations,,,,"",yes

---

## How the Workflow Operates

1. the script creates the required folders if they do not already exist
2. the logging system is initialized
3. the recipient instruction file is located inside the input folder
4. the file is loaded into pandas from CSV or Excel
5. column names are standardized
6. empty rows are removed
7. required fields are cleaned and validated
8. inactive or invalid recipient rows are separated out
9. report files are scanned from the input/reports folder
10. each valid recipient row is matched to a report file
11. an email message is prepared for each valid matched row
12. the workflow runs in dry-run mode by default
13. dispatch results are written to an Excel register
14. a text summary file is created
15. a detailed execution log is written to the log folder

---

## Safety Design

This project is intentionally configured to run in dry-run mode by default.

That means:

- the script prepares the workflow
- validates inputs
- matches files
- builds dispatch results
- writes output files
- does not send real emails unless explicitly enabled

This is useful for safe testing, demonstrations, and portfolio presentation.

---

## Real Email Sending

Real email sending is only enabled when the correct SMTP environment variables are configured and sending is explicitly turned on.

The script checks for these environment variables:

- REPORT_SMTP_HOST
- REPORT_SMTP_PORT
- REPORT_SMTP_USER
- REPORT_SMTP_PASSWORD
- REPORT_SMTP_SENDER
- REPORT_ENABLE_SEND

To allow real sending, REPORT_ENABLE_SEND must be set to:

1

If this is not configured, the script remains in dry-run mode.

---

## Output Files

Each run creates structured output files inside the output folder.

### 1. Dispatch Register

An Excel file is created with a timestamped filename.

Example:

dispatch_register_20260413_143000.xlsx

This workbook includes separate sheets such as:

- dispatch_results
- valid_recipients
- invalid_recipients

The dispatch register is professionally formatted with:

- styled headers
- frozen top row
- bordered cells
- wrapped text
- auto-sized columns
- status highlighting for sent, dry-run, and failed records

### 2. Summary File

A text summary file is created for each run.

Example:

dispatch_summary_20260413_143000.txt

This summary includes:

- run timestamp
- base directory
- output file names
- processing mode
- valid record count
- invalid record count
- matched report count
- sent count
- dry-run success count
- failed count

### 3. Log File

A detailed execution log is written to the log folder.

Example:

report_dispatch_20260413_143000.log

The log captures:

- workflow start
- file loading
- data cleaning actions
- validation results
- file matching actions
- dispatch mode
- delivery status
- errors and exceptions

---

## Validation Rules

The script applies practical validation checks before dispatching.

These include:

- required column validation
- removal of fully empty rows
- blank critical field filtering
- duplicate row removal
- active or inactive recipient handling
- email format validation
- report file existence checks
- invalid row separation

Rows that fail validation are not silently ignored. They are preserved in the invalid_recipients output sheet so the user can review the problems.

---

## Matching Logic

The report_file value in the recipient file is matched against the report files found in input/reports.

The matching logic supports:

- exact filename match
- case-insensitive filename match
- filename stem matching if the extension is omitted

This makes the workflow more robust for real client input scenarios.

---

## Default Email Behavior

If the subject field is blank, the script generates a subject automatically.

Example pattern:

Company Name | Report Name

If no company is provided, it uses a fallback style such as:

Automated Report Delivery | Report Name

If the message field is blank, the script generates a default email body that includes:

- recipient name
- report filename
- a short professional message
- a closing signature from the automation system

---

## Running the Script

Open a terminal in the project folder and run:

python report_email_dispatcher.py

Make sure the input folder contains:

- one supported recipient instruction file
- a reports subfolder with report files to attach

---

## Example Setup Before Running

excel_report_email_dispatcher/
├── input/
│   ├── recipient_distribution.xlsx
│   └── reports/
│       ├── north_region_sales.xlsx
│       ├── finance_summary.xlsx
│       └── operations_report.xlsx
├── output/
├── log/
├── README.md
└── report_email_dispatcher.py

---

## SMTP Configuration for Live Sending

If you want to test real sending later, configure these environment variables in your system:

REPORT_SMTP_HOST
REPORT_SMTP_PORT
REPORT_SMTP_USER
REPORT_SMTP_PASSWORD
REPORT_SMTP_SENDER
REPORT_ENABLE_SEND

Example values might look like this:

REPORT_SMTP_HOST=smtp.office365.com
REPORT_SMTP_PORT=587
REPORT_SMTP_USER=your_email@example.com
REPORT_SMTP_PASSWORD=your_password
REPORT_SMTP_SENDER=your_email@example.com
REPORT_ENABLE_SEND=1

Only after this configuration will the script attempt live sending.

---

## Recommended Workflow for Portfolio Demonstration

For portfolio use, the recommended approach is:

1. prepare a realistic recipient file
2. place sample report files into input/reports
3. run the script in dry-run mode
4. show the output Excel register
5. show the summary file
6. show the log file
7. explain that live sending is intentionally disabled by default for safety

This makes the project look professional and practical while avoiding unnecessary risk during demos.

---

## Skills Demonstrated

This project demonstrates the kind of automation skills a freelance client may want in a reporting workflow.

Key skills shown include:

- Python project structuring
- pandas data cleaning
- Excel automation with openpyxl
- logging system design
- input validation
- file handling
- report distribution logic
- email workflow automation
- dry-run safety controls
- professional output reporting

---

## Ideal Freelance Use Cases

This type of project can be adapted for clients who need:

- recurring Excel report distribution
- customer-specific report delivery
- internal department report circulation
- branch-level performance file delivery
- finance and operations reporting automation
- reduction of manual email attachment work
- improved reporting consistency and traceability

---

## Notes

- the script does not require manual folder creation beyond the main project structure because required folders are created automatically
- the script is safe by default because it runs in dry-run mode unless real sending is explicitly enabled
- invalid rows are preserved in the output instead of being silently dropped
- the project is designed to be easy to extend for future client requirements such as scheduling, templated emails, or department-based filtering

---

## Conclusion

Automated Excel Report Email Dispatcher is a strong portfolio-grade Python automation project that simulates a real client reporting process from data intake through controlled email dispatch preparation and output logging.

It demonstrates practical freelance value by combining data cleaning, file matching, reporting, and delivery workflow automation into a structured and professional solution.
