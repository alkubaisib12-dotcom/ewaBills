# ewaBills - Automated Billing Workflow System

A production-ready Python 3 solution that automates the complete workflow for processing Outlook emails, extracting PDFs, splitting pages, extracting account numbers, merging bills per recipient, and sending consolidated invoices.

## Table of Contents

- [Features](#features)
- [System Overview](#system-overview)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [File Structure](#file-structure)
- [Logging](#logging)
- [Troubleshooting](#troubleshooting)
- [Security Considerations](#security-considerations)

---

## Features

- **Outlook Email Integration**: Connects to Outlook via IMAP to fetch emails and download PDF attachments
- **PDF Processing**: Automatically splits multi-page PDFs into single-page files
- **Account Number Extraction**: Uses regex-based text extraction to identify account numbers from each page
- **Smart Grouping**: Handles multiple accounts per recipient and merges all their bills into one PDF
- **Batch Email Sending**: Sends merged PDFs via SMTP with configurable batch sizes and delays
- **Comprehensive Logging**: CSV-based logging with detailed status tracking for each account
- **Retry Mechanism**: Dedicated script to retry failed email sends
- **Dry-Run Mode**: Test the entire workflow without sending actual emails
- **Idempotent Operations**: Unique filenames prevent data loss on multiple runs

---

## System Overview

### Workflow

```
┌─────────────────────────────────────────────────────────────────┐
│ 1. Connect to Outlook IMAP → Download PDF attachments          │
└─────────────────┬───────────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────────┐
│ 2. Split each PDF → Individual single-page PDF files           │
└─────────────────┬───────────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────────┐
│ 3. Extract account numbers from each page using regex          │
└─────────────────┬───────────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────────┐
│ 4. Load mapping file (account → guy name & email)              │
└─────────────────┬───────────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────────┐
│ 5. Group pages by recipient (supports multiple accounts/guy)   │
└─────────────────┬───────────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────────┐
│ 6. Merge all pages per recipient → Single PDF per guy          │
└─────────────────┬───────────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────────┐
│ 7. Send merged PDFs via SMTP (with batch control)              │
└─────────────────┬───────────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────────┐
│ 8. Write comprehensive CSV log file                            │
└─────────────────────────────────────────────────────────────────┘
```

---

## Requirements

- **Python**: 3.11 or higher (tested on 3.11+)
- **Operating System**: Linux, macOS, or Windows
- **Email Account**: Microsoft 365 / Outlook.com account with IMAP/SMTP access
- **Network**: Internet connection for email operations

---

## Installation

### Step 1: Clone or Download the Repository

```bash
cd /path/to/ewaBills
```

### Step 2: Create a Virtual Environment

```bash
python3 -m venv venv
```

### Step 3: Activate the Virtual Environment

**Linux/macOS:**
```bash
source venv/bin/activate
```

**Windows:**
```cmd
venv\Scripts\activate
```

### Step 4: Install Dependencies

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

### Step 5: Verify Installation

```bash
python extract_and_send.py --help
python retry_failed_emails.py --help
```

---

## Configuration

### 1. Environment Variables Setup

Copy the example environment file and configure it:

```bash
cp .env.example .env
```

Edit the `.env` file with your actual credentials:

```bash
nano .env  # or use your preferred text editor
```

**Critical settings to configure:**

```ini
# Email credentials
EMAIL_USER=your-email@outlook.com
EMAIL_PASSWORD=your-password-or-app-password

# IMAP settings (usually default is fine)
IMAP_SERVER=outlook.office365.com
IMAP_PORT=993

# SMTP settings (usually default is fine)
SMTP_SERVER=smtp.office365.com
SMTP_PORT=587

# Path to mapping file
MAPPING_FILE=./config/accounts_mapping.xlsx
```

### 2. Account Mapping File

The mapping file connects account numbers to recipients. It supports both CSV and Excel formats.

**Required columns:**
- `account_number`: The account number as it appears in PDFs
- `guy_name`: Recipient's full name
- `guy_email`: Recipient's email address

**Example (CSV format):**

```csv
account_number,guy_name,guy_email
100001,John Smith,john.smith@example.com
100002,Jane Doe,jane.doe@example.com
100003,John Smith,john.smith@example.com
```

**Note:** One person can have multiple account numbers (see rows 1 and 3 above). The system will merge all their bills into one PDF.

**Create your mapping file:**

```bash
# Option 1: Use CSV
cp config/accounts_mapping_example.csv config/accounts_mapping.csv
# Then edit config/accounts_mapping.csv with your data

# Option 2: Create Excel file
# Create an Excel file with the three columns above
# Save it as config/accounts_mapping.xlsx
```

### 3. Customize Account Number Regex (Optional)

If your PDFs have a different format for account numbers, update the `ACCOUNT_REGEX` in `.env`:

```ini
# Default pattern (matches "Account No: 123456" or "Account Number: 123456")
ACCOUNT_REGEX=(?:Account\s*(?:No\.?|Number)?\s*:?\s*)(\d+)

# Example: Match "Acct# 123456"
ACCOUNT_REGEX=(?:Acct#\s*)(\d+)

# Example: Match "Customer ID: ABC-123"
ACCOUNT_REGEX=(?:Customer\s*ID:\s*)([A-Z0-9-]+)
```

---

## Usage

### Main Script: `extract_and_send.py`

This is the primary script that performs the complete workflow.

#### Basic Usage

```bash
# Run the full workflow
python extract_and_send.py
```

#### Dry-Run Mode (Recommended for First Run)

Test the entire process without sending emails:

```bash
python extract_and_send.py --dry-run
```

This will:
- Download PDFs from your inbox
- Split and process them
- Extract account numbers
- Create merged PDFs
- Generate log files
- **NOT send any emails** (but log what would be sent)

#### Command-Line Options

```
usage: extract_and_send.py [-h] [--dry-run]

Extract bills from Outlook emails and send merged PDFs

options:
  -h, --help  show this help message and exit
  --dry-run   Process everything but do not send emails
```

### Retry Script: `retry_failed_emails.py`

If some emails failed to send, use this script to retry them.

#### Retry Latest Run

```bash
python retry_failed_emails.py
```

This automatically finds the most recent log file and retries all failed entries.

#### Retry Specific Log File

```bash
python retry_failed_emails.py --log logs/run_20251120_103000.csv
```

#### Dry-Run Retry

```bash
python retry_failed_emails.py --dry-run
```

#### Command-Line Options

```
usage: retry_failed_emails.py [-h] [--log LOG] [--dry-run]

Retry sending failed emails from log file

options:
  -h, --help  show this help message and exit
  --log LOG   Path to specific log file to process (default: most recent)
  --dry-run   Process everything but do not send emails
```

---

## File Structure

```
ewaBills/
├── extract_and_send.py          # Main processing script
├── retry_failed_emails.py       # Retry failed emails script
├── requirements.txt             # Python dependencies
├── .env                         # Your configuration (DO NOT commit!)
├── .env.example                 # Configuration template
├── README.md                    # This file
│
├── config/                      # Configuration files
│   ├── accounts_mapping.xlsx    # Your account mapping (or .csv)
│   └── accounts_mapping_example.csv  # Example mapping file
│
├── downloads_raw/               # Raw PDFs from emails (created automatically)
├── single_pages/                # Single-page PDFs (created automatically)
├── merged_by_guy/               # Merged PDFs per recipient (created automatically)
└── logs/                        # CSV log files (created automatically)
    ├── run_20251120_103000.csv
    └── run_20251120_110500_retry.csv
```

### Directory Details

| Directory | Purpose | Created By |
|-----------|---------|------------|
| `downloads_raw/` | Stores original PDF attachments from emails | Script (auto-created) |
| `single_pages/` | Stores individual single-page PDFs after splitting | Script (auto-created) |
| `merged_by_guy/` | Stores final merged PDFs (one per recipient) | Script (auto-created) |
| `logs/` | Stores CSV log files for each run | Script (auto-created) |
| `config/` | Stores mapping files | User (manual) |

---

## Logging

### Log File Format

Each run creates a timestamped CSV log file: `logs/run_YYYYMMDD_HHMMSS.csv`

**Log Columns:**

| Column | Description |
|--------|-------------|
| `date` | Date of processing (YYYY-MM-DD) |
| `time` | Time of processing (HH:MM:SS) |
| `status` | `Success`, `Failure`, or `Skipped` |
| `account_number` | Account number extracted from PDF |
| `destination_email` | Recipient email (blank if skipped) |
| `num_pages` | Number of pages for this account in merged PDF |
| `failed_pages` | List of page numbers that failed processing (if any) |
| `error_message` | Error details (added by retry script if failures occur) |

### Log Status Definitions

- **Success**: Account was mapped, PDF was merged, and email was sent successfully
- **Failure**: Account was mapped and PDF was merged, but email sending failed
- **Skipped**: Account number was found in PDF but not present in mapping file, OR text extraction failed

### Example Log Entries

```csv
date,time,status,account_number,destination_email,num_pages,failed_pages
2025-11-20,10:30:15,Success,100001,john.smith@example.com,3,
2025-11-20,10:30:16,Success,100003,john.smith@example.com,2,
2025-11-20,10:30:17,Failure,100002,jane.doe@example.com,1,
2025-11-20,10:30:18,Skipped,999999,,0,[1 2 3]
```

---

## Troubleshooting

### Common Issues

#### 1. Authentication Errors

**Error:** `Authentication failed` or `Login failed`

**Solutions:**
- Verify your email and password in `.env`
- For Microsoft 365 accounts with 2FA, use an **App Password** instead of your regular password
  - Go to https://account.microsoft.com/security
  - Click "Advanced security options" → "App passwords"
  - Generate a new app password and use it in `.env`

#### 2. No Emails Found

**Error:** `Found 0 emails to process`

**Solutions:**
- Check the `IMAP_FOLDER` setting (default is `INBOX`)
- If `PROCESS_UNREAD_ONLY=True`, make sure you have unread emails
- Try setting `PROCESS_UNREAD_ONLY=False` to process all emails

#### 3. Account Numbers Not Extracted

**Error:** Multiple pages show `Failed to extract account number`

**Solutions:**
- Open a sample PDF and look at how the account number appears
- Update the `ACCOUNT_REGEX` pattern in `.env` to match your format
- Test regex patterns using online tools like regex101.com
- Add more logging by setting `LOG_LEVEL=DEBUG` in `.env`

#### 4. Mapping File Errors

**Error:** `Missing required columns`

**Solutions:**
- Ensure your CSV/Excel has exactly these column names: `account_number`, `guy_name`, `guy_email`
- Column names are case-insensitive but must match
- Check for extra spaces in column headers

#### 5. PDF Processing Errors

**Error:** `Error reading PDF` or `Error splitting page`

**Solutions:**
- Some PDFs may be password-protected or corrupted
- Check the specific PDF file manually
- The script will skip problematic PDFs and log them

#### 6. Email Sending Failures

**Error:** `Error sending email`

**Solutions:**
- Check SMTP settings in `.env`
- Verify your email account allows SMTP access
- Check rate limits (adjust `BATCH_SIZE` and `BATCH_DELAY_SECONDS`)
- Use the retry script to resend failed emails

### Debug Mode

Enable detailed logging:

```bash
# Edit .env
LOG_LEVEL=DEBUG

# Run script
python extract_and_send.py --dry-run
```

This will show detailed information about each step.

---

## Security Considerations

### Protecting Credentials

1. **Never commit `.env` to version control**
   ```bash
   # Add to .gitignore
   echo ".env" >> .gitignore
   ```

2. **Use App Passwords** instead of your main email password (especially with 2FA enabled)

3. **Restrict file permissions** on `.env`:
   ```bash
   chmod 600 .env
   ```

### Email Security

- The system uses TLS/STARTTLS for encrypted email transmission
- Emails are sent via authenticated SMTP (no anonymous relay)
- Consider using a dedicated service account for automation

### PDF Data Privacy

- Downloaded PDFs and extracted data remain on your local system
- Ensure the server/machine running this script is secure
- Consider encrypting the directories containing PDFs if they contain sensitive data
- Regularly clean up old files from `downloads_raw/` and `single_pages/`

---

## Advanced Configuration

### Running as a Cron Job

To run the script automatically (e.g., daily at 9 AM):

```bash
# Edit crontab
crontab -e

# Add this line (adjust path to your installation)
0 9 * * * cd /path/to/ewaBills && /path/to/ewaBills/venv/bin/python extract_and_send.py >> /path/to/ewaBills/logs/cron.log 2>&1
```

### Processing Specific Email Subjects

Modify the `fetch_emails_and_download_pdfs()` function in `extract_and_send.py`:

```python
# Change search criteria
search_criteria = '(SUBJECT "invoice")'  # Only emails with "invoice" in subject
```

### Custom Email Templates

Edit the `EMAIL_BODY_TEMPLATE` in `.env`:

```ini
EMAIL_BODY_TEMPLATE=Dear {guy_name},

Attached are your bills for this period.

Account summary:
- Total accounts: [will be filled manually]
- Period: November 2025

Thank you for your business.

Best regards,
Billing Team
```

### Batch Processing Tuning

Adjust batch settings to avoid rate limits:

```ini
# Send 20 emails at a time
BATCH_SIZE=20

# Wait 2 minutes between batches
BATCH_DELAY_SECONDS=120
```

---

## Example Workflow

### First-Time Setup

```bash
# 1. Install
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 2. Configure
cp .env.example .env
nano .env  # Add your email credentials

# 3. Create mapping file
cp config/accounts_mapping_example.csv config/accounts_mapping.csv
nano config/accounts_mapping.csv  # Add your account mappings

# 4. Test with dry-run
python extract_and_send.py --dry-run

# 5. Check the generated files
ls -la downloads_raw/
ls -la single_pages/
ls -la merged_by_guy/
cat logs/run_*.csv

# 6. If everything looks good, run for real
python extract_and_send.py
```

### Regular Usage

```bash
# Activate environment
source venv/bin/activate

# Run the script
python extract_and_send.py

# Check logs
cat logs/run_*.csv | tail -20

# Retry any failures
python retry_failed_emails.py
```

---

## Support and Contribution

### Getting Help

1. Check the [Troubleshooting](#troubleshooting) section
2. Review log files in `logs/` directory
3. Run in debug mode with `LOG_LEVEL=DEBUG`
4. Check the generated CSV log for specific account issues

### Reporting Issues

When reporting issues, include:
- Error message from terminal
- Relevant log file entries
- Python version (`python --version`)
- Operating system

---

## License

This project is provided as-is for production use. Modify as needed for your specific requirements.

---

## Changelog

### Version 1.0.0 (2025-11-20)
- Initial production release
- Complete workflow automation
- IMAP email fetching
- PDF splitting and merging
- Account number extraction
- Batch email sending
- CSV logging
- Retry mechanism
- Dry-run mode

---

**Built with Python 3.11+ | Production-Ready | Enterprise-Grade Error Handling**
