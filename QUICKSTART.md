# Quick Start Guide

Get the ewaBills system running in 5 minutes.

## Installation

```bash
# 1. Activate virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt
```

## Configuration

```bash
# 3. Set up environment
cp .env.example .env
nano .env  # Edit with your credentials
```

**Minimum required settings in `.env`:**
```ini
EMAIL_USER=your-email@outlook.com
EMAIL_PASSWORD=your-app-password
MAPPING_FILE=./config/accounts_mapping.csv
```

```bash
# 4. Create mapping file
cp config/accounts_mapping_example.csv config/accounts_mapping.csv
nano config/accounts_mapping.csv  # Add your account numbers, names, and emails
```

## First Run (Dry-Run)

```bash
# Test without sending emails
python extract_and_send.py --dry-run
```

**Check the output:**
- `downloads_raw/` - Downloaded PDFs
- `single_pages/` - Split pages
- `merged_by_guy/` - Merged PDFs per recipient
- `logs/run_*.csv` - Processing log

## Production Run

```bash
# Send emails for real
python extract_and_send.py
```

## Retry Failed Emails

```bash
# If some emails failed, retry them
python retry_failed_emails.py
```

## Common Issues

### App Password Required
If you have 2FA enabled on Outlook:
1. Go to https://account.microsoft.com/security
2. Create an App Password
3. Use it in `.env` instead of your regular password

### No PDFs Found
- Check `IMAP_FOLDER` setting (default: INBOX)
- Set `PROCESS_UNREAD_ONLY=False` to process all emails

### Account Numbers Not Extracted
- Update `ACCOUNT_REGEX` in `.env` to match your PDF format
- Set `LOG_LEVEL=DEBUG` to see detailed extraction info

## Next Steps

Read the full [README.md](README.md) for:
- Detailed configuration options
- Advanced features
- Troubleshooting guide
- Security best practices
