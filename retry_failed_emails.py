#!/usr/bin/env python3
"""
retry_failed_emails.py

Script to retry sending emails for failed entries in the log file.

Author: Senior Python Backend Engineer
Python Version: 3.11+
"""

import csv
import os
import argparse
import logging
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import pandas as pd
from dotenv import load_dotenv


# ==================== CONFIGURATION ====================

# Load environment variables
load_dotenv()

# SMTP Configuration
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
USE_TLS = os.getenv("USE_TLS", "True").lower() == "true"
EMAIL_USER = os.getenv("EMAIL_USER", "")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")

# File Paths
MERGED_BY_GUY_DIR = Path(os.getenv("MERGED_BY_GUY_DIR", "./merged_by_guy"))
LOGS_DIR = Path(os.getenv("LOGS_DIR", "./logs"))
MAPPING_FILE = os.getenv("MAPPING_FILE", "./config/accounts_mapping.xlsx")

# Email Template
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT", "Your Monthly Bills")
EMAIL_BODY_TEMPLATE = os.getenv("EMAIL_BODY_TEMPLATE", """Dear {guy_name},

Please find attached your monthly bills.

This is an automated message. Please do not reply to this email.

Best regards,
Billing Department
""")

# Logging
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL),
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# ==================== UTILITY FUNCTIONS ====================

def get_current_date() -> str:
    """Get current date in YYYY-MM-DD format."""
    return datetime.now().strftime("%Y-%m-%d")


def get_current_time() -> str:
    """Get current time in HH:MM:SS format."""
    return datetime.now().strftime("%H:%M:%S")


def find_latest_log_file() -> Optional[Path]:
    """Find the most recent log file in the logs directory."""
    if not LOGS_DIR.exists():
        logger.error(f"Logs directory does not exist: {LOGS_DIR}")
        return None

    log_files = list(LOGS_DIR.glob("run_*.csv"))

    if not log_files:
        logger.error("No log files found in logs directory")
        return None

    # Sort by modification time, most recent first
    latest_log = max(log_files, key=lambda p: p.stat().st_mtime)
    return latest_log


def load_mapping_file(mapping_file: str) -> Dict[str, Tuple[str, str]]:
    """
    Load account mapping from CSV or Excel file.

    Args:
        mapping_file: Path to mapping file

    Returns:
        Dictionary mapping account_number to (guy_name, guy_email)
    """
    logger.info(f"Loading account mapping from {mapping_file}")

    try:
        if mapping_file.endswith('.csv'):
            df = pd.read_csv(mapping_file)
        elif mapping_file.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(mapping_file)
        else:
            raise ValueError(f"Unsupported file format: {mapping_file}")

        # Normalize column names
        df.columns = df.columns.str.strip().str.lower()

        # Build account_number -> (guy_name, guy_email) mapping
        account_to_guy: Dict[str, Tuple[str, str]] = {}
        guy_to_email: Dict[str, str] = {}  # guy_email -> guy_name

        for _, row in df.iterrows():
            account_num = str(row['account_number']).strip()
            guy_name = str(row['guy_name']).strip()
            guy_email = str(row['guy_email']).strip().lower()
            account_to_guy[account_num] = (guy_name, guy_email)
            guy_to_email[guy_email] = guy_name

        logger.info(f"Loaded {len(account_to_guy)} account mappings")
        return account_to_guy, guy_to_email

    except Exception as e:
        logger.error(f"Error loading mapping file: {e}")
        raise


def find_merged_pdf_for_email(guy_email: str) -> Optional[Path]:
    """
    Find the most recent merged PDF for a guy's email.

    Args:
        guy_email: Email address to search for

    Returns:
        Path to merged PDF or None
    """
    if not MERGED_BY_GUY_DIR.exists():
        logger.error(f"Merged PDFs directory does not exist: {MERGED_BY_GUY_DIR}")
        return None

    # Search for PDFs containing this email in the filename
    sanitized_email = guy_email.replace('@', '_').replace('.', '_')
    pdf_files = list(MERGED_BY_GUY_DIR.glob(f"*{sanitized_email}*.pdf"))

    if not pdf_files:
        # Try a more flexible search
        pdf_files = list(MERGED_BY_GUY_DIR.glob("*.pdf"))
        matching_files = [f for f in pdf_files if sanitized_email in f.name or guy_email in f.name]

        if not matching_files:
            logger.warning(f"No merged PDF found for {guy_email}")
            return None

        pdf_files = matching_files

    # Sort by modification time, most recent first
    latest_pdf = max(pdf_files, key=lambda p: p.stat().st_mtime)
    return latest_pdf


def send_email_with_attachment(
    to_email: str,
    guy_name: str,
    attachment_path: Path,
    subject: str = EMAIL_SUBJECT,
    body_template: str = EMAIL_BODY_TEMPLATE
) -> Tuple[bool, str]:
    """
    Send email with PDF attachment via SMTP.

    Args:
        to_email: Recipient email address
        guy_name: Recipient name
        attachment_path: Path to PDF attachment
        subject: Email subject
        body_template: Email body template

    Returns:
        Tuple of (success: bool, error_message: str)
    """
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = to_email
        msg['Subject'] = subject

        # Format body
        body = body_template.format(guy_name=guy_name)
        msg.attach(MIMEText(body, 'plain'))

        # Attach PDF
        with open(attachment_path, 'rb') as f:
            pdf_attachment = MIMEApplication(f.read(), _subtype='pdf')
            pdf_attachment.add_header('Content-Disposition', 'attachment', filename='bills.pdf')
            msg.attach(pdf_attachment)

        # Connect to SMTP server and send
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            if USE_TLS:
                server.starttls()
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.send_message(msg)

        logger.info(f"Successfully sent email to {to_email}")
        return True, ""

    except Exception as e:
        error_msg = str(e)
        logger.error(f"Error sending email to {to_email}: {error_msg}")
        return False, error_msg


# ==================== RETRY LOGIC ====================

def retry_failed_emails(log_file: Path, dry_run: bool = False) -> Path:
    """
    Retry sending emails for failed entries in the log file.

    Args:
        log_file: Path to the log file to process
        dry_run: If True, don't actually send emails

    Returns:
        Path to updated log file
    """
    logger.info(f"Processing log file: {log_file}")

    # Read the log file
    try:
        with open(log_file, 'r', newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            log_entries = list(reader)
    except Exception as e:
        logger.error(f"Error reading log file: {e}")
        raise

    logger.info(f"Found {len(log_entries)} entries in log file")

    # Load mapping to get guy names
    account_to_guy, guy_to_email = load_mapping_file(MAPPING_FILE)

    # Find failed entries and group by destination_email
    failed_by_email: Dict[str, List[Dict]] = {}

    for entry in log_entries:
        if entry['status'] == 'Failure' and entry['destination_email']:
            email = entry['destination_email']
            if email not in failed_by_email:
                failed_by_email[email] = []
            failed_by_email[email].append(entry)

    logger.info(f"Found {len(failed_by_email)} unique recipients with failures")

    if not failed_by_email:
        logger.info("No failed emails to retry")
        return log_file

    # Retry sending emails
    retry_results: Dict[str, Tuple[bool, str]] = {}

    for guy_email, entries in failed_by_email.items():
        logger.info(f"Retrying email for {guy_email}")

        # Get guy name
        guy_name = guy_to_email.get(guy_email, "Valued Customer")

        # Find merged PDF
        pdf_path = find_merged_pdf_for_email(guy_email)

        if not pdf_path:
            error_msg = f"Merged PDF not found for {guy_email}"
            logger.error(error_msg)
            retry_results[guy_email] = (False, error_msg)
            continue

        if dry_run:
            logger.info(f"[DRY RUN] Would retry sending to {guy_email} with {pdf_path.name}")
            retry_results[guy_email] = (True, "")
        else:
            success, error_msg = send_email_with_attachment(guy_email, guy_name, pdf_path)
            retry_results[guy_email] = (success, error_msg)

    # Update log entries
    for entry in log_entries:
        if entry['status'] == 'Failure' and entry['destination_email']:
            email = entry['destination_email']
            if email in retry_results:
                success, error_msg = retry_results[email]
                if success:
                    entry['status'] = 'Success'
                    entry['date'] = get_current_date()
                    entry['time'] = get_current_time()
                else:
                    # Add error_message column if not exists
                    if 'error_message' not in entry:
                        entry['error_message'] = error_msg

    # Write updated log file
    # Add error_message to fieldnames if any failures occurred
    fieldnames = list(log_entries[0].keys())
    if 'error_message' not in fieldnames:
        fieldnames.append('error_message')

    # Ensure all entries have all fields
    for entry in log_entries:
        for field in fieldnames:
            if field not in entry:
                entry[field] = ''

    # Create a new log file with _retry suffix
    retry_log_path = log_file.parent / f"{log_file.stem}_retry{log_file.suffix}"

    with open(retry_log_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(log_entries)

    logger.info(f"Updated log file written to: {retry_log_path}")

    # Report results
    successful_retries = sum(1 for success, _ in retry_results.values() if success)
    failed_retries = len(retry_results) - successful_retries

    logger.info("=" * 80)
    logger.info(f"Retry complete: {successful_retries} successful, {failed_retries} failed")
    logger.info("=" * 80)

    return retry_log_path


# ==================== MAIN ====================

def main():
    """Main entry point for the retry script."""
    parser = argparse.ArgumentParser(
        description="Retry sending failed emails from log file"
    )
    parser.add_argument(
        '--log',
        type=str,
        help='Path to specific log file to process (default: most recent)'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Process everything but do not send emails'
    )

    args = parser.parse_args()

    logger.info("=" * 80)
    logger.info("Starting retry_failed_emails.py")
    logger.info(f"Dry run mode: {args.dry_run}")
    logger.info("=" * 80)

    try:
        # Determine which log file to process
        if args.log:
            log_file = Path(args.log)
            if not log_file.exists():
                logger.error(f"Log file not found: {log_file}")
                return
        else:
            log_file = find_latest_log_file()
            if not log_file:
                logger.error("No log file found")
                return

        logger.info(f"Using log file: {log_file}")

        # Retry failed emails
        updated_log = retry_failed_emails(log_file, dry_run=args.dry_run)

        logger.info("=" * 80)
        logger.info("Retry processing complete!")
        logger.info(f"Updated log file: {updated_log}")
        logger.info("=" * 80)

    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
