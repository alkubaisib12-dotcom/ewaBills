#!/usr/bin/env python3
"""
extract_and_send.py

Production-ready script for automating Outlook email processing,
PDF splitting, account number extraction, merging, and bill sending.

Author: Senior Python Backend Engineer
Python Version: 3.11+
"""

import imaplib
import email
import os
import re
import csv
import time
import logging
import argparse
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Set
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib

import pandas as pd
from pypdf import PdfReader, PdfWriter
import pdfplumber
from dotenv import load_dotenv


# ==================== CONFIGURATION ====================

# Load environment variables
load_dotenv()

# IMAP Configuration
IMAP_SERVER = os.getenv("IMAP_SERVER", "outlook.office365.com")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
EMAIL_USER = os.getenv("EMAIL_USER", "")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "")
IMAP_FOLDER = os.getenv("IMAP_FOLDER", "INBOX")
PROCESS_UNREAD_ONLY = os.getenv("PROCESS_UNREAD_ONLY", "True").lower() == "true"

# SMTP Configuration
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
USE_TLS = os.getenv("USE_TLS", "True").lower() == "true"

# File Paths
MAPPING_FILE = os.getenv("MAPPING_FILE", "./config/accounts_mapping.xlsx")
DOWNLOADS_RAW_DIR = Path(os.getenv("DOWNLOADS_RAW_DIR", "./downloads_raw"))
SINGLE_PAGES_DIR = Path(os.getenv("SINGLE_PAGES_DIR", "./single_pages"))
MERGED_BY_GUY_DIR = Path(os.getenv("MERGED_BY_GUY_DIR", "./merged_by_guy"))
LOGS_DIR = Path(os.getenv("LOGS_DIR", "./logs"))

# Processing Configuration
ACCOUNT_REGEX = os.getenv("ACCOUNT_REGEX", r"(?:Account\s*(?:No\.?|Number)?\s*:?\s*)(\d+)")
BATCH_SIZE = int(os.getenv("BATCH_SIZE", "30"))
BATCH_DELAY_SECONDS = int(os.getenv("BATCH_DELAY_SECONDS", "60"))

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


# ==================== DATA STRUCTURES ====================

class SinglePageInfo:
    """Information about a single-page PDF."""
    def __init__(
        self,
        single_page_path: Path,
        source_raw_file: str,
        page_number: int,
        account_number: Optional[str] = None,
        extraction_failed: bool = False
    ):
        self.single_page_path = single_page_path
        self.source_raw_file = source_raw_file
        self.page_number = page_number
        self.account_number = account_number
        self.extraction_failed = extraction_failed


class AccountMapping:
    """Mapping between account numbers and guy information."""
    def __init__(self, mapping_df: pd.DataFrame):
        # Normalize column names
        mapping_df.columns = mapping_df.columns.str.strip().str.lower()

        # Build account_number -> (guy_name, guy_email) mapping
        self.account_to_guy: Dict[str, Tuple[str, str]] = {}
        for _, row in mapping_df.iterrows():
            account_num = str(row['account_number']).strip()
            guy_name = str(row['guy_name']).strip()
            guy_email = str(row['guy_email']).strip().lower()
            self.account_to_guy[account_num] = (guy_name, guy_email)

        # Build guy_email -> list of account_numbers mapping
        self.guy_to_accounts: Dict[str, List[str]] = {}
        for account_num, (guy_name, guy_email) in self.account_to_guy.items():
            if guy_email not in self.guy_to_accounts:
                self.guy_to_accounts[guy_email] = []
            self.guy_to_accounts[guy_email].append(account_num)

    def get_guy_info(self, account_number: str) -> Optional[Tuple[str, str]]:
        """Get (guy_name, guy_email) for an account number."""
        return self.account_to_guy.get(account_number)

    def get_accounts_for_guy(self, guy_email: str) -> List[str]:
        """Get all account numbers for a guy."""
        return self.guy_to_accounts.get(guy_email, [])


class LogEntry:
    """Entry for the CSV log file."""
    def __init__(
        self,
        date: str,
        time: str,
        status: str,
        account_number: str,
        destination_email: str = "",
        num_pages: int = 0,
        failed_pages: str = ""
    ):
        self.date = date
        self.time = time
        self.status = status
        self.account_number = account_number
        self.destination_email = destination_email
        self.num_pages = num_pages
        self.failed_pages = failed_pages

    def to_dict(self) -> Dict[str, any]:
        """Convert to dictionary for CSV writing."""
        return {
            'date': self.date,
            'time': self.time,
            'status': self.status,
            'account_number': self.account_number,
            'destination_email': self.destination_email,
            'num_pages': self.num_pages,
            'failed_pages': self.failed_pages
        }


# ==================== UTILITY FUNCTIONS ====================

def ensure_directories_exist():
    """Create all required directories if they don't exist."""
    for directory in [DOWNLOADS_RAW_DIR, SINGLE_PAGES_DIR, MERGED_BY_GUY_DIR, LOGS_DIR]:
        directory.mkdir(parents=True, exist_ok=True)
    logger.info("Ensured all directories exist")


def sanitize_filename(filename: str) -> str:
    """Sanitize filename to remove problematic characters."""
    # Remove or replace characters that are problematic in filenames
    sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
    sanitized = re.sub(r'\s+', '_', sanitized)
    return sanitized


def get_timestamp() -> str:
    """Get current timestamp in YYYYMMDD_HHMMSS format."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def get_current_date() -> str:
    """Get current date in YYYY-MM-DD format."""
    return datetime.now().strftime("%Y-%m-%d")


def get_current_time() -> str:
    """Get current time in HH:MM:SS format."""
    return datetime.now().strftime("%H:%M:%S")


# ==================== EMAIL FETCHING ====================

def fetch_emails_and_download_pdfs(dry_run: bool = False) -> List[Path]:
    """
    Connect to Outlook IMAP, fetch emails, and download PDF attachments.

    Args:
        dry_run: If True, don't mark emails as read

    Returns:
        List of paths to downloaded PDF files
    """
    logger.info(f"Connecting to IMAP server: {IMAP_SERVER}:{IMAP_PORT}")

    downloaded_pdfs: List[Path] = []

    try:
        # Connect to IMAP server
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASSWORD)
        logger.info(f"Successfully logged in as {EMAIL_USER}")

        # Select folder
        mail.select(IMAP_FOLDER)
        logger.info(f"Selected folder: {IMAP_FOLDER}")

        # Search for emails
        if PROCESS_UNREAD_ONLY:
            search_criteria = "UNSEEN"
            logger.info("Searching for unread emails only")
        else:
            search_criteria = "ALL"
            logger.info("Searching for all emails")

        status, message_ids = mail.search(None, search_criteria)

        if status != "OK":
            logger.error("Failed to search emails")
            return downloaded_pdfs

        message_id_list = message_ids[0].split()
        logger.info(f"Found {len(message_id_list)} emails to process")

        # Process each email
        for msg_num in message_id_list:
            try:
                status, msg_data = mail.fetch(msg_num, "(RFC822)")

                if status != "OK":
                    logger.warning(f"Failed to fetch email {msg_num}")
                    continue

                # Parse email
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                # Get message ID for unique filenames
                message_id = email_message.get("Message-ID", str(msg_num.decode()))
                message_id = sanitize_filename(message_id)

                # Process attachments
                attachment_index = 0
                for part in email_message.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue

                    filename = part.get_filename()

                    if filename:
                        # Decode filename if needed
                        decoded_filename = decode_header(filename)[0][0]
                        if isinstance(decoded_filename, bytes):
                            decoded_filename = decoded_filename.decode()

                        # Only process PDF files
                        if not decoded_filename.lower().endswith('.pdf'):
                            logger.debug(f"Skipping non-PDF attachment: {decoded_filename}")
                            continue

                        # Generate unique filename
                        timestamp = get_timestamp()
                        unique_filename = f"{timestamp}_{message_id}_{attachment_index}.pdf"
                        filepath = DOWNLOADS_RAW_DIR / unique_filename

                        # Save attachment
                        with open(filepath, 'wb') as f:
                            f.write(part.get_payload(decode=True))

                        downloaded_pdfs.append(filepath)
                        logger.info(f"Downloaded PDF: {filepath.name}")
                        attachment_index += 1

            except Exception as e:
                logger.error(f"Error processing email {msg_num}: {e}")
                continue

        mail.close()
        mail.logout()
        logger.info(f"Successfully downloaded {len(downloaded_pdfs)} PDF files")

    except Exception as e:
        logger.error(f"Error connecting to IMAP server: {e}")
        raise

    return downloaded_pdfs


# ==================== PDF SPLITTING ====================

def split_pdfs_to_single_pages(pdf_files: List[Path]) -> List[SinglePageInfo]:
    """
    Split each PDF into single-page PDFs.

    Args:
        pdf_files: List of PDF file paths to split

    Returns:
        List of SinglePageInfo objects
    """
    logger.info(f"Starting to split {len(pdf_files)} PDF files")

    single_pages: List[SinglePageInfo] = []

    for pdf_path in pdf_files:
        try:
            reader = PdfReader(str(pdf_path))
            num_pages = len(reader.pages)

            logger.info(f"Processing {pdf_path.name} with {num_pages} pages")

            base_name = pdf_path.stem

            for page_num in range(num_pages):
                try:
                    # Create writer for single page
                    writer = PdfWriter()
                    writer.add_page(reader.pages[page_num])

                    # Generate unique filename for single page
                    single_page_filename = f"{base_name}_page_{page_num + 1}.pdf"
                    single_page_path = SINGLE_PAGES_DIR / single_page_filename

                    # Write single page PDF
                    with open(single_page_path, 'wb') as f:
                        writer.write(f)

                    # Create SinglePageInfo object
                    page_info = SinglePageInfo(
                        single_page_path=single_page_path,
                        source_raw_file=pdf_path.name,
                        page_number=page_num + 1
                    )
                    single_pages.append(page_info)

                    logger.debug(f"Created single page: {single_page_filename}")

                except Exception as e:
                    logger.error(f"Error splitting page {page_num + 1} of {pdf_path.name}: {e}")
                    # Create entry for failed page
                    page_info = SinglePageInfo(
                        single_page_path=SINGLE_PAGES_DIR / f"{base_name}_page_{page_num + 1}_FAILED.pdf",
                        source_raw_file=pdf_path.name,
                        page_number=page_num + 1,
                        extraction_failed=True
                    )
                    single_pages.append(page_info)
                    continue

        except Exception as e:
            logger.error(f"Error reading PDF {pdf_path.name}: {e}")
            continue

    logger.info(f"Successfully split into {len(single_pages)} single-page PDFs")
    return single_pages


# ==================== ACCOUNT NUMBER EXTRACTION ====================

def extract_account_number(pdf_path: Path, regex_pattern: str) -> Optional[str]:
    """
    Extract account number from a single-page PDF.

    Args:
        pdf_path: Path to single-page PDF
        regex_pattern: Regex pattern to match account number

    Returns:
        Extracted account number or None
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                return None

            page = pdf.pages[0]
            text = page.extract_text()

            if not text:
                return None

            # Search for account number using regex
            pattern = re.compile(regex_pattern, re.IGNORECASE)
            match = pattern.search(text)

            if match:
                account_number = match.group(1).strip()
                return account_number

            return None

    except Exception as e:
        logger.error(f"Error extracting text from {pdf_path.name}: {e}")
        return None


def extract_account_numbers_from_pages(
    single_pages: List[SinglePageInfo],
    regex_pattern: str = ACCOUNT_REGEX
) -> List[SinglePageInfo]:
    """
    Extract account numbers from all single-page PDFs.

    Args:
        single_pages: List of SinglePageInfo objects
        regex_pattern: Regex pattern for account number extraction

    Returns:
        Updated list of SinglePageInfo objects with account numbers
    """
    logger.info(f"Extracting account numbers from {len(single_pages)} pages")

    successful_extractions = 0
    failed_extractions = 0

    for page_info in single_pages:
        if page_info.extraction_failed:
            failed_extractions += 1
            continue

        account_number = extract_account_number(page_info.single_page_path, regex_pattern)

        if account_number:
            page_info.account_number = account_number
            successful_extractions += 1
            logger.debug(f"Extracted account {account_number} from {page_info.single_page_path.name}")
        else:
            page_info.extraction_failed = True
            failed_extractions += 1
            logger.warning(f"Failed to extract account number from {page_info.single_page_path.name}")

    logger.info(f"Account extraction: {successful_extractions} successful, {failed_extractions} failed")
    return single_pages


# ==================== MAPPING AND GROUPING ====================

def load_account_mapping(mapping_file: str) -> AccountMapping:
    """
    Load account mapping from CSV or Excel file.

    Args:
        mapping_file: Path to mapping file

    Returns:
        AccountMapping object
    """
    logger.info(f"Loading account mapping from {mapping_file}")

    try:
        if mapping_file.endswith('.csv'):
            df = pd.read_csv(mapping_file)
        elif mapping_file.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(mapping_file)
        else:
            raise ValueError(f"Unsupported file format: {mapping_file}")

        # Validate required columns
        required_columns = {'account_number', 'guy_name', 'guy_email'}
        df_columns = set(df.columns.str.strip().str.lower())

        if not required_columns.issubset(df_columns):
            missing = required_columns - df_columns
            raise ValueError(f"Missing required columns: {missing}")

        mapping = AccountMapping(df)
        logger.info(f"Loaded {len(mapping.account_to_guy)} account mappings")
        logger.info(f"Found {len(mapping.guy_to_accounts)} unique recipients")

        return mapping

    except Exception as e:
        logger.error(f"Error loading mapping file: {e}")
        raise


def group_pages_by_guy(
    single_pages: List[SinglePageInfo],
    mapping: AccountMapping
) -> Dict[str, List[SinglePageInfo]]:
    """
    Group single pages by guy email (accounting for multiple accounts per guy).

    Args:
        single_pages: List of SinglePageInfo objects
        mapping: AccountMapping object

    Returns:
        Dictionary mapping guy_email to list of SinglePageInfo objects
    """
    logger.info("Grouping pages by recipient")

    guy_pages: Dict[str, List[SinglePageInfo]] = {}
    skipped_pages = 0

    for page_info in single_pages:
        # Skip pages without account number
        if not page_info.account_number or page_info.extraction_failed:
            skipped_pages += 1
            continue

        # Get guy info for this account
        guy_info = mapping.get_guy_info(page_info.account_number)

        if not guy_info:
            logger.warning(f"Account {page_info.account_number} not found in mapping")
            skipped_pages += 1
            continue

        guy_name, guy_email = guy_info

        if guy_email not in guy_pages:
            guy_pages[guy_email] = []

        guy_pages[guy_email].append(page_info)

    logger.info(f"Grouped pages for {len(guy_pages)} recipients, skipped {skipped_pages} pages")

    return guy_pages


# ==================== PDF MERGING ====================

def merge_pages_per_guy(
    guy_pages: Dict[str, List[SinglePageInfo]],
    mapping: AccountMapping
) -> Dict[str, Tuple[Path, str, int]]:
    """
    Merge all pages for each guy into a single PDF.

    Args:
        guy_pages: Dictionary mapping guy_email to list of SinglePageInfo
        mapping: AccountMapping object

    Returns:
        Dictionary mapping guy_email to (merged_pdf_path, guy_name, num_pages)
    """
    logger.info(f"Merging PDFs for {len(guy_pages)} recipients")

    merged_pdfs: Dict[str, Tuple[Path, str, int]] = {}

    for guy_email, pages in guy_pages.items():
        try:
            # Sort pages for deterministic ordering
            # Sort by account_number, then source_raw_file, then page_number
            sorted_pages = sorted(
                pages,
                key=lambda p: (p.account_number or "", p.source_raw_file, p.page_number)
            )

            # Get guy name from first page's account number
            first_account = sorted_pages[0].account_number
            guy_info = mapping.get_guy_info(first_account)
            guy_name = guy_info[0] if guy_info else "Unknown"

            # Create merged PDF
            merger = PdfWriter()

            for page_info in sorted_pages:
                try:
                    reader = PdfReader(str(page_info.single_page_path))
                    merger.add_page(reader.pages[0])
                except Exception as e:
                    logger.error(f"Error adding page {page_info.single_page_path.name}: {e}")
                    continue

            # Generate unique filename for merged PDF
            timestamp = get_timestamp()
            sanitized_name = sanitize_filename(guy_name)
            sanitized_email = sanitize_filename(guy_email)
            merged_filename = f"{sanitized_name}_{sanitized_email}_{timestamp}.pdf"
            merged_path = MERGED_BY_GUY_DIR / merged_filename

            # Write merged PDF
            with open(merged_path, 'wb') as f:
                merger.write(f)

            num_pages = len(sorted_pages)
            merged_pdfs[guy_email] = (merged_path, guy_name, num_pages)

            logger.info(f"Created merged PDF for {guy_name} ({guy_email}): {num_pages} pages")

        except Exception as e:
            logger.error(f"Error merging PDFs for {guy_email}: {e}")
            continue

    logger.info(f"Successfully created {len(merged_pdfs)} merged PDFs")
    return merged_pdfs


# ==================== EMAIL SENDING ====================

def send_email_with_attachment(
    to_email: str,
    guy_name: str,
    attachment_path: Path,
    subject: str = EMAIL_SUBJECT,
    body_template: str = EMAIL_BODY_TEMPLATE
) -> bool:
    """
    Send email with PDF attachment via SMTP.

    Args:
        to_email: Recipient email address
        guy_name: Recipient name
        attachment_path: Path to PDF attachment
        subject: Email subject
        body_template: Email body template

    Returns:
        True if successful, False otherwise
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
        return True

    except Exception as e:
        logger.error(f"Error sending email to {to_email}: {e}")
        return False


def send_emails_for_merged_pdfs(
    merged_pdfs: Dict[str, Tuple[Path, str, int]],
    dry_run: bool = False,
    batch_size: int = BATCH_SIZE,
    batch_delay: int = BATCH_DELAY_SECONDS
) -> Dict[str, bool]:
    """
    Send emails with merged PDFs to recipients.

    Args:
        merged_pdfs: Dictionary mapping guy_email to (merged_pdf_path, guy_name, num_pages)
        dry_run: If True, don't actually send emails
        batch_size: Number of emails to send per batch
        batch_delay: Seconds to wait between batches

    Returns:
        Dictionary mapping guy_email to success status
    """
    logger.info(f"Sending emails to {len(merged_pdfs)} recipients (dry_run={dry_run})")

    send_results: Dict[str, bool] = {}
    emails_sent = 0

    for guy_email, (pdf_path, guy_name, num_pages) in merged_pdfs.items():
        if dry_run:
            logger.info(f"[DRY RUN] Would send email to {guy_email} with attachment {pdf_path.name}")
            send_results[guy_email] = True
            emails_sent += 1
        else:
            success = send_email_with_attachment(guy_email, guy_name, pdf_path)
            send_results[guy_email] = success

            if success:
                emails_sent += 1

            # Batch delay
            if emails_sent > 0 and emails_sent % batch_size == 0:
                logger.info(f"Sent {emails_sent} emails, pausing for {batch_delay} seconds...")
                time.sleep(batch_delay)

    successful = sum(1 for success in send_results.values() if success)
    failed = len(send_results) - successful

    logger.info(f"Email sending complete: {successful} successful, {failed} failed")
    return send_results


# ==================== LOGGING ====================

def write_run_log(
    single_pages: List[SinglePageInfo],
    mapping: AccountMapping,
    guy_pages: Dict[str, List[SinglePageInfo]],
    merged_pdfs: Dict[str, Tuple[Path, str, int]],
    send_results: Dict[str, bool]
) -> Path:
    """
    Write CSV log file for this run.

    Args:
        single_pages: All single page information
        mapping: Account mapping
        guy_pages: Pages grouped by guy
        merged_pdfs: Merged PDF information
        send_results: Email sending results

    Returns:
        Path to log file
    """
    timestamp = get_timestamp()
    log_filename = f"run_{timestamp}.csv"
    log_path = LOGS_DIR / log_filename

    logger.info(f"Writing run log to {log_path}")

    log_entries: List[LogEntry] = []
    processed_accounts: Set[str] = set()

    # Process successfully sent emails
    for guy_email, (pdf_path, guy_name, num_pages) in merged_pdfs.items():
        # Get all accounts for this guy
        accounts = mapping.get_accounts_for_guy(guy_email)

        # Find pages for this guy to count per account
        pages_for_guy = guy_pages.get(guy_email, [])

        # Group by account number
        account_page_counts: Dict[str, int] = {}
        for page in pages_for_guy:
            if page.account_number:
                account_page_counts[page.account_number] = account_page_counts.get(page.account_number, 0) + 1

        # Create log entry for each account
        for account_num in account_page_counts.keys():
            processed_accounts.add(account_num)

            success = send_results.get(guy_email, False)
            status = "Success" if success else "Failure"

            entry = LogEntry(
                date=get_current_date(),
                time=get_current_time(),
                status=status,
                account_number=account_num,
                destination_email=guy_email,
                num_pages=account_page_counts[account_num],
                failed_pages=""
            )
            log_entries.append(entry)

    # Process skipped pages (no mapping found)
    account_skip_reasons: Dict[str, List[int]] = {}

    for page in single_pages:
        if page.account_number and page.account_number not in processed_accounts:
            # Account number found but not in mapping
            guy_info = mapping.get_guy_info(page.account_number)
            if not guy_info:
                if page.account_number not in account_skip_reasons:
                    account_skip_reasons[page.account_number] = []
                account_skip_reasons[page.account_number].append(page.page_number)

    for account_num, failed_page_nums in account_skip_reasons.items():
        entry = LogEntry(
            date=get_current_date(),
            time=get_current_time(),
            status="Skipped",
            account_number=account_num,
            destination_email="",
            num_pages=0,
            failed_pages=str(failed_page_nums)
        )
        log_entries.append(entry)

    # Process pages with failed extraction (no account number found)
    failed_extraction_pages = [p for p in single_pages if p.extraction_failed or not p.account_number]
    if failed_extraction_pages:
        failed_page_nums = [p.page_number for p in failed_extraction_pages]
        entry = LogEntry(
            date=get_current_date(),
            time=get_current_time(),
            status="Skipped",
            account_number="",
            destination_email="",
            num_pages=0,
            failed_pages=str(failed_page_nums)
        )
        log_entries.append(entry)

    # Write to CSV
    with open(log_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['date', 'time', 'status', 'account_number', 'destination_email', 'num_pages', 'failed_pages']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        for entry in log_entries:
            writer.writerow(entry.to_dict())

    logger.info(f"Successfully wrote {len(log_entries)} log entries to {log_path}")
    return log_path


# ==================== MAIN ====================

def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description="Extract bills from Outlook emails and send merged PDFs"
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Process everything but do not send emails'
    )

    args = parser.parse_args()

    logger.info("=" * 80)
    logger.info("Starting extract_and_send.py")
    logger.info(f"Dry run mode: {args.dry_run}")
    logger.info("=" * 80)

    try:
        # Step 0: Ensure directories exist
        ensure_directories_exist()

        # Step 1: Fetch emails and download PDFs
        logger.info("\n[STEP 1] Fetching emails and downloading PDFs...")
        downloaded_pdfs = fetch_emails_and_download_pdfs(dry_run=args.dry_run)

        if not downloaded_pdfs:
            logger.warning("No PDF attachments found in emails")
            return

        # Step 2: Split PDFs into single pages
        logger.info("\n[STEP 2] Splitting PDFs into single pages...")
        single_pages = split_pdfs_to_single_pages(downloaded_pdfs)

        if not single_pages:
            logger.warning("No pages extracted from PDFs")
            return

        # Step 3: Extract account numbers
        logger.info("\n[STEP 3] Extracting account numbers from pages...")
        single_pages = extract_account_numbers_from_pages(single_pages)

        # Step 4: Load account mapping
        logger.info("\n[STEP 4] Loading account mapping...")
        mapping = load_account_mapping(MAPPING_FILE)

        # Step 5: Group pages by guy
        logger.info("\n[STEP 5] Grouping pages by recipient...")
        guy_pages = group_pages_by_guy(single_pages, mapping)

        if not guy_pages:
            logger.warning("No pages matched to recipients")
            return

        # Step 6: Merge pages per guy
        logger.info("\n[STEP 6] Merging PDFs per recipient...")
        merged_pdfs = merge_pages_per_guy(guy_pages, mapping)

        if not merged_pdfs:
            logger.warning("No merged PDFs created")
            return

        # Step 7: Send emails
        logger.info("\n[STEP 7] Sending emails...")
        send_results = send_emails_for_merged_pdfs(merged_pdfs, dry_run=args.dry_run)

        # Step 8: Write log
        logger.info("\n[STEP 8] Writing run log...")
        log_path = write_run_log(single_pages, mapping, guy_pages, merged_pdfs, send_results)

        logger.info("=" * 80)
        logger.info("Processing complete!")
        logger.info(f"Log file: {log_path}")
        logger.info("=" * 80)

    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
