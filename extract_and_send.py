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
ACCOUNT_REGEX = os.getenv(
    "ACCOUNT_REGEX",
    r"(?:Account\s*(?:No\.?|Number)?\s*:?\s*)(\d+)",
)
BATCH_SIZE = int(os.getenv("BATCH_SIZE", "30"))
BATCH_DELAY_SECONDS = int(os.getenv("BATCH_DELAY_SECONDS", "60"))

# Email Template
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT", "Your Monthly Bills")
EMAIL_BODY_TEMPLATE = os.getenv(
    "EMAIL_BODY_TEMPLATE",
    """Dear {guy_name},

Please find attached your monthly bills.

This is an automated message. Please do not reply to this email.

Best regards,
Billing Department
""",
)

# Logging
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL),
    format="%(asctime)s - %(levelname)s - %(message)s",
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
        extraction_failed: bool = False,
    ) -> None:
        self.single_page_path = single_page_path
        self.source_raw_file = source_raw_file
        self.page_number = page_number
        self.account_number = account_number
        self.extraction_failed = extraction_failed


class AccountMapping:
    """Mapping between account numbers and guy information."""

    def __init__(self, mapping_df: pd.DataFrame) -> None:
        # Normalize column names
        mapping_df.columns = mapping_df.columns.str.strip().str.lower()

        # Build account_number -> (guy_name, guy_email) mapping
        self.account_to_guy: Dict[str, Tuple[str, str]] = {}
        for _, row in mapping_df.iterrows():
            account_num = str(row["account_number"]).strip()
            guy_name = str(row["guy_name"]).strip()
            guy_email = str(row["guy_email"]).strip().lower()
            self.account_to_guy[account_num] = (guy_name, guy_email)

        # Build guy_email -> list of account_numbers mapping
        self.guy_to_accounts: Dict[str, List[str]] = {}
        for account_num, (_, guy_email) in self.account_to_guy.items():
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
        failed_pages: str = "",
    ) -> None:
        self.date = date
        self.time = time
        self.status = status
        self.account_number = account_number
        self.destination_email = destination_email
        self.num_pages = num_pages
        self.failed_pages = failed_pages

    def to_dict(self) -> Dict[str, object]:
        """Convert to dictionary for CSV writing."""

        return {
            "date": self.date,
            "time": self.time,
            "status": self.status,
            "account_number": self.account_number,
            "destination_email": self.destination_email,
            "num_pages": self.num_pages,
            "failed_pages": self.failed_pages,
        }


# ==================== UTILITY FUNCTIONS ====================


def ensure_directories_exist() -> None:
    """Create all required directories if they don't exist."""

    for directory in [
        DOWNLOADS_RAW_DIR,
        SINGLE_PAGES_DIR,
        MERGED_BY_GUY_DIR,
        LOGS_DIR,
    ]:
        directory.mkdir(parents=True, exist_ok=True)
    logger.info("Ensured all directories exist")


def sanitize_filename(filename: str) -> str:
    """Sanitize filename to remove problematic characters."""

    sanitized = re.sub(r'[<>:"/\\|?*]', "_", filename)
    sanitized = re.sub(r"\s+", "_", sanitized)
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


# ==================== CONSOLE OUTPUT HELPERS ====================


def print_header(text: str) -> None:
    """Print a formatted header."""

    print("\n" + "=" * 80)
    print(f"  {text}")
    print("=" * 80)


def print_section(text: str) -> None:
    """Print a section header."""

    print(f"\n{'─' * 80}")
    print(f"▶ {text}")
    print("─" * 80)


def print_success(text: str) -> None:
    """Print success message with checkmark."""

    print(f"  ✓ {text}")


def print_error(text: str) -> None:
    """Print error message with X."""

    print(f"  ✗ {text}")


def print_warning(text: str) -> None:
    """Print warning message."""

    print(f"  ⚠ {text}")


def print_info(text: str) -> None:
    """Print info message."""

    print(f"  → {text}")


def print_progress(current: int, total: int, text: str) -> None:
    """Print progress indicator."""

    print(f"  [{current}/{total}] {text}")


def print_summary_box(title: str, items: Dict[str, object]) -> None:
    """Print a summary box."""

    print(f"\n╔{'═' * 78}╗")
    print(f"║ {title.center(76)} ║")
    print(f"╠{'═' * 78}╣")
    for key, value in items.items():
        line = f"║  {key}: {value}"
        print(line + " " * (79 - len(line)) + "║")
    print(f"╚{'═' * 78}╝")


# ==================== EMAIL FETCHING ====================


def fetch_emails_and_download_pdfs(dry_run: bool = False) -> List[Path]:
    """Connect to IMAP, fetch emails, and download PDF attachments."""

    print_info(f"Connecting to IMAP: {IMAP_SERVER}")
    logger.info(f"Connecting to IMAP server: {IMAP_SERVER}:{IMAP_PORT}")

    downloaded_pdfs: List[Path] = []

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASSWORD)
        print_success(f"Logged in as {EMAIL_USER}")

        mail.select(IMAP_FOLDER)
        logger.info(f"Selected folder: {IMAP_FOLDER}")

        if PROCESS_UNREAD_ONLY:
            search_criteria = "UNSEEN"
            print_info("Searching for unread emails with PDF attachments...")
        else:
            search_criteria = "ALL"
            print_info("Searching all emails for PDF attachments...")

        logger.info(f"Search criteria: {search_criteria}")
        status, message_ids = mail.search(None, search_criteria)

        if status != "OK":
            print_error("Failed to search emails")
            logger.error("Failed to search emails")
            return downloaded_pdfs

        message_id_list = message_ids[0].split()
        total_emails = len(message_id_list)
        print_success(f"Found {total_emails} emails")
        logger.info(f"Found {total_emails} emails to process")

        email_counter = 0
        pdfs_found = 0

        print_info(
            f"Processing {total_emails} emails (showing only emails with PDFs)...",
        )

        for msg_num in message_id_list:
            email_counter += 1

            if email_counter % 50 == 0 or email_counter == total_emails:
                print_info(
                    f"Progress: {email_counter}/{total_emails} emails checked, {pdfs_found} PDFs found so far...",
                )

            try:
                status, msg_data = mail.fetch(msg_num, "(RFC822)")

                if status != "OK":
                    logger.warning(f"Failed to fetch email {msg_num}")
                    continue

                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                message_id = email_message.get("Message-ID", str(msg_num.decode()))
                subject = email_message.get("Subject", "No Subject")

                message_id = sanitize_filename(message_id)

                attachment_index = 0
                email_pdf_count = 0

                for part in email_message.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    if part.get("Content-Disposition") is None:
                        continue

                    filename = part.get_filename()

                    if not filename:
                        continue

                    decoded_filename = decode_header(filename)[0][0]
                    if isinstance(decoded_filename, bytes):
                        decoded_filename = decoded_filename.decode()

                    if not decoded_filename.lower().endswith(".pdf"):
                        logger.debug(
                            f"Skipping non-PDF attachment: {decoded_filename}",
                        )
                        continue

                    timestamp = get_timestamp()
                    unique_filename = f"{timestamp}_{message_id}_{attachment_index}.pdf"
                    filepath = DOWNLOADS_RAW_DIR / unique_filename

                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))

                    downloaded_pdfs.append(filepath)
                    email_pdf_count += 1
                    pdfs_found += 1
                    logger.debug(f"Downloaded PDF: {filepath.name}")
                    attachment_index += 1

                if email_pdf_count > 0:
                    if len(subject) > 60:
                        subject = subject[:60] + "..."
                    print_success(f"Found {email_pdf_count} PDF(s) in: {subject}")

            except Exception as exc:  # noqa: BLE001
                logger.error(f"Error processing email {msg_num}: {exc}")
                continue

        mail.close()
        mail.logout()
        print_summary_box(
            "EMAIL DOWNLOAD COMPLETE",
            {
                "Total PDFs Downloaded": len(downloaded_pdfs),
                "From Emails": total_emails,
            },
        )
        logger.info(f"Successfully downloaded {len(downloaded_pdfs)} PDF files")

    except Exception as exc:  # noqa: BLE001
        logger.error(f"Error connecting to IMAP server: {exc}")
        raise

    return downloaded_pdfs


# ==================== PDF SPLITTING ====================


def split_pdfs_to_single_pages(pdf_files: List[Path]) -> List[SinglePageInfo]:
    """Split each PDF into single-page PDFs and name them with account numbers."""

    total_pdfs = len(pdf_files)
    print_info(
        f"Splitting {total_pdfs} PDF files and extracting account numbers...",
    )
    logger.info(f"Starting to split {total_pdfs} PDF files")

    single_pages: List[SinglePageInfo] = []
    pdf_counter = 0

    account_page_count: Dict[str, int] = {}

    for pdf_path in pdf_files:
        pdf_counter += 1
        try:
            reader = PdfReader(str(pdf_path))
            num_pages = len(reader.pages)

            print_progress(
                pdf_counter,
                total_pdfs,
                f"Splitting {pdf_path.name} ({num_pages} pages)",
            )
            logger.info(f"Processing {pdf_path.name} with {num_pages} pages")

            for page_num in range(num_pages):
                try:
                    writer = PdfWriter()
                    writer.add_page(reader.pages[page_num])

                    temp_filename = f"temp_page_{page_num + 1}.pdf"
                    temp_path = SINGLE_PAGES_DIR / temp_filename

                    with open(temp_path, "wb") as f:
                        writer.write(f)

                    account_number = extract_account_number(
                        temp_path,
                        ACCOUNT_REGEX,
                    )

                    if account_number:
                        if account_number not in account_page_count:
                            account_page_count[account_number] = 0
                        account_page_count[account_number] += 1

                        final_filename = (
                            f"{account_number}-page-"
                            f"{account_page_count[account_number]}.pdf"
                        )
                        final_path = SINGLE_PAGES_DIR / final_filename

                        temp_path.rename(final_path)

                        page_info = SinglePageInfo(
                            single_page_path=final_path,
                            source_raw_file=pdf_path.name,
                            page_number=page_num + 1,
                            account_number=account_number,
                        )
                        single_pages.append(page_info)
                        logger.debug(
                            f"Created: {final_filename} (account: {account_number})",
                        )
                    else:
                        unknown_count = account_page_count.get("UNKNOWN", 0) + 1
                        account_page_count["UNKNOWN"] = unknown_count

                        final_filename = f"UNKNOWN-page-{unknown_count}.pdf"
                        final_path = SINGLE_PAGES_DIR / final_filename

                        temp_path.rename(final_path)

                        page_info = SinglePageInfo(
                            single_page_path=final_path,
                            source_raw_file=pdf_path.name,
                            page_number=page_num + 1,
                            extraction_failed=True,
                        )
                        single_pages.append(page_info)
                        logger.debug(
                            f"Created: {final_filename} (no account number found)",
                        )

                except Exception as exc:  # noqa: BLE001
                    logger.error(
                        f"Error splitting page {page_num + 1} of {pdf_path.name}: {exc}",
                    )
                    page_info = SinglePageInfo(
                        single_page_path=SINGLE_PAGES_DIR
                        / f"ERROR-page-{page_num + 1}.pdf",
                        source_raw_file=pdf_path.name,
                        page_number=page_num + 1,
                        extraction_failed=True,
                    )
                    single_pages.append(page_info)
                    continue

        except Exception as exc:  # noqa: BLE001
            print_error(f"Error reading PDF {pdf_path.name}: {exc}")
            logger.error(f"Error reading PDF {pdf_path.name}: {exc}")
            continue

    successful = sum(1 for p in single_pages if p.account_number is not None)
    failed = len(single_pages) - successful

    print_summary_box(
        "PDF SPLITTING & EXTRACTION COMPLETE",
        {
            "Total Pages Created": len(single_pages),
            "Account Numbers Extracted": successful,
            "Failed Extractions": failed,
            "Unique Accounts Found": len(account_page_count)
            - (1 if "UNKNOWN" in account_page_count else 0),
        },
    )
    logger.info(
        "Successfully split into %s single-page PDFs with %s account numbers extracted",
        len(single_pages),
        successful,
    )
    return single_pages


# ==================== ACCOUNT NUMBER EXTRACTION ====================


def extract_account_number(pdf_path: Path, regex_pattern: str) -> Optional[str]:
    """Extract account number from a single-page PDF.

    Works with both patterns that use capturing groups and patterns
    without groups (in which case the full match is returned).
    """

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return None

            page = pdf.pages[0]
            text = page.extract_text()

            if not text:
                logger.debug(
                    "No text extracted from %s when looking for account number",
                    pdf_path.name,
                )
                return None

            pattern = re.compile(regex_pattern, re.IGNORECASE)
            match = pattern.search(text)

            if not match:
                logger.debug(
                    "No regex match for account number in %s",
                    pdf_path.name,
                )
                return None

            # If there are capture groups, prefer them
            if match.groups():
                for i in range(1, len(match.groups()) + 1):
                    group_val = match.group(i)
                    if group_val:
                        account_number = group_val.strip()
                        return account_number

            # No groups defined in pattern → use full match
            full_match = match.group(0).strip()
            if full_match:
                return full_match

            return None

    except Exception as exc:  # noqa: BLE001
        logger.debug(
            "Error extracting text from %s: %s",
            pdf_path.name,
            exc,
        )
        return None


def extract_account_numbers_from_pages(
    single_pages: List[SinglePageInfo],
    regex_pattern: str = ACCOUNT_REGEX,
) -> List[SinglePageInfo]:
    """Extract account numbers from all single-page PDFs (unused helper)."""

    total_pages = len(single_pages)
    print_info(f"Extracting account numbers from {total_pages} pages...")
    logger.info(f"Extracting account numbers from {total_pages} pages")

    successful_extractions = 0
    failed_extractions = 0
    page_counter = 0

    for page_info in single_pages:
        page_counter += 1

        if page_counter % 10 == 0 or page_counter == total_pages:
            print_progress(page_counter, total_pages, "Extracting account numbers...")

        if page_info.extraction_failed:
            failed_extractions += 1
            continue

        account_number = extract_account_number(page_info.single_page_path, regex_pattern)

        if account_number:
            page_info.account_number = account_number
            successful_extractions += 1
            logger.debug(
                "Extracted account %s from %s",
                account_number,
                page_info.single_page_path.name,
            )
        else:
            page_info.extraction_failed = True
            failed_extractions += 1
            logger.debug(
                "Failed to extract account number from %s",
                page_info.single_page_path.name,
            )

    print_summary_box(
        "ACCOUNT EXTRACTION COMPLETE",
        {
            "Successful Extractions": successful_extractions,
            "Failed Extractions": failed_extractions,
            "Success Rate": (
                f"{(successful_extractions / total_pages * 100):.1f}%"
                if total_pages > 0
                else "N/A"
            ),
        },
    )
    logger.info(
        "Account extraction: %s successful, %s failed",
        successful_extractions,
        failed_extractions,
    )
    return single_pages


# ==================== MAPPING AND GROUPING ====================


def load_account_mapping(mapping_file: str) -> AccountMapping:
    """Load account mapping from CSV or Excel file."""

    logger.info(f"Loading account mapping from {mapping_file}")

    try:
        if mapping_file.endswith(".csv"):
            df = pd.read_csv(mapping_file)
        elif mapping_file.endswith((".xlsx", ".xls")):
            df = pd.read_excel(mapping_file)
        else:
            raise ValueError(f"Unsupported file format: {mapping_file}")

        required_columns = {"account_number", "guy_name", "guy_email"}
        df_columns = set(df.columns.str.strip().str.lower())

        if not required_columns.issubset(df_columns):
            missing = required_columns - df_columns
            raise ValueError(f"Missing required columns: {missing}")

        mapping = AccountMapping(df)
        logger.info(f"Loaded {len(mapping.account_to_guy)} account mappings")
        logger.info(f"Found {len(mapping.guy_to_accounts)} unique recipients")

        return mapping

    except Exception as exc:  # noqa: BLE001
        logger.error(f"Error loading mapping file: {exc}")
        raise


def group_pages_by_guy(
    single_pages: List[SinglePageInfo], mapping: AccountMapping
) -> Dict[str, List[SinglePageInfo]]:
    """Group single pages by guy email (accounts per guy)."""

    logger.info("Grouping pages by recipient")

    guy_pages: Dict[str, List[SinglePageInfo]] = {}
    skipped_pages = 0

    for page_info in single_pages:
        if not page_info.account_number or page_info.extraction_failed:
            skipped_pages += 1
            continue

        guy_info = mapping.get_guy_info(page_info.account_number)

        if not guy_info:
            logger.warning(
                "Account %s not found in mapping", page_info.account_number
            )
            skipped_pages += 1
            continue

        guy_name, guy_email = guy_info

        if guy_email not in guy_pages:
            guy_pages[guy_email] = []

        guy_pages[guy_email].append(page_info)

    logger.info(
        "Grouped pages for %s recipients, skipped %s pages",
        len(guy_pages),
        skipped_pages,
    )

    return guy_pages


# ==================== PDF MERGING ====================


def merge_pages_per_guy(
    guy_pages: Dict[str, List[SinglePageInfo]],
    mapping: AccountMapping,
) -> Dict[str, Tuple[Path, str, int]]:
    """Merge all pages for each guy into a single PDF."""

    total_recipients = len(guy_pages)
    print_info(f"Merging PDFs for {total_recipients} recipients...")
    logger.info(f"Merging PDFs for {total_recipients} recipients")

    merged_pdfs: Dict[str, Tuple[Path, str, int]] = {}
    recipient_counter = 0

    for guy_email, pages in guy_pages.items():
        recipient_counter += 1
        try:
            sorted_pages = sorted(
                pages,
                key=lambda p: (
                    p.account_number or "",
                    p.source_raw_file,
                    p.page_number,
                ),
            )

            first_account = sorted_pages[0].account_number
            guy_info = mapping.get_guy_info(first_account) if first_account else None
            guy_name = guy_info[0] if guy_info else "Unknown"

            merger = PdfWriter()

            for page_info in sorted_pages:
                try:
                    reader = PdfReader(str(page_info.single_page_path))
                    merger.add_page(reader.pages[0])
                except Exception as exc:  # noqa: BLE001
                    logger.error(
                        "Error adding page %s: %s",
                        page_info.single_page_path.name,
                        exc,
                    )
                    continue

            timestamp = get_timestamp()
            sanitized_name = sanitize_filename(guy_name)
            sanitized_email = sanitize_filename(guy_email)
            merged_filename = (
                f"{sanitized_name}_{sanitized_email}_{timestamp}.pdf"
            )
            merged_path = MERGED_BY_GUY_DIR / merged_filename

            with open(merged_path, "wb") as f:
                merger.write(f)

            num_pages = len(sorted_pages)
            merged_pdfs[guy_email] = (merged_path, guy_name, num_pages)

            print_progress(
                recipient_counter,
                total_recipients,
                f"✓ {guy_name}: {num_pages} pages merged",
            )
            logger.info(
                "Created merged PDF for %s (%s): %s pages",
                guy_name,
                guy_email,
                num_pages,
            )

        except Exception as exc:  # noqa: BLE001
            print_error(
                f"[{recipient_counter}/{total_recipients}] Error merging for {guy_email}: {exc}",
            )
            logger.error(f"Error merging PDFs for {guy_email}: {exc}")
            continue

    print_summary_box(
        "PDF MERGING COMPLETE",
        {
            "Total Recipients": total_recipients,
            "Merged PDFs Created": len(merged_pdfs),
        },
    )
    logger.info("Successfully created %s merged PDFs", len(merged_pdfs))
    return merged_pdfs


# ==================== EMAIL SENDING ====================


def send_email_with_attachment(
    to_email: str,
    guy_name: str,
    attachment_path: Path,
    subject: str = EMAIL_SUBJECT,
    body_template: str = EMAIL_BODY_TEMPLATE,
) -> bool:
    """Send email with PDF attachment via SMTP."""

    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_USER
        msg["To"] = to_email
        msg["Subject"] = subject

        body = body_template.format(guy_name=guy_name)
        msg.attach(MIMEText(body, "plain"))

        with open(attachment_path, "rb") as f:
            pdf_attachment = MIMEApplication(f.read(), _subtype="pdf")
            pdf_attachment.add_header(
                "Content-Disposition",
                "attachment",
                filename="bills.pdf",
            )
            msg.attach(pdf_attachment)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            if USE_TLS:
                server.starttls()
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.send_message(msg)

        logger.info("Successfully sent email to %s", to_email)
        return True

    except Exception as exc:  # noqa: BLE001
        logger.error("Error sending email to %s: %s", to_email, exc)
        return False


def send_emails_for_merged_pdfs(
    merged_pdfs: Dict[str, Tuple[Path, str, int]],
    dry_run: bool = False,
    batch_size: int = BATCH_SIZE,
    batch_delay: int = BATCH_DELAY_SECONDS,
) -> Dict[str, bool]:
    """Send emails with merged PDFs to recipients."""

    total_emails = len(merged_pdfs)
    if dry_run:
        print_warning("DRY RUN MODE - No emails will actually be sent")
    print_info(f"Sending emails to {total_emails} recipients...")
    logger.info(
        "Sending emails to %s recipients (dry_run=%s)", total_emails, dry_run
    )

    send_results: Dict[str, bool] = {}
    emails_sent = 0
    email_counter = 0

    for guy_email, (pdf_path, guy_name, num_pages) in merged_pdfs.items():
        email_counter += 1
        if dry_run:
            print_progress(
                email_counter,
                total_emails,
                f"[DRY RUN] Would send to {guy_name} ({guy_email}) - {num_pages} pages",
            )
            logger.info(
                "[DRY RUN] Would send email to %s with attachment %s",
                guy_email,
                pdf_path.name,
            )
            send_results[guy_email] = True
            emails_sent += 1
        else:
            success = send_email_with_attachment(guy_email, guy_name, pdf_path)
            send_results[guy_email] = success

            if success:
                print_progress(
                    email_counter,
                    total_emails,
                    f"✓ Sent to {guy_name} ({guy_email}) - {num_pages} pages",
                )
                emails_sent += 1
            else:
                print_progress(
                    email_counter,
                    total_emails,
                    f"✗ Failed to send to {guy_name} ({guy_email})",
                )

            if emails_sent > 0 and emails_sent % batch_size == 0:
                print_info(
                    f"Batch complete ({emails_sent} emails). Pausing for {batch_delay} seconds...",
                )
                logger.info(
                    "Sent %s emails, pausing for %s seconds...",
                    emails_sent,
                    batch_delay,
                )
                time.sleep(batch_delay)

    successful = sum(1 for success in send_results.values() if success)
    failed = len(send_results) - successful

    print_summary_box(
        "EMAIL SENDING COMPLETE",
        {
            "Total Emails": total_emails,
            "Successful": successful,
            "Failed": failed,
            "Mode": "DRY RUN" if dry_run else "PRODUCTION",
        },
    )
    logger.info(
        "Email sending complete: %s successful, %s failed", successful, failed
    )
    return send_results


# ==================== LOGGING ====================


def write_run_log(
    single_pages: List[SinglePageInfo],
    mapping: AccountMapping,
    guy_pages: Dict[str, List[SinglePageInfo]],
    merged_pdfs: Dict[str, Tuple[Path, str, int]],
    send_results: Dict[str, bool],
) -> Path:
    """Write CSV log file for this run."""

    timestamp = get_timestamp()
    log_filename = f"run_{timestamp}.csv"
    log_path = LOGS_DIR / log_filename

    logger.info(f"Writing run log to {log_path}")

    log_entries: List[LogEntry] = []
    processed_accounts: Set[str] = set()

    for guy_email, (_, _, _) in merged_pdfs.items():
        accounts = mapping.get_accounts_for_guy(guy_email)

        pages_for_guy = guy_pages.get(guy_email, [])

        account_page_counts: Dict[str, int] = {}
        for page in pages_for_guy:
            if page.account_number:
                account_page_counts[page.account_number] = (
                    account_page_counts.get(page.account_number, 0) + 1
                )

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
                failed_pages="",
            )
            log_entries.append(entry)

    account_skip_reasons: Dict[str, List[int]] = {}

    for page in single_pages:
        if page.account_number and page.account_number not in processed_accounts:
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
            failed_pages=str(failed_page_nums),
        )
        log_entries.append(entry)

    failed_extraction_pages = [
        p for p in single_pages if p.extraction_failed or not p.account_number
    ]
    if failed_extraction_pages:
        failed_page_nums = [p.page_number for p in failed_extraction_pages]
        entry = LogEntry(
            date=get_current_date(),
            time=get_current_time(),
            status="Skipped",
            account_number="",
            destination_email="",
            num_pages=0,
            failed_pages=str(failed_page_nums),
        )
        log_entries.append(entry)

    with open(log_path, "w", newline="", encoding="utf-8") as csvfile:
        fieldnames = [
            "date",
            "time",
            "status",
            "account_number",
            "destination_email",
            "num_pages",
            "failed_pages",
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        for entry in log_entries:
            writer.writerow(entry.to_dict())

    logger.info("Successfully wrote %s log entries to %s", len(log_entries), log_path)
    return log_path


# ==================== MAIN ====================


def main() -> None:
    """Main entry point for the script."""

    parser = argparse.ArgumentParser(
        description="Extract bills from Outlook emails and send merged PDFs",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Process everything but do not send emails",
    )

    args = parser.parse_args()

    print_header("ewaBills - Automated Billing Workflow System")
    print(f"  Mode: {'DRY RUN' if args.dry_run else 'PRODUCTION'}")
    print(f"  Server: {IMAP_SERVER}")
    print(f"  User: {EMAIL_USER}")
    print("=" * 80)

    logger.info("=" * 80)
    logger.info("Starting extract_and_send.py")
    logger.info("Dry run mode: %s", args.dry_run)
    logger.info("=" * 80)

    try:
        ensure_directories_exist()

        print_section("STEP 1: Fetching Emails and Downloading PDFs")
        downloaded_pdfs = fetch_emails_and_download_pdfs(dry_run=args.dry_run)

        if not downloaded_pdfs:
            print_warning("No PDF attachments found in emails")
            return

        print_section("STEP 2: Splitting PDFs and Extracting Account Numbers")
        single_pages = split_pdfs_to_single_pages(downloaded_pdfs)

        if not single_pages:
            print_warning("No pages extracted from PDFs")
            return

        print_section("STEP 3: Loading Account Mapping")
        mapping = load_account_mapping(MAPPING_FILE)
        print_success(
            f"Loaded mapping for {len(mapping.guy_to_accounts)} recipients",
        )

        print_section("STEP 4: Grouping Pages by Recipient")
        guy_pages = group_pages_by_guy(single_pages, mapping)

        if not guy_pages:
            print_warning("No pages matched to recipients")
            return

        print_section("STEP 5: Merging PDFs per Recipient")
        merged_pdfs = merge_pages_per_guy(guy_pages, mapping)

        if not merged_pdfs:
            print_warning("No merged PDFs created")
            return

        print_section("STEP 6: Sending Emails")
        send_results = send_emails_for_merged_pdfs(
            merged_pdfs,
            dry_run=args.dry_run,
        )

        print_section("STEP 7: Writing Log File")
        log_path = write_run_log(
            single_pages,
            mapping,
            guy_pages,
            merged_pdfs,
            send_results,
        )
        print_success(f"Log file created: {log_path}")

        print_header("PROCESSING COMPLETE!")
        print(f"  Log file: {log_path}")
        print(f"  Mode: {'DRY RUN' if args.dry_run else 'PRODUCTION'}")
        print("=" * 80)

        logger.info("=" * 80)
        logger.info("Processing complete!")
        logger.info("Log file: %s", log_path)
        logger.info("=" * 80)

    except Exception as exc:  # noqa: BLE001
        logger.error("Fatal error: %s", exc, exc_info=True)
        raise


if __name__ == "__main__":
    main()
