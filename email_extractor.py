#!/usr/bin/env python3
"""
IMAP Email Search and Export Tool
Connects to IMAP servers, searches messages, and exports to .eml and .pdf formats
Enhanced: Multi-folder search capability and Unicode character support
"""

import imaplib
import email
import argparse
import os
import re
import sys
import logging
from datetime import datetime
from dotenv import load_dotenv
from email.header import decode_header
from email.utils import parsedate_to_datetime
from fpdf import FPDF
import html2text
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def clean_text_for_pdf(text):
    """Clean text for PDF rendering, removing problematic Unicode characters"""
    if not text:
        return ""
    
    # Remove Zero Width characters that cause rendering issues
    text = re.sub(r'[\u200B-\u200F\uFEFF\u202A-\u202E]', '', text)
    
    # Replace smart quotes and special characters with ASCII equivalents
    text = text.replace('\u201C', '"').replace('\u201D', '"')  # Smart double quotes  
    text = text.replace('\u2018', "'").replace('\u2019', "'")  # Smart single quotes/apostrophes
    text = text.replace('\u2013', '-').replace('\u2014', '-')  # En dash, em dash
    text = text.replace('\u2026', '...')  # Ellipsis
    text = text.replace('\u2022', '*')    # Bullet point
    text = text.replace('\u00B0', ' degrees ')  # Degree symbol
    
    # Remove control characters except newlines and tabs
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', text)
    
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    
    # Final safety check: keep only characters that can be encoded in latin-1
    try:
        text.encode('latin-1')
        return text
    except UnicodeEncodeError:
        # Filter to ASCII printable characters only
        safe_chars = []
        for char in text:
            if ord(char) < 128 and (char.isprintable() or char in '\n\t '):
                safe_chars.append(char)
            else:
                safe_chars.append('?')  # Replace problematic chars with ?
        return ''.join(safe_chars)

class EmailPDF(FPDF):
    """Custom FPDF class for email formatting with headers and footers"""
    
    def __init__(self, email_subject=""):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.email_subject = clean_text_for_pdf(email_subject)
        # Set proper margins to prevent horizontal space issues
        self.set_margins(left=20, top=20, right=20)
        self.set_auto_page_break(auto=True, margin=25)
        
        # Add Unicode font support for special characters
        try:
            # Try to add DejaVu font for Unicode support
            self.add_font("DejaVu", "", "DejaVuSans.ttf")
            self.unicode_font_available = True
        except:
            # Fall back to Arial if DejaVu not available
            self.unicode_font_available = False
    
    def header(self):
        """Page header"""
        font_name = "DejaVu" if self.unicode_font_available else "Arial"
        self.set_font(font_name, 'B', 12)
        # Truncate and clean subject for header
        header_text = self.email_subject[:50] + "..." if len(self.email_subject) > 50 else self.email_subject
        try:
            self.cell(0, 10, f'Email: {header_text}', 0, 1, 'C')
        except:
            # Fallback for problematic characters
            self.cell(0, 10, 'Email: [Subject contains unsupported characters]', 0, 1, 'C')
        self.ln(5)
    
    def footer(self):
        """Page footer with page numbers"""
        self.set_y(-15)
        font_name = "DejaVu" if self.unicode_font_available else "Arial"
        self.set_font(font_name, 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def sanitize_filename(text):
    """Replace non-alphanumeric characters with hyphens for safe filenames"""
    if not text:
        return ""
    # Replace non-alphanumeric characters with hyphens
    sanitized = re.sub(r'[^A-Za-z0-9]', '-', text)
    # Remove multiple consecutive hyphens
    sanitized = re.sub(r'-+', '-', sanitized)
    # Remove leading/trailing hyphens
    return sanitized.strip('-')

def parse_folder_list(folder_line):
    """Parse IMAP folder list response to extract folder name"""
    try:
        # Decode bytes to string
        folder_str = folder_line.decode() if isinstance(folder_line, bytes) else folder_line
        
        # Extract folder name from IMAP list response
        # Format: (flags) "delimiter" "folder_name"
        parts = folder_str.split(' "')
        if len(parts) >= 3:
            folder_name = parts[2].rstrip('"')
            return folder_name
        
        # Alternative parsing for different server responses
        match = re.search(r'"([^"]*)"$', folder_str)
        if match:
            return match.group(1)
        
        return None
    except Exception as e:
        logger.warning(f"Error parsing folder: {e}")
        return None

def get_all_folders(imap_conn, verbose=False):
    """Get list of all available folders in the mailbox"""
    try:
        typ, folder_data = imap_conn.list()
        
        if typ != "OK":
            logger.warning("Failed to get folder list")
            return ["INBOX"]  # Fallback to INBOX only
        
        folders = []
        for folder_line in folder_data:
            folder_name = parse_folder_list(folder_line)
            if folder_name:
                folders.append(folder_name)
        
        if verbose:
            print(f"Available folders: {folders}")
        
        # Ensure INBOX is included if not found
        if "INBOX" not in folders:
            folders.insert(0, "INBOX")
        
        return folders
    
    except Exception as e:
        logger.warning(f"Error getting folder list: {e}")
        return ["INBOX"]  # Fallback to INBOX only

def build_export_folder(base_path, search_date, sender=None, recipient=None, keywords=None):
    """Build export folder path based on search criteria"""
    folder_parts = []
    
    # Always include date
    folder_parts.append(sanitize_filename(search_date))
    
    # Add other criteria if provided
    if sender:
        folder_parts.append(sanitize_filename(sender))
    if recipient:
        folder_parts.append(sanitize_filename(recipient))
    if keywords:
        folder_parts.append(sanitize_filename(keywords))
    
    folder_name = '_'.join(folder_parts)
    full_path = Path(base_path) / folder_name
    full_path.mkdir(parents=True, exist_ok=True)
    return str(full_path)

def decode_mime_words(s):
    """Decode MIME encoded words in email headers"""
    if s is None:
        return ""
    if isinstance(s, str):
        return s
    
    decoded_parts = decode_header(s)
    decoded_string = ""
    
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            try:
                decoded_string += part.decode(encoding or "utf-8", errors="ignore")
            except (LookupError, TypeError):
                decoded_string += part.decode("utf-8", errors="ignore")
        else:
            decoded_string += part
    
    return decoded_string

def extract_text_from_email(msg):
    """Extract text content from email message, handling multipart and HTML"""
    text_content = ""
    html_content = ""
    
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition", ""))
            
            # Skip attachments
            if "attachment" in content_disposition:
                continue
                
            try:
                if content_type == "text/plain":
                    charset = part.get_content_charset() or "utf-8"
                    payload = part.get_payload(decode=True)
                    if payload:
                        text_content += payload.decode(charset, errors="ignore") + "\n"
                elif content_type == "text/html":
                    charset = part.get_content_charset() or "utf-8"
                    payload = part.get_payload(decode=True)
                    if payload:
                        html_content += payload.decode(charset, errors="ignore") + "\n"
            except Exception as e:
                logger.warning(f"Error extracting content: {e}")
                continue
    else:
        # Single part message
        content_type = msg.get_content_type()
        try:
            if content_type == "text/plain":
                charset = msg.get_content_charset() or "utf-8"
                payload = msg.get_payload(decode=True)
                if payload:
                    text_content = payload.decode(charset, errors="ignore")
            elif content_type == "text/html":
                charset = msg.get_content_charset() or "utf-8"
                payload = msg.get_payload(decode=True)
                if payload:
                    html_content = payload.decode(charset, errors="ignore")
        except Exception as e:
            logger.warning(f"Error extracting single part content: {e}")
    
    # Prefer plain text, fall back to converted HTML
    if text_content.strip():
        return text_content.strip()
    elif html_content.strip():
        # Convert HTML to plain text
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.ignore_images = True
        h.body_width = 0  # No line wrapping
        return h.handle(html_content).strip()
    else:
        return "No readable content found"

def load_env_config(env_path):
    """Load configuration from .env file"""
    if not os.path.exists(env_path):
        logger.info(f"No .env file found at {env_path}")
        return {}
    
    load_dotenv(env_path)
    config = {}
    
    env_vars = [
        'MAILHOST', 'MAILPORT', 'CRYPT', 'USERNAME', 'PASSWORD',
        'SENDER', 'RECIPIENT', 'KEYWORDS', 'START_DATE', 'END_DATE'
    ]
    
    for var in env_vars:
        value = os.getenv(var)
        if value:
            config[var.lower()] = value
    
    return config

def connect_imap(host, port, crypt, username, password):
    """Connect to IMAP server with specified encryption"""
    logger.info(f"Connecting to {host}:{port} with {crypt} encryption")
    
    crypt = crypt.lower()
    
    try:
        if crypt == "ssl":
            connection = imaplib.IMAP4_SSL(host, int(port))
        else:
            connection = imaplib.IMAP4(host, int(port))
            if crypt == "starttls":
                connection.starttls()
        
        connection.login(username, password)
        logger.info("Successfully connected and authenticated")
        return connection
    
    except Exception as e:
        logger.error(f"Failed to connect to IMAP server: {e}")
        sys.exit(1)

def search_folder_messages(imap_conn, folder_name, sender, recipient, keywords, start_date, end_date, case_sensitive, verbose=False):
    """Search for messages in a specific folder matching criteria"""
    try:
        # Select folder (some folders may need quotes)
        try:
            typ, data = imap_conn.select(f'"{folder_name}"', readonly=True)
        except:
            typ, data = imap_conn.select(folder_name, readonly=True)
        
        if typ != "OK":
            if verbose:
                print(f"  ❌ Cannot access folder: {folder_name}")
            return []
        
        folder_message_count = int(data[0]) if data[0] else 0
        if verbose:
            print(f"  📁 Folder '{folder_name}': {folder_message_count} messages")
        
        if folder_message_count == 0:
            return []
        
        # Build search criteria - default to ALL (no date restriction)
        search_criteria = ["ALL"]
        
        # Date range criteria (only if provided)
        if start_date:
            since_str = start_date.strftime("%d-%b-%Y")
            search_criteria.append(f'SINCE "{since_str}"')
        
        if end_date:
            before_str = end_date.strftime("%d-%b-%Y") 
            search_criteria.append(f'BEFORE "{before_str}"')
        
        # Execute search
        typ, message_ids = imap_conn.search(None, *search_criteria)
        
        if typ != "OK":
            if verbose:
                print(f"  ❌ Search failed in folder: {folder_name}")
            return []
        
        ids = message_ids[0].split()
        if verbose and len(ids) > 0:
            print(f"  🔍 Found {len(ids)} messages in date range")
        
        # Filter messages by sender, recipient, and keywords
        matching_messages = []
        
        for i, msg_id in enumerate(ids, 1):
            try:
                typ, msg_data = imap_conn.fetch(msg_id, '(RFC822)')
                if typ != "OK":
                    continue
                
                raw_email = msg_data[0][1]
                email_message = email.message_from_bytes(raw_email)
                
                # Extract and decode headers
                from_addr = decode_mime_words(email_message.get("From", ""))
                to_addr = decode_mime_words(email_message.get("To", ""))
                cc_addr = decode_mime_words(email_message.get("Cc", ""))
                bcc_addr = decode_mime_words(email_message.get("Bcc", ""))
                subject = decode_mime_words(email_message.get("Subject", ""))
                date_header = email_message.get("Date", "")
                
                # Combine all recipient fields
                all_recipients = f"{to_addr} {cc_addr} {bcc_addr}".strip()
                
                # Extract body text
                body_text = extract_text_from_email(email_message)
                
                # Verbose output - show message being reviewed
                if verbose:
                    print(f"\n    [{i}/{len(ids)}] Reviewing in {folder_name}:")
                    print(f"      Date: {date_header}")
                    print(f"      From: {from_addr}")
                    print(f"      To: {to_addr}")
                    print(f"      Subject: {subject}")
                    print(f"      Body preview: {body_text[:100]}{'...' if len(body_text) > 100 else ''}")
                
                # Apply filters
                def matches(search_term, text_to_search):
                    if not search_term:
                        return True
                    if case_sensitive:
                        return search_term in text_to_search
                    else:
                        return search_term.lower() in text_to_search.lower()
                
                sender_match = matches(sender, from_addr)
                recipient_match = matches(recipient, all_recipients)
                keyword_match = matches(keywords, f"{subject} {body_text}")
                
                if verbose:
                    print(f"      Sender match: {sender_match} (searching for: {sender or 'ANY'})")
                    print(f"      Recipient match: {recipient_match} (searching for: {recipient or 'ANY'})")
                    print(f"      Keyword match: {keyword_match} (searching for: {keywords or 'ANY'})")
                    print(f"      Overall match: {sender_match and recipient_match and keyword_match}")
                
                if sender_match and recipient_match and keyword_match:
                    # Parse email date
                    try:
                        email_date = parsedate_to_datetime(date_header)
                    except Exception:
                        email_date = datetime.now()
                    
                    matching_messages.append({
                        "id": msg_id.decode(),
                        "folder": folder_name,
                        "from": from_addr,
                        "to": to_addr,
                        "subject": subject,
                        "date": email_date,
                        "date_header": date_header,
                        "raw": raw_email,
                        "body": body_text,
                        "email_obj": email_message
                    })
                    
                    if verbose:
                        print(f"      ✅ MATCH FOUND - Added to results")
                
            except Exception as e:
                logger.warning(f"Error processing message {msg_id} in {folder_name}: {e}")
                if verbose:
                    print(f"      ❌ ERROR processing message: {e}")
                continue
        
        if verbose and matching_messages:
            print(f"  ✅ Found {len(matching_messages)} matching messages in {folder_name}")
        
        return matching_messages
    
    except Exception as e:
        logger.warning(f"Error searching folder {folder_name}: {e}")
        return []

def search_all_messages(imap_conn, sender, recipient, keywords, start_date, end_date, case_sensitive, search_all_folders=False, verbose=False):
    """Search for messages across all folders or just INBOX"""
    logger.info("Searching for messages...")
    
    if search_all_folders:
        folders = get_all_folders(imap_conn, verbose)
        logger.info(f"Searching across {len(folders)} folders")
    else:
        folders = ["INBOX"]
        logger.info("Searching INBOX only")
    
    all_matching_messages = []
    
    for folder in folders:
        if verbose:
            print(f"\n🔍 Searching folder: {folder}")
        
        folder_matches = search_folder_messages(
            imap_conn, folder, sender, recipient, keywords, 
            start_date, end_date, case_sensitive, verbose
        )
        
        all_matching_messages.extend(folder_matches)
    
    logger.info(f"Found {len(all_matching_messages)} total matching messages across all searched folders")
    return all_matching_messages

def save_message_files(message, export_folder, verbose=False):
    """Save message as both .eml and .pdf files"""
    try:
        # Generate filename components
        date_str = message["date"].strftime('%Y-%m-%d')
        sender = message["from"][:30]  # Limit length
        recipient = message["to"][:30]  # Limit length  
        subject = message["subject"][:16]  # First 16 chars as specified
        
        # Build filename
        filename_parts = [
            sanitize_filename(date_str),
            sanitize_filename(sender),
            sanitize_filename(recipient), 
            sanitize_filename(subject)
        ]
        
        base_filename = '_'.join(part for part in filename_parts if part)
        
        # Ensure filename isn't too long
        if len(base_filename) > 200:
            base_filename = base_filename[:200]
        
        if verbose:
            print(f"Saving files with base name: {base_filename}")
            print(f"  Source folder: {message.get('folder', 'UNKNOWN')}")
        
        # Save .eml file
        eml_path = Path(export_folder) / f"{base_filename}.eml"
        with open(eml_path, "wb") as f:
            f.write(message["raw"])
        
        if verbose:
            print(f"  ✅ Saved EML: {eml_path}")
        
        # Create PDF with Unicode support
        pdf = EmailPDF(message["subject"])
        pdf.add_page()
        
        # Clean all text for PDF rendering
        safe_from = clean_text_for_pdf(message["from"])
        safe_to = clean_text_for_pdf(message["to"])
        safe_subject = clean_text_for_pdf(message["subject"])
        safe_body = clean_text_for_pdf(message["body"])
        safe_date = clean_text_for_pdf(message["date_header"])
        safe_folder = clean_text_for_pdf(message.get("folder", "UNKNOWN"))
        
        # Choose font based on availability
        font_name = "DejaVu" if pdf.unicode_font_available else "Arial"
        
        # Header information with safe layout
        pdf.set_font(font_name, "B", 14)
        pdf.cell(0, 10, "Email Message", ln=1, align="C")
        pdf.ln(5)
        
        # Email metadata using only cell() to prevent multi_cell issues
        pdf.set_font(font_name, "B", 10)
        pdf.cell(25, 6, "Folder:", border=0)
        pdf.set_font(font_name, "", 10)
        pdf.cell(145, 6, safe_folder[:70], ln=1, border=0)
        
        pdf.set_font(font_name, "B", 10)
        pdf.cell(25, 6, "Date:", border=0)
        pdf.set_font(font_name, "", 10)
        pdf.cell(145, 6, safe_date[:70], ln=1, border=0)
        
        pdf.set_font(font_name, "B", 10)
        pdf.cell(25, 6, "From:", border=0)
        pdf.set_font(font_name, "", 10)
        pdf.cell(145, 6, safe_from[:70], ln=1, border=0)
        
        pdf.set_font(font_name, "B", 10)
        pdf.cell(25, 6, "To:", border=0)
        pdf.set_font(font_name, "", 10)
        pdf.cell(145, 6, safe_to[:70], ln=1, border=0)
        
        pdf.set_font(font_name, "B", 10)
        pdf.cell(25, 6, "Subject:", border=0)
        pdf.set_font(font_name, "", 10)
        pdf.cell(145, 6, safe_subject[:70], ln=1, border=0)
        
        pdf.ln(8)
        pdf.set_font(font_name, "B", 10)
        pdf.cell(0, 6, "Message Body:", ln=1, border=0)
        pdf.ln(3)
        
        # Message body with conservative approach
        pdf.set_font(font_name, "", 9)
        
        # Truncate extremely long bodies
        if len(safe_body) > 5000:
            safe_body = safe_body[:5000] + "\n\n[MESSAGE TRUNCATED - Original too long for PDF display]"
        
        # Split body into lines and use cell() for better control
        body_lines = safe_body.split('\n')
        for line in body_lines:
            if len(line) > 80:
                # Split very long lines
                words = line.split(' ')
                current_line = ""
                for word in words:
                    if len(current_line + word) < 80:
                        current_line += word + " "
                    else:
                        if current_line:
                            try:
                                pdf.cell(0, 5, current_line.strip(), ln=1, border=0)
                            except:
                                pdf.cell(0, 5, "[Line contains unsupported characters]", ln=1, border=0)
                        current_line = word + " "
                if current_line:
                    try:
                        pdf.cell(0, 5, current_line.strip(), ln=1, border=0)
                    except:
                        pdf.cell(0, 5, "[Line contains unsupported characters]", ln=1, border=0)
            else:
                try:
                    pdf.cell(0, 5, line, ln=1, border=0)
                except:
                    pdf.cell(0, 5, "[Line contains unsupported characters]", ln=1, border=0)
        
        # Save PDF
        pdf_path = Path(export_folder) / f"{base_filename}.pdf"
        pdf.output(str(pdf_path))
        
        if verbose:
            print(f"  ✅ Saved PDF: {pdf_path}")
        
        logger.info(f"Saved: {base_filename}")
        return True
        
    except Exception as e:
        logger.error(f"Error saving message files: {e}")
        if verbose:
            print(f"  ❌ ERROR saving files: {e}")
        return False

def main():
    """Main function"""
    parser = argparse.ArgumentParser(
        description="Search and export emails from IMAP servers",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --mailhost imap.gmail.com --username user@gmail.com --password pass --sender boss@company.com
  %(prog)s --env config.env --keywords "urgent project" --start-date 2023-01-01 --all-folders
  %(prog)s --mailhost mail.company.com --crypt starttls --recipient client@company.com --verbose --all-folders
        """
    )
    
    # Connection arguments
    parser.add_argument('--mailhost', help='IMAP server hostname')
    parser.add_argument('--mailport', type=int, help='IMAP server port')
    parser.add_argument('--crypt', choices=['none', 'starttls', 'ssl'], 
                       help='Connection encryption (default: auto-detect)')
    parser.add_argument('--username', help='IMAP username')
    parser.add_argument('--password', help='IMAP password')
    
    # Search criteria
    parser.add_argument('--sender', help='Filter by sender email (partial match)')
    parser.add_argument('--recipient', help='Filter by recipient email (partial match)')
    parser.add_argument('--keywords', help='Search keywords in subject/body')
    parser.add_argument('--start-date', help='Start date (YYYY-MM-DD format, default: all time)')
    parser.add_argument('--end-date', help='End date (YYYY-MM-DD format, default: all time)')
    
    # Options
    parser.add_argument('--case-sensitive', action='store_true',
                       help='Enable case-sensitive matching')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='Enable verbose output showing messages under review')
    parser.add_argument('--all-folders', action='store_true',
                       help='Search all folders in mailbox (default: INBOX only)')
    parser.add_argument('--env', default='.env', 
                       help='Path to .env file (default: .env)')
    parser.add_argument('--export-dir', default='./export',
                       help='Export directory (default: ./export)')
    
    args = parser.parse_args()
    
    # Load environment configuration
    env_config = load_env_config(args.env)
    
    # Merge configuration (command line takes precedence)
    config = {}
    for key in ['mailhost', 'mailport', 'crypt', 'username', 'password',
               'sender', 'recipient', 'keywords', 'start_date', 'end_date']:
        
        # Command line argument takes precedence
        cli_value = getattr(args, key, None)
        if cli_value is not None:
            config[key] = cli_value
        # Fall back to environment variable
        elif key in env_config:
            config[key] = env_config[key]
    
    # Validate required parameters
    required_params = ['mailhost', 'username', 'password']
    missing_params = [p for p in required_params if not config.get(p)]
    if missing_params:
        logger.error(f"Missing required parameters: {', '.join(missing_params)}")
        parser.print_help()
        sys.exit(1)
    
    # Set defaults
    if not config.get('mailport'):
        config['mailport'] = 993 if config.get('crypt', '').lower() == 'ssl' else 143
    
    if not config.get('crypt'):
        # Auto-detect based on port
        config['crypt'] = 'ssl' if config['mailport'] == 993 else 'starttls'
    
    # Parse dates (default to None for all-time search)
    start_date = None
    end_date = None
    
    if config.get('start_date'):
        try:
            start_date = datetime.strptime(config['start_date'], '%Y-%m-%d')
        except ValueError:
            logger.error("Invalid start date format. Use YYYY-MM-DD")
            sys.exit(1)
    
    if config.get('end_date'):
        try:
            end_date = datetime.strptime(config['end_date'], '%Y-%m-%d')
        except ValueError:
            logger.error("Invalid end date format. Use YYYY-MM-DD")
            sys.exit(1)
    
    if args.verbose:
        print(f"\n🔍 Search Configuration:")
        print(f"  IMAP Server: {config['mailhost']}:{config['mailport']} ({config['crypt']})")
        print(f"  Username: {config['username']}")
        print(f"  Search scope: {'ALL FOLDERS' if args.all_folders else 'INBOX ONLY'}")
        print(f"  Sender filter: {config.get('sender', 'ANY')}")
        print(f"  Recipient filter: {config.get('recipient', 'ANY')}")
        print(f"  Keywords filter: {config.get('keywords', 'ANY')}")
        print(f"  Date range: {start_date.strftime('%Y-%m-%d') if start_date else 'ALL TIME'} to {end_date.strftime('%Y-%m-%d') if end_date else 'ALL TIME'}")
        print(f"  Case sensitive: {args.case_sensitive}")
        print(f"  Export directory: {args.export_dir}\n")
    
    # Connect to IMAP server
    imap_conn = connect_imap(
        config['mailhost'],
        config['mailport'], 
        config['crypt'],
        config['username'],
        config['password']
    )
    
    try:
        # Search for messages
        messages = search_all_messages(
            imap_conn,
            config.get('sender'),
            config.get('recipient'), 
            config.get('keywords'),
            start_date,
            end_date,
            args.case_sensitive,
            args.all_folders,
            args.verbose
        )
        
        if not messages:
            logger.info("No matching messages found")
            if args.verbose:
                print("\n❌ No messages matched your search criteria")
            return
        
        # Create export folder
        search_date = datetime.now().strftime('%Y-%m-%d')
        export_folder = build_export_folder(
            args.export_dir,
            search_date,
            config.get('sender'),
            config.get('recipient'),
            config.get('keywords')
        )
        
        logger.info(f"Export folder: {export_folder}")
        if args.verbose:
            print(f"\n📁 Export folder: {export_folder}")
            print(f"\n💾 Saving {len(messages)} matching messages...")
        
        # Export messages
        success_count = 0
        for i, message in enumerate(messages, 1):
            if args.verbose:
                print(f"\n[{i}/{len(messages)}] Processing from {message.get('folder', 'UNKNOWN')}:")
                print(f"  From: {message['from']}")
                print(f"  Subject: {message['subject']}")
            
            if save_message_files(message, export_folder, args.verbose):
                success_count += 1
        
        logger.info(f"Successfully exported {success_count}/{len(messages)} messages")
        if args.verbose:
            print(f"\n✅ Export completed: {success_count}/{len(messages)} messages saved successfully")
        
    finally:
        # Clean up connection
        try:
            imap_conn.close()
        except:
            pass
        try:
            imap_conn.logout()
        except:
            pass

if __name__ == "__main__":
    main()
