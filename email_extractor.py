#!/usr/bin/env python3
"""
IMAP Email Search and Export Tool
Connects to IMAP servers, searches messages, and exports to .eml and .pdf formats
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

class EmailPDF(FPDF):
    """Custom FPDF class for email formatting with headers and footers"""
    
    def __init__(self, email_subject=""):
        super().__init__()
        self.email_subject = email_subject
        self.set_auto_page_break(auto=True, margin=15)
    
    def header(self):
        """Page header"""
        self.set_font('Arial', 'B', 12)
        # Truncate long subjects for header
        header_text = self.email_subject[:50] + "..." if len(self.email_subject) > 50 else self.email_subject
        self.cell(0, 10, f'Email: {header_text}', 0, 1, 'C')
        self.ln(5)
    
    def footer(self):
        """Page footer with page numbers"""
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
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

def sanitize_folder_name(text):
    """Sanitize text for folder names (same as filename but for consistency)"""
    return sanitize_filename(text)

def build_export_folder(base_path, search_date, sender=None, recipient=None, keywords=None):
    """Build export folder path based on search criteria"""
    folder_parts = []
    
    # Always include date
    folder_parts.append(sanitize_folder_name(search_date))
    
    # Add other criteria if provided
    if sender:
        folder_parts.append(sanitize_folder_name(sender))
    if recipient:
        folder_parts.append(sanitize_folder_name(recipient))
    if keywords:
        folder_parts.append(sanitize_folder_name(keywords))
    
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

def search_messages(imap_conn, sender, recipient, keywords, start_date, end_date, case_sensitive):
    """Search for messages matching criteria"""
    logger.info("Searching for messages...")
    
    try:
        imap_conn.select("INBOX")
        
        # Build search criteria
        search_criteria = ["ALL"]
        
        # Date range criteria
        if start_date:
            since_str = start_date.strftime("%d-%b-%Y")
            search_criteria.append(f'SINCE "{since_str}"')
        
        if end_date:
            before_str = end_date.strftime("%d-%b-%Y") 
            search_criteria.append(f'BEFORE "{before_str}"')
        
        # Execute search
        typ, message_ids = imap_conn.search(None, *search_criteria)
        
        if typ != "OK":
            logger.error("Failed to search messages")
            return []
        
        ids = message_ids[0].split()
        logger.info(f"Found {len(ids)} messages in date range")
        
        # Filter messages by sender, recipient, and keywords
        matching_messages = []
        
        for msg_id in ids:
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
                
                if sender_match and recipient_match and keyword_match:
                    # Parse email date
                    try:
                        email_date = parsedate_to_datetime(date_header)
                    except Exception:
                        email_date = datetime.now()
                    
                    matching_messages.append({
                        "id": msg_id.decode(),
                        "from": from_addr,
                        "to": to_addr,
                        "subject": subject,
                        "date": email_date,
                        "date_header": date_header,
                        "raw": raw_email,
                        "body": body_text,
                        "email_obj": email_message
                    })
            
            except Exception as e:
                logger.warning(f"Error processing message {msg_id}: {e}")
                continue
        
        logger.info(f"Found {len(matching_messages)} matching messages")
        return matching_messages
    
    except Exception as e:
        logger.error(f"Error during message search: {e}")
        return []

def save_message_files(message, export_folder):
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
        
        # Save .eml file
        eml_path = Path(export_folder) / f"{base_filename}.eml"
        with open(eml_path, "wb") as f:
            f.write(message["raw"])
        
        # Create PDF with unbreakable header section
        pdf = EmailPDF(message["subject"])
        pdf.add_page()
        pdf.set_font("Arial", size=10)
        
        # Header information in unbreakable section
        with pdf.unbreakable() as doc:
            doc.set_font("Arial", "B", 12)
            doc.cell(0, 8, "Email Message", ln=1, align="C")
            doc.ln(3)
            
            doc.set_font("Arial", "B", 10)
            doc.cell(30, 6, "Date:", border=0)
            doc.set_font("Arial", size=10)
            doc.cell(0, 6, message["date_header"], ln=1, border=0)
            
            doc.set_font("Arial", "B", 10)  
            doc.cell(30, 6, "From:", border=0)
            doc.set_font("Arial", size=10)
            doc.multi_cell(0, 6, message["from"], border=0)
            
            doc.set_font("Arial", "B", 10)
            doc.cell(30, 6, "To:", border=0)
            doc.set_font("Arial", size=10)
            doc.multi_cell(0, 6, message["to"], border=0)
            
            doc.set_font("Arial", "B", 10)
            doc.cell(30, 6, "Subject:", border=0)
            doc.set_font("Arial", size=10)
            doc.multi_cell(0, 6, message["subject"], border=0)
            
            doc.ln(5)
            doc.set_font("Arial", "B", 10)
            doc.cell(0, 6, "Message Body:", ln=1, border=0)
            doc.ln(2)
        
        # Message body with automatic page breaks
        pdf.set_font("Arial", size=10)
        pdf.multi_cell(0, 5, message["body"])
        
        # Save PDF
        pdf_path = Path(export_folder) / f"{base_filename}.pdf"
        pdf.output(str(pdf_path))
        
        logger.info(f"Saved: {base_filename}")
        return True
        
    except Exception as e:
        logger.error(f"Error saving message files: {e}")
        return False

def main():
    """Main function"""
    parser = argparse.ArgumentParser(
        description="Search and export emails from IMAP servers",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --mailhost imap.gmail.com --username user@gmail.com --password pass --sender boss@company.com
  %(prog)s --env config.env --keywords "urgent project" --start-date 2023-01-01
  %(prog)s --mailhost mail.company.com --crypt starttls --recipient client@company.com
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
    parser.add_argument('--start-date', help='Start date (YYYY-MM-DD format)')
    parser.add_argument('--end-date', help='End date (YYYY-MM-DD format)')
    
    # Options
    parser.add_argument('--case-sensitive', action='store_true',
                       help='Enable case-sensitive matching')
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
    
    # Parse dates
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
        messages = search_messages(
            imap_conn,
            config.get('sender'),
            config.get('recipient'), 
            config.get('keywords'),
            start_date,
            end_date,
            args.case_sensitive
        )
        
        if not messages:
            logger.info("No matching messages found")
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
        
        # Export messages
        success_count = 0
        for message in messages:
            if save_message_files(message, export_folder):
                success_count += 1
        
        logger.info(f"Successfully exported {success_count}/{len(messages)} messages")
        
    finally:
        # Clean up connection
        try:
            imap_conn.close()
            imap_conn.logout()
        except:
            pass

if __name__ == "__main__":
    main()