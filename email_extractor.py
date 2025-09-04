#!/usr/bin/env python3
"""
IMAP Email Search and Export Tool
Connects to IMAP servers, searches messages, and exports to .eml and .pdf formats
Enhanced: Multi-folder search, HTML-to-PDF rendering for professional output
Updated: Letter size paper with 0.5" margins for optimal US business format
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
import html2text
from pathlib import Path
from html import escape

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Try to import HTML-to-PDF libraries with better error handling
PDF_ENGINE = None
WEASYPRINT_AVAILABLE = False
PDFKIT_AVAILABLE = False
XHTML2PDF_AVAILABLE = False

# Test WeasyPrint
try:
    import weasyprint
    from weasyprint import HTML, CSS
    WEASYPRINT_AVAILABLE = True
    PDF_ENGINE = "weasyprint"
    logger.info("WeasyPrint detected and available")
except ImportError as e:
    logger.debug(f"WeasyPrint not available: {e}")
except Exception as e:
    logger.debug(f"WeasyPrint import error: {e}")

# Test pdfkit
if not PDF_ENGINE:
    try:
        import pdfkit
        # Test if wkhtmltopdf is available
        pdfkit.configuration()
        PDFKIT_AVAILABLE = True
        PDF_ENGINE = "pdfkit"
        logger.info("pdfkit detected and available")
    except ImportError as e:
        logger.debug(f"pdfkit not available: {e}")
    except Exception as e:
        logger.debug(f"pdfkit error (may need wkhtmltopdf): {e}")

# Test xhtml2pdf
if not PDF_ENGINE:
    try:
        from xhtml2pdf import pisa
        XHTML2PDF_AVAILABLE = True
        PDF_ENGINE = "xhtml2pdf"
        logger.info("xhtml2pdf detected and available")
    except ImportError as e:
        logger.debug(f"xhtml2pdf not available: {e}")
    except Exception as e:
        logger.debug(f"xhtml2pdf error: {e}")

# Show what's available for debugging
logger.info(f"PDF Library Status: WeasyPrint={WEASYPRINT_AVAILABLE}, pdfkit={PDFKIT_AVAILABLE}, xhtml2pdf={XHTML2PDF_AVAILABLE}")

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

def extract_email_content(msg):
    """Extract both HTML and plain text content from email message"""
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
    
    # Return both HTML and plain text, with fallbacks
    if not html_content and text_content:
        # Convert plain text to basic HTML
        html_content = f"<pre style='white-space: pre-wrap; font-family: monospace;'>{escape(text_content)}</pre>"
    elif not text_content and html_content:
        # Convert HTML to plain text for search
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.ignore_images = True
        h.body_width = 0
        text_content = h.handle(html_content).strip()
    
    return {
        'html': html_content.strip() if html_content else "",
        'text': text_content.strip() if text_content else "No readable content found"
    }

def create_email_html(message, folder_name="UNKNOWN"):
    """Create a professional HTML representation of the email optimized for Letter size paper"""
    
    html_template = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Email: {subject}</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            max-width: none;
            margin: 0;
            padding: 0;
            line-height: 1.5;
            background-color: #ffffff;
            color: #212529;
            font-size: 11pt;
        }}
        
        .email-container {{
            background-color: white;
            padding: 20px;
            width: 100%;
            box-sizing: border-box;
        }}
        
        .email-header {{
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 4px;
            padding: 15px;
            margin-bottom: 20px;
        }}
        
        .header-table {{
            width: 100%;
            border-collapse: collapse;
        }}
        
        .header-table td {{
            padding: 6px 0;
            vertical-align: top;
            border-bottom: 1px solid #f1f3f4;
            font-size: 10pt;
        }}
        
        .header-table tr:last-child td {{
            border-bottom: none;
        }}
        
        .header-label {{
            font-weight: 600;
            color: #495057;
            width: 80px;
            padding-right: 12px;
        }}
        
        .header-value {{
            color: #212529;
            word-wrap: break-word;
        }}
        
        .email-subject {{
            font-size: 14pt;
            font-weight: 600;
            color: #1a73e8;
            margin-bottom: 12px;
            padding-bottom: 12px;
            border-bottom: 2px solid #e8f0fe;
            line-height: 1.3;
        }}
        
        .email-body {{
            background-color: #ffffff;
            padding: 18px;
            border-left: 3px solid #1a73e8;
            margin-top: 20px;
            border-radius: 0 3px 3px 0;
        }}
        
        .email-body pre {{
            white-space: pre-wrap;
            word-wrap: break-word;
            font-family: 'SF Mono', Monaco, 'Cascadia Code', 'Roboto Mono', Consolas, 'Courier New', monospace;
            font-size: 10pt;
            line-height: 1.4;
            margin: 0;
            color: #24292f;
        }}
        
        .email-body p {{
            margin-bottom: 0.8em;
            font-size: 10pt;
        }}
        
        .email-body a {{
            color: #1a73e8;
            text-decoration: none;
        }}
        
        .email-body a:hover {{
            text-decoration: underline;
        }}
        
        .footer {{
            margin-top: 30px;
            padding-top: 15px;
            border-top: 1px solid #e9ecef;
            color: #6c757d;
            font-size: 9pt;
            text-align: center;
        }}
        
        /* Optimize for Letter size paper (8.5" x 11") with 0.5" margins */
        @media print {{
            body {{ 
                margin: 0; 
                padding: 0; 
                background-color: white; 
                font-size: 10pt;
            }}
            .email-container {{ 
                padding: 0;
                margin: 0;
            }}
            .email-header {{
                break-inside: avoid;
                page-break-inside: avoid;
            }}
            .email-subject {{
                page-break-after: avoid;
            }}
            .email-body {{
                orphans: 3;
                widows: 3;
            }}
        }}
        
        /* Letter size (8.5" x 11") with 0.5" margins on all sides */
        @page {{
            margin: 0.5in;
            size: letter;
        }}
    </style>
</head>
<body>
    <div class="email-container">
        <div class="email-subject">{subject}</div>
        
        <div class="email-header">
            <table class="header-table">
                <tr>
                    <td class="header-label">Folder:</td>
                    <td class="header-value"><strong>{folder}</strong></td>
                </tr>
                <tr>
                    <td class="header-label">Date:</td>
                    <td class="header-value">{date}</td>
                </tr>
                <tr>
                    <td class="header-label">From:</td>
                    <td class="header-value">{from_addr}</td>
                </tr>
                <tr>
                    <td class="header-label">To:</td>
                    <td class="header-value">{to_addr}</td>
                </tr>
                {cc_row}
                {bcc_row}
            </table>
        </div>
        
        <div class="email-body">
            {body_content}
        </div>
        
        <div class="footer">
            Generated by IMAP Email Extractor on {export_date}
        </div>
    </div>
</body>
</html>"""

    # Prepare CC and BCC rows only if they exist
    cc_row = ""
    if message.get('cc') and message['cc'].strip():
        cc_row = f"""<tr>
            <td class="header-label">CC:</td>
            <td class="header-value">{escape(message['cc'])}</td>
        </tr>"""
    
    bcc_row = ""
    if message.get('bcc') and message['bcc'].strip():
        bcc_row = f"""<tr>
            <td class="header-label">BCC:</td>
            <td class="header-value">{escape(message['bcc'])}</td>
        </tr>"""
    
    # Use HTML content if available, otherwise format plain text
    if message['body_html'] and message['body_html'].strip():
        # For HTML emails, embed the HTML directly
        body_content = message['body_html']
    else:
        # For plain text emails, wrap in <pre> tags with better styling
        body_content = f"<pre>{escape(message['body_text'])}</pre>"
    
    return html_template.format(
        subject=escape(message['subject']),
        folder=escape(folder_name),
        date=escape(message['date_header']),
        from_addr=escape(message['from']),
        to_addr=escape(message['to']),
        cc_row=cc_row,
        bcc_row=bcc_row,
        body_content=body_content,
        export_date=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    )

def html_to_pdf(html_content, output_path, verbose=False):
    """Convert HTML content to PDF using available library with Letter size and 0.5" margins"""
    if not PDF_ENGINE:
        if verbose:
            print("  ❌ No PDF engine available")
        return False
    
    try:
        if PDF_ENGINE == "weasyprint":
            # WeasyPrint - best CSS support and output quality
            HTML(string=html_content).write_pdf(output_path)
            
        elif PDF_ENGINE == "pdfkit":
            # pdfkit - requires wkhtmltopdf system installation
            # Letter size: 8.5" x 11" with 0.5" margins
            options = {
                'page-size': 'Letter',
                'margin-top': '0.5in',
                'margin-right': '0.5in',
                'margin-bottom': '0.5in',
                'margin-left': '0.5in',
                'encoding': "UTF-8",
                'no-outline': None,
                'enable-local-file-access': None,
                'print-media-type': None,
                'disable-smart-shrinking': None
            }
            pdfkit.from_string(html_content, output_path, options=options)
            
        elif PDF_ENGINE == "xhtml2pdf":
            # xhtml2pdf - pure Python but limited CSS support
            with open(output_path, "w+b") as result_file:
                pisa_status = pisa.CreatePDF(html_content, dest=result_file)
                if pisa_status.err:
                    logger.error(f"xhtml2pdf error: {pisa_status.err}")
                    return False
        
        if verbose:
            print(f"  ✅ Generated PDF using {PDF_ENGINE} (Letter size, 0.5\" margins)")
        return True
        
    except Exception as e:
        logger.error(f"PDF generation failed with {PDF_ENGINE}: {e}")
        if verbose:
            print(f"  ❌ PDF generation error: {e}")
        return False

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
                
                # Extract both HTML and text content
                content = extract_email_content(email_message)
                
                # Verbose output - show message being reviewed
                if verbose:
                    print(f"\n    [{i}/{len(ids)}] Reviewing in {folder_name}:")
                    print(f"      Date: {date_header}")
                    print(f"      From: {from_addr}")
                    print(f"      To: {to_addr}")
                    print(f"      Subject: {subject}")
                    print(f"      Body preview: {content['text'][:100]}{'...' if len(content['text']) > 100 else ''}")
                
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
                keyword_match = matches(keywords, f"{subject} {content['text']}")
                
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
                        "cc": cc_addr,
                        "bcc": bcc_addr,
                        "subject": subject,
                        "date": email_date,
                        "date_header": date_header,
                        "raw": raw_email,
                        "body_text": content['text'],
                        "body_html": content['html'],
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
    """Save message as .eml, .html, and .pdf files"""
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
        
        # Generate professional HTML
        html_content = create_email_html(message, message.get('folder', 'UNKNOWN'))
        
        # Save .html file for reference/debugging
        html_path = Path(export_folder) / f"{base_filename}.html"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        
        if verbose:
            print(f"  ✅ Saved HTML: {html_path}")
        
        # Convert HTML to PDF if engine available
        if PDF_ENGINE:
            pdf_path = Path(export_folder) / f"{base_filename}.pdf"
            if html_to_pdf(html_content, str(pdf_path), verbose):
                if verbose:
                    print(f"  ✅ Saved PDF: {pdf_path}")
            else:
                logger.warning(f"Failed to generate PDF for {base_filename}")
        else:
            if verbose:
                print(f"  ⚠️  PDF skipped (no engine available)")
        
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
        description="Search and export emails from IMAP servers with professional PDF output (Letter size, 0.5\" margins)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --mailhost imap.gmail.com --username user@gmail.com --password pass --sender boss@company.com
  %(prog)s --env config.env --keywords "urgent project" --start-date 2023-01-01 --all-folders
  %(prog)s --mailhost mail.company.com --crypt starttls --recipient client@company.com --verbose --all-folders

PDF Requirements:
  Install one of the following for PDF generation:
  - pip install weasyprint (recommended)
  - pip install pdfkit (requires wkhtmltopdf system installation)
  - pip install xhtml2pdf (basic support)

PDF Format: US Letter size (8.5" x 11") with 0.5" margins on all sides
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
    
    # Show detailed PDF engine status
    print(f"\n📋 PDF Engine Status:")
    print(f"  WeasyPrint: {'✅ Available' if WEASYPRINT_AVAILABLE else '❌ Not available'}")
    print(f"  pdfkit: {'✅ Available' if PDFKIT_AVAILABLE else '❌ Not available'}")
    print(f"  xhtml2pdf: {'✅ Available' if XHTML2PDF_AVAILABLE else '❌ Not available'}")
    print(f"  Selected engine: {PDF_ENGINE or 'None'}")
    print(f"  PDF format: US Letter (8.5\" x 11\") with 0.5\" margins")
    
    # Check PDF generation capability
    if not PDF_ENGINE:
        print("\n⚠️  WARNING: No PDF generation library working properly!")
        print("Install one of the following for PDF output:")
        print("  pip install weasyprint      (recommended)")
        print("  pip install pdfkit          (requires wkhtmltopdf)")
        print("  pip install xhtml2pdf       (basic support)")
        print("\nContinuing with .eml and .html output only...\n")
    else:
        print(f"\n✅ PDF generation available using {PDF_ENGINE}\n")
    
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
        print(f"  Export directory: {args.export_dir}")
        print(f"  PDF format: Letter size, 0.5\" margins")
        print(f"  PDF engine: {PDF_ENGINE or 'None'}\n")
    
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
        
        # Summary of what was created
        output_formats = [".eml (original)", ".html (formatted)"]
        if PDF_ENGINE:
            output_formats.append(".pdf (Letter size, 0.5\" margins)")
        
        if args.verbose:
            print(f"\n✅ Export completed: {success_count}/{len(messages)} messages saved successfully")
            print(f"Files saved per message: {', '.join(output_formats)}")
        
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
