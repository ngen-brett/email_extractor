# IMAP Email Extractor with Privacy Mode

A professional Python tool for searching and extracting emails from IMAP servers with organized export to `.eml`, `.html`, and `.pdf` formats. Now includes privacy mode for email address redaction and multi-folder search capabilities.

[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)  
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 🆕 New Features

✅ **Privacy Mode** – Redact email addresses for compliance and confidentiality (e.g., `billgates@microsoft.com` → `b-------s@microsoft.com`)  
✅ **Letter Size PDFs** – US Letter (8.5" × 11") format with 0.5" margins for business use  
✅ **Enhanced PDF Quality** – Professional HTML-to-PDF rendering with proper typography  
✅ **Multi-Folder Search** – Search across all folders (INBOX, Sent, Drafts, etc.) or specific folders  
✅ **Verbose Mode** – Detailed progress tracking and message review output  

## Core Features

✅ **Flexible IMAP Connectivity** – Supports SSL, STARTTLS, and plaintext connections with auto-detection  
✅ **Multi-Provider Support** – Works with Gmail, Outlook, Yahoo, and corporate email servers  
✅ **Smart Search Filtering** – Filter by sender, recipient, keywords, and date ranges  
✅ **Professional PDF Export** – High-quality PDFs with proper formatting and Unicode support  
✅ **HTML Email Handling** – Preserves rich formatting from HTML emails  
✅ **Organized File Structure** – Creates structured export folders based on search criteria  
✅ **Dual Configuration** – Command-line arguments and `.env` file support  
✅ **Comprehensive Logging** – Detailed logging for troubleshooting and monitoring  
✅ **Enterprise-Ready** – Robust error handling suitable for compliance and archival use  

## Quick Start

1. **Clone and Install**  
   ```bash
   git clone https://github.com/yourusername/imap-email-extractor.git
   cd imap-email-extractor
   pip install -r requirements.txt
   ```

2. **Configure Authentication**  
   ```bash
   cp .env.example .env
   # Edit .env with your email server settings
   ```

3. **Run Your First Search**  
   ```bash
   python email_extractor_privacy.py --sender "boss@company.com"
   ```

4. **Privacy Mode Example**  
   ```bash
   python email_extractor_privacy.py --sender "client@company.com" --privacy
   ```

## Installation

### Requirements
- Python 3.7 or higher  
- IMAP-enabled email account  

### Dependencies
```bash
pip install -r requirements.txt
```

Required packages:
- `python-dotenv` – Environment variable management  
- `html2text` – HTML email content conversion  
- At least one PDF library:
  - `weasyprint` – **Recommended** (best quality, full CSS support)
  - `pdfkit` – Excellent quality (requires system `wkhtmltopdf`)
  - `xhtml2pdf` – Basic support (pure Python, no dependencies)

## Configuration

### Environment Variables (`.env`)
```env
MAILHOST=imap.gmail.com
MAILPORT=993
CRYPT=ssl

USERNAME=your-email@example.com
PASSWORD=your-app-password

SENDER=boss@company.com
RECIPIENT=client@example.com
KEYWORDS=urgent project
START_DATE=2023-01-01
END_DATE=2023-12-31
```

### Email Provider Settings

| Provider              | Hostname                   | Port | Encryption |
|-----------------------|----------------------------|------|------------|
| **Gmail**             | `imap.gmail.com`           | 993  | SSL        |
| **Outlook/Office 365**| `outlook.office365.com`    | 993  | SSL        |
| **Yahoo**             | `imap.mail.yahoo.com`      | 993  | SSL        |
| **Corporate IMAP**    | `mail.company.com`         | 143/993 | STARTTLS/SSL |

## Usage

### Basic Commands
```bash
# Search INBOX only (default)
python email_extractor_privacy.py --sender "manager@company.com"

# Search ALL folders
python email_extractor_privacy.py --sender "manager@company.com" --all-folders

# Privacy mode with redacted email addresses
python email_extractor_privacy.py --sender "client@company.com" --privacy
```

### Advanced Search Examples
```bash
# Multi-criteria search with date range
python email_extractor_privacy.py \
  --sender "boss@company.com" \
  --keywords "quarterly report" \
  --start-date 2023-01-01 \
  --end-date 2023-12-31 \
  --all-folders

# Verbose mode for detailed progress
python email_extractor_privacy.py \
  --recipient "team@company.com" \
  --privacy \
  --verbose

# Case-sensitive keyword search
python email_extractor_privacy.py \
  --keywords "Project Alpha" \
  --case-sensitive
```

### Full Argument Reference

| Argument         | Description                                      |
|------------------|--------------------------------------------------|
| `--mailhost`     | IMAP server hostname                             |
| `--mailport`     | IMAP server port                                 |
| `--crypt`        | `ssl`, `starttls`, or `none`                     |
| `--username`     | IMAP username                                    |
| `--password`     | IMAP password                                    |
| `--sender`       | Filter by sender (partial match)                 |
| `--recipient`    | Filter by recipient (partial match)              |
| `--keywords`     | Search keywords in subject and body              |
| `--start-date`   | Start date (`YYYY-MM-DD`), default all time      |
| `--end-date`     | End date (`YYYY-MM-DD`), default all time        |
| `--case-sensitive`| Enable case-sensitive matching                  |
| `--all-folders`  | Search all folders (default: INBOX only)         |
| `--privacy`      | **NEW:** Redact email addresses for privacy      |
| `--verbose`, `-v`| Show detailed progress and message review        |
| `--env`          | Path to `.env` file                              |
| `--export-dir`   | Export directory (default `./export`)            |

## Privacy Mode

### Email Address Redaction
When `--privacy` flag is used:
- **Email addresses are redacted**: `billgates@microsoft.com` → `b-------s@microsoft.com`
- **Files include `-redacted` suffix**: `message-redacted.pdf`
- **Privacy notice added to PDFs**: Clear indication of redacted content
- **All formats redacted**: EML, HTML, and PDF files

### Privacy Mode Examples
```bash
# Basic privacy mode
python email_extractor_privacy.py --sender "client@company.com" --privacy

# Privacy mode with multi-folder search
python email_extractor_privacy.py \
  --keywords "confidential" \
  --all-folders \
  --privacy \
  --verbose
```

### Use Cases for Privacy Mode
- **Compliance reviews** – Share emails without exposing personal information
- **Legal discovery** – Redact sensitive contact information
- **Training materials** – Use real emails without privacy concerns
- **External sharing** – Safe sharing with third parties

## Output Structure

### Standard Mode
```
export/
└── 2025-09-04_boss-company-com/
    ├── 2025-09-04_boss-company-com_client-example-com_project-deadline.eml
    ├── 2025-09-04_boss-company-com_client-example-com_project-deadline.html
    └── 2025-09-04_boss-company-com_client-example-com_project-deadline.pdf
```

### Privacy Mode
```
export/
└── 2025-09-04_boss-company-com/
    ├── 2025-09-04_boss-company-com_client-example-com_project-deadline-redacted.eml
    ├── 2025-09-04_boss-company-com_client-example-com_project-deadline-redacted.html
    └── 2025-09-04_boss-company-com_client-example-com_project-deadline-redacted.pdf
```

### File Naming Convention
- **Folder**: `{date}_{sender}_{recipient}_{keywords}` (sanitized)  
- **Files**: `{email-date}_{sender}_{recipient}_{first-16-subject}` (sanitized)  
- **Privacy**: Adds `-redacted` suffix before file extension

## PDF Features

### Professional Business Format
- **US Letter size** (8.5" × 11") optimized for business use
- **0.5" margins** on all sides for maximum content space
- **Professional typography** with proper font hierarchy
- **Print-optimized** CSS for clean printed output

### Advanced PDF Capabilities
- **Unicode Support**: Full international character support with DejaVu font fallback
- **HTML Preservation**: Maintains rich formatting from HTML emails
- **Smart Character Handling**: Normalizes smart quotes, dashes, and special characters
- **Multi-Page Support**: Automatic page breaks with headers and footers
- **Privacy Indicators**: Clear privacy notices in redacted documents

### PDF Library Options
| Library    | Quality | Setup | CSS Support | Recommended Use |
|------------|---------|-------|-------------|-----------------|
| WeasyPrint | Excellent | Medium | Full CSS3 | **Production** |
| pdfkit     | Excellent | Complex | Full CSS3 | Advanced users |
| xhtml2pdf  | Good | Easy | Basic CSS | Quick setup |

## Security & Best Practices

### Authentication Security
- Use **app passwords** instead of account passwords
- Enable **2FA** on email accounts
- Never commit `.env` files with credentials
- Use secure credential storage in production

### Privacy Compliance
- **Privacy mode** for GDPR/CCPA compliance
- **Redaction verification** in verbose mode
- **Audit trail** with comprehensive logging
- **Secure file handling** with proper permissions

### Production Deployment
- Use environment variables for credentials
- Implement proper logging and monitoring
- Regular security updates for dependencies
- Backup and retention policies for exports

## Troubleshooting

### Common Issues

**PDF Generation Errors:**
```bash
# Install recommended PDF library
pip install weasyprint

# Or fallback option
pip install xhtml2pdf
```

**IMAP Connection Issues:**
```bash
# Test connection settings
python email_extractor_privacy.py --verbose --sender test
```

**Unicode Character Issues:**
- Privacy mode automatically handles problematic characters
- WeasyPrint provides best Unicode support
- Check email encoding settings

### Verbose Mode Diagnostics
Use `--verbose` for detailed troubleshooting:
```bash
python email_extractor_privacy.py --sender "test@example.com" --verbose
```

Shows:
- PDF library status and capabilities
- Folder access and message counts
- Individual message processing details
- Privacy redaction confirmation
- File save operations and paths

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/privacy-enhancements`)
3. Commit your changes (`git commit -am 'Add privacy features'`)
4. Push to the branch (`git push origin feature/privacy-enhancements`)
5. Open a Pull Request

## License

MIT License – see the [LICENSE](LICENSE.md) file for details.

## Changelog

### v2.0.0 (2025-09-04)
- ✅ **Added Privacy Mode** with email address redaction
- ✅ **US Letter PDF format** with 0.5" margins
- ✅ **Enhanced HTML-to-PDF** rendering quality
- ✅ **Multi-folder search** capabilities
- ✅ **Verbose progress mode** for detailed output
- ✅ **Improved error handling** and diagnostics

### v1.0.0 (2025-09-03)
- ✅ Initial release with IMAP search and export
- ✅ Multi-format output (EML, HTML, PDF)
- ✅ Basic search filtering and date ranges
