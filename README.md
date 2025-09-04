# IMAP Email Extractor

A professional Python tool for searching and extracting emails from IMAP servers with organized export to both `.eml` and `.pdf` formats. Now supports multi-folder search.

[![Python 3.7+](https://img.shields.io/badge/pythonicenseectivity** – Supports SSL, STARTTLS, and plaintext connections with auto-detection  
✅ **Multi-Provider Support** – Works with Gmail, Outlook, Yahoo, and corporate email servers  
✅ **Smart Search Filtering** – Filter by sender, recipient, keywords, and date ranges  
✅ **Multi-Folder Search** – Optionally search all folders (INBOX, Sent, Drafts, Spam, custom folders)  
✅ **Professional PDF Export** – Multi-page PDFs with unbreakable headers, Unicode support, and automatic pagination  
✅ **HTML Email Handling** – Converts HTML email content to readable plain text  
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
   python email_extractor_multi_folder.py --sender "boss@company.com"
   ```

## Installation

### Requirements
- Python 3.7 or higher  
- IMAP-enabled email account  

### Dependencies
```bash
pip install -r requirements.txt
```
- `python-dotenv` – Environment variable management  
- `fpdf2` – Advanced PDF generation with multi-page and Unicode support  
- `html2text` – HTML email content conversion  

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
python email_extractor_multi_folder.py --sender "manager@company.com"

# Search ALL folders
python email_extractor_multi_folder.py --sender "manager@company.com" --all-folders
```

### Search Criteria Examples
```bash
# Keywords and date range
python email_extractor_multi_folder.py --keywords "invoice" --start-date 2023-01-01

# Case-sensitive search
python email_extractor_multi_folder.py --sender "Boss@Company.com" --case-sensitive

# Multi-folder verbose search
python email_extractor_multi_folder.py \
  --recipient "team@company.com" \
  --all-folders \
  --verbose
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
| `--verbose`, `-v`| Show messages under review and detailed progress |
| `--env`          | Path to `.env` file                              |
| `--export-dir`   | Export directory (default `./export`)            |

## Output Structure

```
export/
└── 2025-09-04_boss-company-com/
    ├── 2025-09-04_boss-company-com_client-example-com_project-deadline-.eml
    ├── 2025-09-04_boss-company-com_client-example-com_project-deadline-.pdf
    ├── 2025-09-04_boss-company-com_Sent_report-summary-.eml
    └── 2025-09-04_boss-company-com_Sent_report-summary-.pdf
```

- **Folder Naming**: `{date}_{sender}_{recipient}_{keywords}`, sanitized  
- **File Naming**: `{email-date}_{sender}_{recipient}_{first-16-subject}`, sanitized  

## PDF Features

- **Unicode Support**: DejaVu Unicode font with Arial fallback  
- **Smart Character Handling**: Normalizes smart quotes, dashes, and removes zero-width characters  
- **Professional Formatting**: Headers, footers, page numbers  
- **Multi-Page Support**: Automatic page breaks  
- **Controlled Layout**: Fixed margins and conservative cell widths  

## Security & Best Practices

- Use **app passwords** and enable **2FA**  
- Never commit `.env` with credentials to version control  
- Always use **SSL/STARTTLS** when available  
- For production, consider secure credential storage  

## Contributing

1. Fork the repo  
2. Create a feature branch  
3. Commit your changes  
4. Open a Pull Request  

## License

Internal Use Only – see the [LICENSE](LICENSE) file.
