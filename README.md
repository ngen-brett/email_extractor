# IMAP Email Extractor

A professional Python tool for searching and extracting emails from IMAP servers with organized export to both `.eml` and `.pdf` formats.

[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

✅ **Flexible IMAP Connectivity** - Supports SSL, STARTTLS, and plaintext connections with auto-detection  
✅ **Multi-Provider Support** - Works with Gmail, Outlook, Yahoo, and corporate email servers  
✅ **Smart Search Filtering** - Filter by sender, recipient, keywords, and date ranges  
✅ **Professional PDF Export** - Multi-page PDFs with unbreakable headers and automatic pagination  
✅ **HTML Email Handling** - Converts HTML email content to readable plain text  
✅ **Organized File Structure** - Creates structured export folders based on search criteria  
✅ **Dual Configuration** - Command-line arguments and `.env` file support  
✅ **Comprehensive Logging** - Detailed logging for troubleshooting and monitoring  
✅ **Enterprise-Ready** - Robust error handling suitable for compliance and archival use

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
   python email_extractor.py --sender "boss@company.com"
   ```

## Installation

### Requirements
- Python 3.7 or higher
- IMAP-enabled email account

### Dependencies
```bash
pip install -r requirements.txt
```

**Required packages:**
- `python-dotenv` - Environment variable management
- `fpdf2` - Advanced PDF generation with multi-page support
- `html2text` - HTML email content conversion

## Configuration

### Environment Variables (.env)

Create a `.env` file for your email server configuration:

```env
# IMAP Server Configuration
MAILHOST=imap.gmail.com
MAILPORT=993
CRYPT=ssl

# Authentication (Required)
USERNAME=your-email@example.com
PASSWORD=your-app-password

# Search Criteria (Optional)
SENDER=sender@company.com
RECIPIENT=recipient@company.com
KEYWORDS=urgent project deadline
START_DATE=2023-01-01
END_DATE=2023-12-31
```

### Email Provider Settings

| Provider | Hostname | Port | Encryption |
|----------|----------|------|------------|
| **Gmail** | `imap.gmail.com` | 993 | SSL |
| **Outlook/Office 365** | `outlook.office365.com` | 993 | SSL |
| **Yahoo** | `imap.mail.yahoo.com` | 993 | SSL |
| **Corporate IMAP** | `mail.company.com` | 143/993 | STARTTLS/SSL |

## Usage

### Basic Commands

```bash
# Search by sender
python email_extractor.py --sender "manager@company.com"

# Search by keywords
python email_extractor.py --keywords "project deadline"

# Search by recipient
python email_extractor.py --recipient "team@company.com"

# Date range search
python email_extractor.py --start-date 2023-11-01 --end-date 2023-11-30
```

### Advanced Usage

```bash
# Multiple search criteria
python email_extractor.py \
    --sender "project-manager@company.com" \
    --recipient "development-team@company.com" \
    --keywords "sprint review" \
    --start-date 2023-12-01

# Case-sensitive search
python email_extractor.py --sender "Boss@company.com" --case-sensitive

# Custom IMAP server
python email_extractor.py \
    --mailhost mail.company.com \
    --mailport 143 \
    --crypt starttls \
    --username your-username \
    --password your-password \
    --keywords "quarterly report"

# Using custom .env file
python email_extractor.py --env production.env --sender "client@company.com"
```

### Gmail Setup Example

```bash
# Gmail with app password (recommended)
python email_extractor.py \
    --mailhost imap.gmail.com \
    --username your-email@gmail.com \
    --password your-16-digit-app-password \
    --keywords "invoice"
```

## Command Line Arguments

| Argument | Description | Example |
|----------|-------------|---------|
| `--mailhost` | IMAP server hostname | `imap.gmail.com` |
| `--mailport` | IMAP server port | `993` |
| `--crypt` | Encryption type | `ssl`, `starttls`, `none` |
| `--username` | IMAP username | `user@company.com` |
| `--password` | IMAP password | `app-password-123` |
| `--sender` | Filter by sender (partial match) | `boss@company.com` |
| `--recipient` | Filter by recipient (partial match) | `team@company.com` |
| `--keywords` | Search keywords in subject/body | `urgent project` |
| `--start-date` | Start date (YYYY-MM-DD) | `2023-01-01` |
| `--end-date` | End date (YYYY-MM-DD) | `2023-12-31` |
| `--case-sensitive` | Enable case-sensitive matching | (flag) |
| `--env` | Path to .env file | `.env`, `config.env` |
| `--export-dir` | Export directory | `./exports` |

## Output Structure

The tool creates organized export folders and files:

```
export/
└── 2023-09-04_boss-company-com_urgent/
    ├── 2023-09-04_boss-company-com_client-example-com_project-deadline-.eml
    ├── 2023-09-04_boss-company-com_client-example-com_project-deadline-.pdf
    ├── 2023-09-03_manager-company-com_team-company-com_status-update.eml
    └── 2023-09-03_manager-company-com_team-company-com_status-update.pdf
```

### Naming Conventions

**Folder Structure:**
- Format: `{search-date}_{sender}_{recipient}_{keywords}`
- Only includes provided search criteria
- Non-alphanumeric characters → hyphens
- Fields separated by underscores

**File Naming:**
- Format: `{email-date}_{sender}_{recipient}_{first-16-subject-chars}`
- Both `.eml` (original) and `.pdf` (formatted) versions
- Filesystem-safe character sanitization

## PDF Features

Generated PDFs include:

- 📄 **Professional Headers** - Email subject in page headers
- 🔒 **Unbreakable Sections** - Email metadata stays together across pages  
- 📖 **Multi-Page Support** - Automatic page breaks for long emails
- 🔢 **Page Numbers** - Footer pagination for easy navigation
- 🎨 **Text Formatting** - Preserves email structure and readability
- 🌐 **HTML Conversion** - Converts HTML emails to readable plain text

## Security & Best Practices

### Authentication
- **Use app passwords** instead of account passwords when available
- **Enable 2FA** on email accounts for additional security
- **Never commit** `.env` files with credentials to version control

### Email Provider Setup

**Gmail:**
1. Enable 2-Factor Authentication
2. Generate App Password: Google Account → Security → App passwords
3. Use the 16-digit app password in the script

**Outlook/Office 365:**
1. Enable IMAP in Outlook settings
2. Use account password or app password (if required by organization)
3. Some corporate accounts may require additional authentication

**Corporate Email:**
- Contact IT department for IMAP server details
- Verify firewall rules allow IMAP connections
- Test connection parameters before bulk operations

## Error Handling & Troubleshooting

### Common Issues

**🔐 Authentication Failed**
- Verify username/password credentials
- Check if 2FA requires app password generation
- Ensure IMAP access is enabled in email account settings

**🌐 Connection Timeout**
- Verify server hostname and port numbers
- Check corporate firewall/network restrictions
- Try different encryption methods (SSL vs STARTTLS)

**📭 No Messages Found**
- Review search criteria (check case sensitivity settings)
- Verify date range format (YYYY-MM-DD)
- Test with broader search terms first

**💾 PDF Generation Errors**
- Ensure sufficient disk space in export directory
- Check file system permissions
- Verify special characters in email content are handled

### Debug Mode

Enable verbose logging by setting the log level:

```python
logging.basicConfig(level=logging.DEBUG)
```

## Use Cases

### Compliance & Legal
- **Email Discovery** - Extract communications for legal proceedings
- **Compliance Audits** - Archive emails matching specific criteria
- **Regulatory Reporting** - Export emails for regulatory submissions

### Business Operations  
- **Project Communication** - Gather all emails related to specific projects
- **Client Correspondence** - Extract client communication threads
- **Contract Management** - Find emails containing contract-related keywords

### IT & Security
- **Incident Response** - Extract emails during security investigations
- **Backup & Archive** - Create searchable email archives
- **Migration Planning** - Export emails before system migrations

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- 📚 **Documentation** - Comprehensive README and inline code comments
- 🐛 **Issues** - Report bugs via GitHub Issues
- 💡 **Feature Requests** - Suggest improvements via GitHub Issues

## Changelog

### v1.0.0
- Initial release with core IMAP extraction functionality
- SSL/STARTTLS/plaintext connection support
- Advanced PDF generation with unbreakable sections
- Comprehensive search and filtering capabilities
- Professional file organization and naming

---

⚠️ **Security Notice**: This tool handles email credentials. Always use app passwords, enable 2FA, and follow your organization's security policies when deploying in production environments.
