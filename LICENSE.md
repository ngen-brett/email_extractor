# MIT License

**Copyright (c) 2025 NGEN Networks, LLC, and IMAP Email Extractor Contributors**

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

---

## Third-Party Dependencies and License Compatibility

This software depends on the following third-party libraries, each with their own licensing terms:

### Compatible Dependencies (Permissive Licenses)

| Library | License | Compatibility |
|---------|---------|---------------|
| **python-dotenv** | BSD 3-Clause | ✅ Compatible with MIT |
| **weasyprint** | BSD 3-Clause | ✅ Compatible with MIT |
| **pdfkit** | MIT-compatible | ✅ Compatible with MIT |
| **xhtml2pdf** | Apache 2.0-compatible | ✅ Compatible with MIT |

### GPL Dependency

| Library | License | Usage Model |
|---------|---------|-------------|
| **html2text** | GPL 3.0 or later | 🔄 Dynamic dependency (pip install) |

**Important:** `html2text` is used as an external dependency installed via pip. This is generally considered acceptable usage that does not trigger GPL copyleft requirements for the main application.

---

## Usage Rights Summary

### ✅ **Permitted Uses**

- **Commercial Use** - Use in commercial applications without licensing fees
- **Modification** - Modify source code for your needs and create derivative works
- **Distribution** - Distribute original or modified versions, including in proprietary software
- **Private Use** - Use for internal business purposes with no disclosure requirements

### 📋 **Requirements**

- **Attribution Required** - Include this license file with distributions and retain copyright notices
- **License Preservation** - Include the full MIT license text in all copies or substantial portions

### 🚫 **No Additional Obligations**

- **No Source Disclosure** - Modifications can remain proprietary
- **No Same License** - Derivative works can use different licenses (but must retain MIT attribution)
- **No Patent Retaliation** - No patent licensing restrictions

---

## GPL Dependency Considerations

### Dynamic Linking Model

This software uses `html2text`, which is licensed under GPL 3.0. Key considerations:

1. **Separate Installation** - `html2text` is installed as an independent package via pip
2. **No Code Mixing** - The main codebase does not include `html2text` source code
3. **Runtime Dependency** - Used only at runtime through Python imports
4. **No Modifications** - `html2text` is used as-is without modifications

### Alternative Options

Users concerned about GPL compatibility may substitute `html2text` with:
- **BeautifulSoup** with custom text extraction
- **Custom HTML parsing** solutions
- **Other permissively-licensed** HTML-to-text converters

### Compliance Guidelines

- ✅ **Using as pip dependency** - No GPL obligations for your code
- ❌ **Modifying html2text source** - Would require GPL 3.0 licensing for modifications
- ❌ **Statically bundling html2text** - Would trigger GPL copyleft requirements

---

## Patent Grant

Contributors grant users a license to any patents they own that are necessarily infringed by the software, subject to the same conditions as this MIT License.

---

## Disclaimer of Warranties

**THE SOFTWARE IS PROVIDED "AS-IS" WITHOUT WARRANTY.** The authors disclaim all warranties including:

- Fitness for a particular purpose
- Non-infringement of third-party rights
- Security or reliability guarantees
- Compliance with specific regulations

### User Responsibilities

Users are responsible for:
- Testing in their environment
- Compliance with applicable laws and regulations
- Security assessment for their specific use case
- Proper handling of email data and privacy regulations (GDPR, HIPAA, etc.)

---

## Contribution Policy

By contributing to this project, you agree that:

1. Your contributions will be licensed under this same MIT License
2. You have the legal right to make the contribution
3. You understand the GPL dependency implications
4. You provide contributions without warranty

---

## Commercial Support

For licensing questions, commercial support, or enterprise inquiries:

- Open an issue on the project repository
- Contact the maintainers directly
- Consult with legal counsel for complex licensing scenarios

---

**License effective date:** September 4, 2025  
**Last updated:** September 4, 2025
