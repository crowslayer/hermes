# 📧 Hermes - Email Automation with Excel + Outlook

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)
[![Version](https://img.shields.io/badge/version-1.0.3-green.svg)](CHANGELOG.md)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)
[![Windows](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue.svg)]()
[![VBA](https://img.shields.io/badge/Language-VBA-red.svg)]()

** Excel + Outlook automation for sending personalized emails with templates, dynamic variables and attachments**

<p align="center">
  <a href="README.md">🇪🇸 Español</a> | <strong>🇬🇧 English</strong>
</p>

---

## 📋 Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Screenshots](#screenshots)
- [Requirements](#requirements)
- [Excel BASE Sheet Structure](#excel-base-sheet-structure)
- [Dynamic Variables](#dynamic-variables)
- [OFT Templates](#oft-templates)
- [Execution Modes](#execution-modes)
- [Installation & Setup](#installation--setup)
- [Code Architecture](#code-architecture)
- [Processing Flow](#processing-flow)
- [Complete Practical Example](#complete-practical-example)
- [FAQ](#faq)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)
- [Author](#author)

---

## Overview

Hermes automates personalized email sending using **Excel + Outlook**, with full support for:

- Outlook **.OFT** templates
- Automatic attachments based on filename matching rules
- Unlimited dynamic variables from Excel columns
- Configurable delay between bulk emails
- Automatic processing status logging

Designed for enterprise workloads: accounting, billing, collections, administration, support, and invoicing.

---

## Key Features

| Feature | Description |
|---------|-------------|
| ✅ Multiple .OFT templates | Use different templates per row |
| ✅ Unlimited dynamic variables | Any column becomes a template variable |
| ✅ Smart attachments | Auto-match files by base filename |
| ✅ Safe sending | Configurable delay between emails |
| ✅ Preview mode | Review drafts before sending |
| ✅ Auto-logging | Status and timestamp columns auto-created |
| ✅ High performance | Handles thousands of rows without degradation |

---

## Screenshots

| Main Menu | File Selection | Results |
|-----------|----------------|---------|
| ![Main](docs/images/main.png) | ![File Dialog](docs/images/file_dialog.png) | ![Results](docs/images/result_simple.png) |

---

## Requirements

| Requirement | Details |
|-------------|---------|
| OS | Windows 10 / 11 |
| Software | Microsoft Excel + Microsoft Outlook |
| Config | Macros enabled, COM automation permissions |

---

## Excel BASE Sheet Structure

| Column | Content | Example |
|--------|---------|---------|
| A | ATTACHMENTS (separated by `\|`) | `invoice01\|certificate` |
| B | TO | `client@mail.com` |
| C | CC | `manager@company.com` |
| D | BCC | `auditor@company.com` |
| E | Subject | Monthly Invoice |
| F | Template | `invoice_client.oft` |
| G→∞ | Dynamic variables | NAME, RFC, TOTAL, DATE, etc. |

**Auto-generated columns:**
- Processing status (`Sent`, `Display`, `Error`)
- Processing timestamp

---

## Dynamic Variables

**Step 1:** Define column headers as variable names:
```
NAME | DATE | RFC | TOTAL | NOTE
```

**Step 2:** Use them in your .OFT template (HTML):
```html
Hello [NAME], your invoice from [DATE] was issued for [TOTAL].
```

**Result:** The engine replaces ALL variables automatically.

---

## OFT Templates

Save templates from Outlook:

**File → Save As → Outlook Template (.oft)**

Templates support full HTML: images, styles, tables, signatures, etc.

Example variables:
```
[NOMBRE]
[TOTAL]
[DEADLINE]
[PRODUCT]
```

---

## Execution Modes

### 1. Preview Mode
Processes each row and shows the email draft. User sends manually.
- Useful for verifying variable replacement
- Interactive pagination between pages

### 2. Auto-Send Mode
Prompts for delay (in seconds) between emails.
- Automatically sends all emails
- Built-in rate limiting

---

## Installation & Setup

1. Open Excel
2. Press `ALT + F11` (Visual Basic Editor)
3. Menu: *File → Import File...*
4. Import all `.bas` files from `/src`
5. Open your Excel file with the **BASE** sheet
6. Run macros:
   - `ProcessDisplay` (preview mode)
   - `ProcessSend` (auto-send mode)

---

## Code Architecture

| Module | Role |
|--------|------|
| **hermes.bas** | Main logic, flow control |
| **ProcessRow.bas** | Row processing (attachments, template, sending) |
| **Utils.bas** | Variable building and application |
| **Variables.bas** | Generic reusable functions |

**Benefits:**
- Easy maintenance
- Extend without breaking existing functions
- GitHub collaboration ready
- Future CI/CD integration

---

## Processing Flow

![General Flow](docs/diagramas/FlujoGeneral/FlujoGeneral.svg)

![Processing Flow](docs/diagramas/Procesar/Procesar.svg)

---

## Complete Practical Example

### 1. Excel Data:

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| invoice01 | client@x.com | cc@x.com | | Invoice | invoice.oft | NAME | TOTAL |
| invoice02 | client2@x.com | | | Invoice | invoice.oft | Luis | 1200 |

### 2. Template:
```html
Hello [NAME],

Your attached invoice has a total of [TOTAL].

Best regards.
```

### 3. Files in folder:
```
invoice01.pdf
invoice02.pdf
```

### 4. Result:
- Email sent or previewed
- Template with variables replaced
- Correct attachments attached
- Status marked in new columns

---

## FAQ

**Can I use images in templates?**
Yes. Outlook handles full HTML.

**Does Outlook need to be open?**
No. It opens automatically if needed.

**Does it work on Mac?**
No. Outlook COM automation only works on Windows.

**Can I send thousands of emails?**
Yes, but use delay between emails to avoid blocking.

---

## Troubleshooting

### Outlook blocks sending
Check: `File → Trust Center → Programmatic Access Settings`

### Template not found
Verify `.oft` extension and file path.

### Attachments not included
Ensure the **base filename** matches (without extension).

---

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

---

## License

This project is licensed under **GNU AGPLv3**. See [LICENSE](LICENSE) for details.

---

## Author

**Alex Herrera**
📧 crowslayer@gmail.com

---

## Support

If you find this project useful, consider supporting it:

<a href="https://www.paypal.com/donate/?hosted_button_id=3VLCPQZWUGACS">
  <img src="https://img.shields.io/badge/Donate-PayPal-green.svg" alt="Donate via PayPal">
</a>
