# MSG → TXT / PDF Converter

Bulk convert Outlook `.msg` files into a single consolidated export.  
Outputs include:

- **TXT** — plain text file with full metadata and raw body.
- **Markdown** — structured output with metadata sections and raw body in grey code blocks.
- **PDF** — nicely formatted export generated from Markdown.

---

## Features

- Drag-and-drop multiple `.msg` files at once.
- Extracts **metadata**:
  - From / FromEmail
  - To, Cc, Bcc
  - Subject
  - Date (best-effort parsed from message, headers, or body)
  - Attachments
  - Raw Headers (optional)
- Includes **raw email body** (no formatting stripped) in fenced code blocks.
- Sort emails:
  - By Date (asc/desc)
  - As Uploaded
- Export options:
  - Combined TXT
  - Markdown
  - PDF (via Markdown → HTML → xhtml2pdf)

---

## Requirements

Python 3.9+ recommended.

Install dependencies:

```bash
pip install -r requirements.txt
