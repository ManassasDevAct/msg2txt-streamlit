import io
import os
import re
import tempfile
from datetime import datetime
from typing import Optional, Dict, List

import streamlit as st
from dateutil import parser as dtparser
import extract_msg
import markdown2
from xhtml2pdf import pisa


# ── Basics ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="MSG → One TXT", page_icon="📥", layout="centered")
st.title("Bulk convert .msg to a single .txt")
st.caption("Drop many Outlook .msg files. Get one combined .txt with full metadata.")

# ── Helpers ────────────────────────────────────────────────────────────────────
def ensure_str(x) -> str:
    if x is None:
        return ""
    if isinstance(x, bytes):
        try:
            return x.decode("utf-8", "replace")
        except Exception:
            return x.decode("latin-1", "replace")
    return str(x)

def safe_join(sep: str, items) -> str:
    return sep.join(ensure_str(i) for i in items)

def safe_filename(name: str, max_len: int = 180) -> str:
    name = re.sub(r'[<>:"/\\|?*\n\r\t]+', "_", ensure_str(name)).strip().strip(".")
    return name[:max_len] if len(name) > max_len else name

def stringify_addrs(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (list, tuple, set)):
        return ", ".join(ensure_str(x).strip() for x in v if ensure_str(x).strip())
    return ensure_str(v)

def try_parse_datetime(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    try:
        return dtparser.parse(s)
    except Exception:
        return None

def md_inline_escape(text: str) -> str:
    s = ensure_str(text)
    # light escape so metadata prints neatly
    return s.replace("|", r"\|").replace("*", r"\*").replace("_", r"\_").replace("`", r"\`")


# ── Date extraction (compact + robust) ─────────────────────────────────────────
DATE_BODY_PATTERNS = [r'^\s*sent:\s*(.+)$', r'^\s*date:\s*(.+)$']
DATE_BODY_COMPILED = [re.compile(p, re.IGNORECASE | re.MULTILINE) for p in DATE_BODY_PATTERNS]

def coalesce(*vals) -> str:
    for v in vals:
        if v:
            s = ensure_str(v).strip()
            if s: return s
    return ""

def normalize_email_pair(display: str, email_field: str) -> tuple[str, str]:
    disp = ensure_str(display) or ""
    eml = ensure_str(email_field).strip()
    if not eml:
        m = re.search(r'<([^>]+@[^>]+)>', disp)
        if m: eml = m.group(1).strip()
    if not disp and eml: disp = eml
    disp = re.sub(r'\s*<[^>]+>\s*', '', disp).strip() or disp
    return disp, eml

def parse_date_from_headers(headers: str) -> Optional[str]:
    if not headers: return None
    lines = ensure_str(headers).splitlines()
    collected, cur = [], None
    for line in lines:
        if re.match(r'^[\t ]', line) and cur is not None:
            cur += " " + line.strip()
        else:
            if cur is not None: collected.append(cur)
            cur = line
    if cur is not None: collected.append(cur)
    for line in collected:
        if line.lower().startswith("date:"):
            return line.split(":", 1)[1].strip()
    return None

def parse_date_from_body(body: str) -> Optional[str]:
    if not body: return None
    for pat in DATE_BODY_COMPILED:
        m = pat.search(body)
        if m:
            cand = m.group(1).strip()
            cand = re.split(r'\b(subject|to|from):', cand, flags=re.IGNORECASE)[0].strip()
            if cand: return cand
    return None

def best_effort_parse_datetime(*candidates: Optional[str]) -> tuple[str, str]:
    for c in candidates:
        if not c: continue
        try:
            dt = dtparser.parse(c)
            return dt.isoformat(), ensure_str(c)
        except Exception:
            pass
    return "", ensure_str(coalesce(*candidates))


# ── MSG parsing ────────────────────────────────────────────────────────────────
def read_msg_from_bytes(data: bytes, debug: bool = False) -> Dict[str, str]:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
        tmp.write(data); tmp_path = tmp.name
    try:
        msg = extract_msg.Message(tmp_path)

        sender_display = ensure_str(getattr(msg, "sender", ""))
        sender_email   = ensure_str(getattr(msg, "senderemail", ""))
        from_display, from_email = normalize_email_pair(sender_display, sender_email)

        headers_raw = ensure_str(getattr(msg, "headers", "") or "")
        body_text   = ensure_str(getattr(msg, "body", "") or "")

        to_val  = stringify_addrs(getattr(msg, "to", None))
        cc_val  = stringify_addrs(getattr(msg, "cc", None))
        bcc_val = stringify_addrs(getattr(msg, "bcc", None))

        att_names = []
        for att in (getattr(msg, "attachments", []) or []):
            longn = ensure_str(getattr(att, "longFilename", "") or "")
            shortn = ensure_str(getattr(att, "shortFilename", "") or "")
            att_names.append(longn or shortn)
        att_str = ", ".join(a for a in att_names if a)

        cand_date = coalesce(
            getattr(msg, "date", None),
            getattr(msg, "clientSubmitTime", None),
            getattr(msg, "messageDeliveryTime", None),
            getattr(msg, "lastModificationTime", None),
            getattr(msg, "creationTime", None),
        )
        cand_header = parse_date_from_headers(headers_raw)
        cand_body   = parse_date_from_body(body_text)
        iso_date, raw_used = best_effort_parse_datetime(cand_date, cand_header, cand_body)

        meta = {
            "OriginalFilename": os.path.basename(tmp_path),
            "From": from_display,
            "FromEmail": from_email,
            "To": to_val,
            "Cc": cc_val,
            "Bcc": bcc_val,
            "Subject": ensure_str(getattr(msg, "subject", "") or ""),
            "Date": ensure_str(iso_date),
            "DateRaw": ensure_str(raw_used),
            "HeadersRaw": headers_raw,
            "Body": body_text,               # RAW body; we will display as-is
            "AttachmentNames": att_str,
        }

        if debug:
            meta["_date_debug"] = {
                "msg.date": ensure_str(getattr(msg, "date", None)),
                "clientSubmitTime": ensure_str(getattr(msg, "clientSubmitTime", None)),
                "messageDeliveryTime": ensure_str(getattr(msg, "messageDeliveryTime", None)),
                "lastModificationTime": ensure_str(getattr(msg, "lastModificationTime", None)),
                "creationTime": ensure_str(getattr(msg, "creationTime", None)),
                "headers.Date": ensure_str(cand_header or ""),
                "body_sent_line": ensure_str(cand_body or ""),
                "raw_used": ensure_str(raw_used),
                "iso": ensure_str(iso_date),
            }
        return meta
    finally:
        try: os.unlink(tmp_path)
        except Exception: pass


# ── Formatting (TXT/MD/PDF) ────────────────────────────────────────────────────
def format_record_txt(rec: Dict[str, str], idx: int, total: int) -> str:
    sep = "=" * 78
    body_raw = ensure_str(rec.get("Body", ""))
    lines = [
        sep,
        f"Email {idx} of {total}",
        sep,
        f"From: {ensure_str(rec.get('From',''))}",
        f"FromEmail: {ensure_str(rec.get('FromEmail',''))}",
        f"To: {ensure_str(rec.get('To',''))}",
        f"Cc: {ensure_str(rec.get('Cc',''))}",
        f"Bcc: {ensure_str(rec.get('Bcc',''))}",
        f"Subject: {ensure_str(rec.get('Subject',''))}",
        f"Date: {ensure_str(rec.get('Date') or rec.get('DateRaw') or '')}",
        f"AttachmentNames: {ensure_str(rec.get('AttachmentNames',''))}",
        "",
        "Headers:",
        ensure_str(rec.get("HeadersRaw","")).strip(),
        "",
        "Body:",
        body_raw.rstrip(),
        "",
    ]
    return safe_join("\n", lines)

def build_markdown(records: List[Dict[str, str]]) -> str:
    parts = []
    parts.append("# Email Export\n")
    parts.append(f"_Total emails_: **{len(records)}**\n")
    parts.append("---\n")
    for i, r in enumerate(records, start=1):
        parts.append(f"## Email {i}\n")
        parts.append(f"**From:** {md_inline_escape(r.get('From',''))}  ")
        if r.get("FromEmail"):
            parts.append(f"**FromEmail:** `{ensure_str(r.get('FromEmail',''))}`  ")
        parts.append(f"**To:** {md_inline_escape(r.get('To',''))}  ")
        if r.get("Cc"):  parts.append(f"**Cc:** {md_inline_escape(r.get('Cc',''))}  ")
        if r.get("Bcc"): parts.append(f"**Bcc:** {md_inline_escape(r.get('Bcc',''))}  ")
        parts.append(f"**Subject:** {md_inline_escape(r.get('Subject',''))}  ")
        parts.append(f"**Date:** `{ensure_str(r.get('Date') or r.get('DateRaw') or '')}`  ")
        if r.get("AttachmentNames"):
            parts.append(f"**Attachments:** {md_inline_escape(r.get('AttachmentNames',''))}  ")
        if r.get('OriginalFilename'):
            parts.append(f"**Source File:** `{ensure_str(r['OriginalFilename'])}`  ")

        # Optional headers in a collapsible block if present
        headers = ensure_str(r.get("HeadersRaw","")).strip()
        if headers:
            parts.append("\n<details>\n<summary><strong>Headers</strong></summary>\n\n```text")
            parts.append(headers)
            parts.append("```\n</details>\n")

        # RAW body in a grey box (fenced code block)
        parts.append("\n**Body**\n")
        parts.append("```text")
        parts.append(ensure_str(r.get("Body","")).rstrip())
        parts.append("```")

        parts.append('\n<div class="pagebreak"></div>\n')
        parts.append("---\n")
    return safe_join("\n", parts)

def markdown_to_pdf(md_text: str) -> bytes:
    html_body = markdown2.markdown(ensure_str(md_text), extras=["tables", "fenced-code-blocks"])
    html_full = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @page {{ size: Letter; margin: 0.6in; }}
  body {{ font-family: DejaVu Sans, Arial, Helvetica, sans-serif; font-size: 11pt; line-height: 1.35; }}
  h1, h2, h3, h4 {{ margin: 0.4em 0 0.2em; }}
  code, pre {{ font-family: "DejaVu Sans Mono", "Courier New", monospace; font-size: 10pt; white-space: pre-wrap; word-wrap: break-word; }}
  pre {{ border: 1px solid #ddd; padding: 8px; border-radius: 4px; background: #fafafa; }}
  .pagebreak {{ page-break-before: always; }}
  details > summary {{ cursor: pointer; margin: 0.2em 0; }}
  table {{ border-collapse: collapse; width: 100%; margin: 8px 0; font-size: 10pt; }}
  th, td {{ border: 1px solid #ddd; padding: 4px 6px; vertical-align: top; }}
  th {{ background: #f0f0f0; }}
</style>
</head>
<body>
{html_body}
</body>
</html>
    """.strip()
    out = io.BytesIO()
    pisa.CreatePDF(io.StringIO(ensure_str(html_full)), dest=out)
    return out.getvalue()

def sort_key(rec: Dict[str, str]):
    dt = try_parse_datetime(ensure_str(rec.get("Date") or rec.get("DateRaw")))
    subj = ensure_str(rec.get("Subject", ""))
    fn = ensure_str(rec.get("OriginalFilename", ""))
    return (0 if dt else 1, dt or datetime.min, subj, fn)


# ── UI Controls ────────────────────────────────────────────────────────────────
with st.expander("Options", expanded=False):
    sort_choice = st.selectbox("Order emails in output",
                               ["By Date (asc)", "By Date (desc)", "As Uploaded"], index=0)
    include_headers = st.checkbox("Include raw headers", value=True)
    include_body = st.checkbox("Include body", value=True)
    show_date_debug = st.checkbox("Show date debug info", value=False)
    make_pdf = st.checkbox("Prepare PDF export (Markdown → PDF)", value=True)

uploaded = st.file_uploader(
    "Drop .msg files here", type=["msg"], accept_multiple_files=True,
    help="Select many .msg files. They will be merged into one .txt file.",
)

ordered: List[Dict[str, str]] = []

# ── Main flow ──────────────────────────────────────────────────────────────────
if uploaded:
    st.info(f"Loaded {len(uploaded)} file(s). Parsing locally...")
    records, errors = [], []

    for f in uploaded:
        try:
            rec = read_msg_from_bytes(f.read(), debug=show_date_debug)
            rec["OriginalFilename"] = ensure_str(f.name)

            # Respect toggles
            if not include_headers:
                rec["HeadersRaw"] = ""
            if not include_body:
                rec["Body"] = ""

            records.append(rec)
        except Exception as e:
            errors.append(f"{f.name}: {e}")

    if errors:
        st.warning("Some files could not be parsed:")
        for e in errors:
            st.code(e)

    ordered = sorted(records, key=sort_key, reverse=("desc" in sort_choice.lower())) \
        if sort_choice.startswith("By Date") else records

    total = len(ordered)
    combined_parts = [format_record_txt(r, i, total) for i, r in enumerate(ordered, start=1)]
    combined_text = ensure_str(safe_join("\n", combined_parts)).encode("utf-8")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = safe_filename(f"merged_emails_{ts}.txt")

    # TXT download
    st.download_button(
        "Download combined TXT",
        data=combined_text,
        file_name=out_name,
        mime="text/plain",
        use_container_width=True,
    )

    # Markdown + PDF (with RAW body in grey box)
    if make_pdf and ordered:
        md_doc = build_markdown(ordered)
        st.download_button(
            "Download Markdown",
            data=md_doc.encode("utf-8"),
            file_name=out_name.replace(".txt", ".md"),
            mime="text/markdown",
            use_container_width=True,
        )
        try:
            pdf_bytes = markdown_to_pdf(md_doc)
            st.download_button(
                "Download PDF",
                data=pdf_bytes,
                file_name=out_name.replace(".txt", ".pdf"),
                mime="application/pdf",
                use_container_width=True,
            )
        except Exception as e:
            st.warning(f"PDF generation failed: {e}")

    with st.expander("Preview first email section"):
        if ordered:
            st.text(format_record_txt(ordered[0], 1, total)[:5000])
        else:
            st.write("No records parsed yet.")
else:
    st.write("Select one or more .msg files to enable conversion.")
