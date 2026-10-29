"""RFQ Data Processor - Python-only verification edition.

Preserves the Excel/PDF workflows without external AI APIs.
"""
import re
import csv
import html
import traceback
from collections import Counter
import streamlit.components.v1 as components
from collections import defaultdict
from copy import copy
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from PyPDF2 import PdfReader
try:
    import fitz  # PyMuPDF: layout-aware extraction for tabular RFQ PDFs
except ImportError:
    fitz = None

st.set_page_config(page_title="RFQ Data Processor", layout="wide", initial_sidebar_state="auto")
st.markdown("""<style>
.stApp {background:linear-gradient(160deg,#0d1220,#111827 52%,#0d1421)}
/* Dialog: restore a single reachable vertical scroll surface.
   Do not clip Streamlit's own dialog content or reserve a side gutter. */
[data-testid="stDialog"] [role="dialog"] {
  width:min(1280px, calc(100vw - 24px))!important;
  max-width:calc(100vw - 24px)!important;
  max-height:calc(100dvh - 24px)!important;
  overflow-y:auto!important;
  overflow-x:hidden!important;
  box-sizing:border-box!important;
  overscroll-behavior-y:contain;
}
[data-testid="stDialog"] [data-testid="stDialogContent"] {
  width:100%!important;
  min-width:0!important;
  max-width:100%!important;
  height:auto!important;
  max-height:none!important;
  overflow:visible!important;
  box-sizing:border-box!important;
  padding:1rem 1.1rem!important;
}
[data-testid="stDialog"] [data-testid="stDialogContent"] > div {
  min-width:0!important;
  width:100%!important;
  max-width:100%!important;
  box-sizing:border-box!important;
}
[data-testid="stDialog"] [data-testid="stVerticalBlock"] {gap:.55rem;min-width:0}
[data-testid="stDialog"] [data-testid="column"],
[data-testid="stDialog"] [data-testid="stHorizontalBlock"],
[data-testid="stDialog"] iframe {min-width:0!important;max-width:100%!important;box-sizing:border-box}
[data-testid="stDialog"] [data-testid="stSelectbox"] {min-width:0;overflow:visible!important}
[data-testid="stDialog"] [data-testid="stSelectbox"] [data-baseweb="select"] {overflow:visible!important}
[data-testid="stDialog"] [role="dialog"] {scroll-padding-bottom:14px}
/* Match native button and selectbox heights in review navigation. */
[data-testid="stDialog"] [data-testid="stSelectbox"] [data-baseweb="select"],
[data-testid="stDialog"] [data-testid="stSelectbox"] [data-baseweb="select"] > div,
[data-testid="stDialog"] .stButton button {min-height:42px!important}
.rfq-review-stat{background:#192436;border:1px solid #344158;border-radius:10px;padding:12px 14px;min-height:72px;box-sizing:border-box}
.rfq-stat-label{color:#dae2f0;font-size:14px;font-weight:600}
.rfq-stat-count{color:#f5f8fd;font-size:25px;font-weight:600;line-height:1.15;margin-top:3px}
.rfq-orange{color:#ffa742!important}.rfq-red{color:#ff6868!important}
.rfq-nav-label{height:23px;font-size:14px;font-weight:600;line-height:21px;margin:0;position:relative;top:-6px}
/* Make Streamlit's editable field visually match the read-only comparison. */
[data-testid="stDialog"] [data-testid="stTextArea"] textarea {
  background:#ffffff!important;
  color:#172033!important;
  border:1px solid #aebccd!important;
  border-radius:12px!important;
  font-family:ui-monospace,Consolas,monospace!important;
  font-size:12.5px!important;
  line-height:1.65!important;
  padding:14px!important;
  overflow-wrap:normal!important;
  word-break:normal!important;
  white-space:pre!important;
  overflow-x:auto!important;
  resize:none!important;
}
[data-testid="stDialog"] [data-testid="stTextArea"] textarea:focus {
  outline:1px solid #54769f!important;
  box-shadow:none!important;
}
[data-testid="stDialog"] .stButton button {white-space:nowrap;padding-left:.5rem;padding-right:.5rem}
.stButton button[kind="primary"] {background:#ef8545;border-color:#ef8545;color:white;border-radius:9px}
.stButton button {border-radius:9px;min-height:2.35rem}
[data-testid="stFileUploader"] {border-radius:12px}
[data-testid="stMetric"] {background:#192436;border:1px solid #344158;border-radius:10px;padding:8px 12px}

/* Compact no-format confirmation: override global full-width dialog rules. */
[data-testid="stDialog"] [role="dialog"]:has(#rfq-noformat-modal) {
  width:min(430px, calc(100vw - 32px))!important;
  max-width:min(430px, calc(100vw - 32px))!important;
  min-width:0!important;
  height:auto!important;
}
[data-testid="stDialog"] [role="dialog"]:has(#rfq-noformat-modal) [data-testid="stAlert"] {
  width:100%!important;
  min-width:0!important;
  max-width:100%!important;
  overflow-wrap:anywhere!important;
  box-sizing:border-box!important;
}
[data-testid="stDialog"] [role="dialog"]:has(#rfq-noformat-modal) [data-testid="stHorizontalBlock"] {
  justify-content:center!important;
  gap:12px!important;
}
[data-testid="stDialog"] [role="dialog"]:has(#rfq-noformat-modal) [data-testid="column"] {
  min-width:0!important;
  flex:0 1 150px!important;
  width:150px!important;
}

/* Only the Verify All dialog should be compact; not the main review dialog. */
[data-testid="stDialog"] [role="dialog"]:has(#rfq-confirm-modal) {
  width:min(490px, calc(100vw - 32px))!important;
  max-width:min(490px, calc(100vw - 32px))!important;
  min-width:0!important;
  height:auto!important;
  max-height:calc(100dvh - 32px)!important;
}
[data-testid="stDialog"] [role="dialog"]:has(#rfq-confirm-modal) [data-testid="stAlert"] {
  width:100%!important;
  max-width:100%!important;
  box-sizing:border-box!important;
  padding-right:.5rem!important;
}
[data-testid="stDialog"] [role="dialog"]:has(#rfq-confirm-modal) [data-testid="stHorizontalBlock"] {
  justify-content:center!important;
  gap:12px!important;
}
[data-testid="stDialog"] [role="dialog"]:has(#rfq-confirm-modal) [data-testid="column"] {
  flex:0 1 150px!important;
  width:150px!important;
  min-width:0!important;
}
</style>""", unsafe_allow_html=True)

# All formatting is performed locally by deterministic Python rules.
# No API keys, external inference, models, or network calls are required.

def _clean(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (ValueError, TypeError):
        pass
    return str(value).strip()

def normalize_qty(value):
    """Remove thousands-group separators without changing nonnumeric quantities."""
    if isinstance(value, str):
        raw = value.strip()
        if re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?", raw):
            return raw.replace(",", "")
        return value
    return value

def workbook_bytes(workbook):
    """Export workbooks using Aptos Narrow without changing other cell styles."""
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value is None and not cell.has_style:
                    continue
                current_font = copy(cell.font)
                if current_font.name != "Aptos Narrow":
                    current_font.name = "Aptos Narrow"
                    cell.font = current_font
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()

def wrap_all(worksheet):
    for row in worksheet.iter_rows():
        for cell in row:
            align = copy(cell.alignment)
            align.wrap_text = True
            if not align.vertical:
                align.vertical = "top"
            cell.alignment = align

def clear_template_rows(worksheet, start_row=2):
    for row in worksheet.iter_rows(min_row=start_row, max_row=worksheet.max_row):
        for cell in row:
            cell.value = None

def show_error(title, exc):
    st.error(f"{title}: {type(exc).__name__}: {exc}")
    with st.expander("Error details"):
        st.code("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)), language="text")

# Recognize headings even in reversed/extracted text (e.g. some PDFs).
SECTION_PATTERNS = {
    "Basic Data Text": re.compile(r"^basic\s+data(?:\s+text)?\s*:?\s*$", re.I),
    "Classification Text": re.compile(r"^classification\s+text\s*:?\s*$", re.I),
    "MPN Text": re.compile(r"^mpn\s+text\s*:?\s*$", re.I),
    "Inspection Text": re.compile(r"^inspection\s+text\s*:?\s*$", re.I),
    "PO Material Text": re.compile(r"^(?:po\s+material|purchase\s+order)\s+text\s*:?\s*$", re.I),
}
SECTION_ORDER = ("PO Material Text", "Basic Data Text", "Classification Text", "MPN Text", "Inspection Text")
EMPTY_MARKERS = {"", "-", "--", "---", "n/a", "na", "nil", "none", "null"}
META_LINES = set()
FOLLOWED_BY_DATA_LABELS = {"tag nos", "tag numbers", "mpn details"}
PRESERVED_LABELS = {"characteristics", "cell/battery measurement", "battery stand measurement"}
# Narrative subsection titles must not be mistaken for blank attribute fields.
NARRATIVE_HEADING_PATTERN = re.compile(
    r"\b(?:requirements?|instructions?|conditions?|notes?|remarks?|scope|deliverables?|documentation)\b",
    re.I,
)


def is_empty_field(line, next_nonempty=""):
    """Drop only genuinely empty attribute fields; preserve prose and technical values."""
    line = line.strip()
    if not line:
        return True
    if line.casefold() in META_LINES:
        return True
    # Section subheadings can introduce values on subsequent lines.
    # Keep them only when meaningful following content exists.
    normalized_heading = line.rstrip(": \t").strip().casefold()
    if normalized_heading in FOLLOWED_BY_DATA_LABELS and next_nonempty and not match_section(next_nonempty):
        return False
    if ":" in line:
        left, right = line.split(":", 1)
        # Empty field label, including extraction placeholders.
        label = left.strip()
        if label.casefold() in PRESERVED_LABELS or NARRATIVE_HEADING_PATTERN.search(label):
            return False
        if label and right.strip().casefold() in EMPTY_MARKERS:
            return True
    return False


def match_section(line):
    for name, pattern in SECTION_PATTERNS.items():
        if pattern.fullmatch(line.strip()):
            return name
    return None


def extract_manufacturers(text):
    """Only extract explicitly marked manufacturers; never infer from a product name."""
    names = []
    patterns = (
        re.compile(r"^manufacturer\s+number\s*:\s*[^\s]+\s*[-–]\s*(.+)$", re.I),
        re.compile(r"^(?:manufacturer(?:\s+name)?|mfr|mfg|make)\s*:\s*(.+)$", re.I),
    )
    for line in text.splitlines():
        value = line.strip()
        for pattern in patterns:
            match = pattern.match(value)
            if match:
                name = match.group(1).strip()
                if name and name.casefold() not in EMPTY_MARKERS and name.casefold() not in {n.casefold() for n in names}:
                    names.append(name)
                break
    return " - ".join(names)


# Ariba sometimes places the entire PO description, including multiple section
# headings and attributes, into one Excel cell / physical line. Split only
# recognizable labels; preserve all values and the original source for review.
_INLINE_FIELD_LABELS = (
    "NUT MATERIAL SPECIFICATION", "STUD MATERIAL SPECIFICATION",
    "NUT THREAD CLASS", "STUD THREAD CLASS", "NUT DESIGN STANDARD",
    "STUD DESIGN STANDARD", "NO.OF NUT&WASHER", "NO. OF NUT&WASHER",
    "TOTAL LENGTH", "THREAD TYPE", "NUT MATERIAL", "STUD MATERIAL",
    "COATING", "DIAMETER", "CHARACTERISTICS", "CLASS", "TYPE",
    "MATERIAL", "DESIGN STANDARD", "APPLICATION", "STANDARD",
)
_INLINE_LABEL_RE = re.compile(
    r"(?<![\w])(?:" + "|".join(re.escape(t) for t in sorted(_INLINE_FIELD_LABELS, key=len, reverse=True)) + r")\s*:",
    re.I,
)
_INLINE_SECTION_RE = re.compile(
    r"(?:PO Material Text|Purchase Order Text|Material PO Text|Basic Data Text|"
    r"Classification Text|MPN Text|Inspection Text)\s*:", re.I,
)


def expand_inline_rfq_fields(raw):
    """Split concatenated Ariba attributes while retaining original wording.

    Only activate attribute segmentation on lines with multiple known labels,
    or on a line containing an embedded section heading. Single ordinary
    technical lines and narrative prose remain untouched.
    """
    expanded = []
    for physical_line in str(raw).splitlines():
        matches = list(_INLINE_SECTION_RE.finditer(physical_line))
        sections_split = physical_line
        for match in reversed(matches):
            # A heading appearing after narrative text is a new section.
            if match.start() > 0 and sections_split[:match.start()].strip():
                sections_split = sections_split[:match.start()].rstrip() + "\n" + sections_split[match.start():]
            # Split heading from any field text immediately following it.
            end = match.end()
            if end < len(sections_split) and sections_split[end:].strip():
                sections_split = sections_split[:end] + "\n" + sections_split[end:].lstrip()
        for subline in sections_split.splitlines():
            fields = list(_INLINE_LABEL_RE.finditer(subline))
            if len(fields) < 2 and not (fields and fields[0].start() > 0 and matches):
                expanded.append(subline)
                continue
            for field in reversed(fields):
                if field.start() and subline[:field.start()].strip():
                    subline = subline[:field.start()].rstrip() + "\n" + subline[field.start():]
            expanded.extend(subline.splitlines())
    return "\n".join(expanded)


def format_rfq_python(text):
    """Preservation-first deterministic formatter. Never fabricate or paraphrase values.

    The original text is always retained in the review pane, and the user must
    verify every line. Unrecognized content is included instead of discarded.
    """
    raw = _clean(text).replace("\r\n", "\n").replace("\r", "\n")
    # Corrupted Ariba carrier heading (not a technical specification).
    # Only remove the malformed heading itself; retain the following data.
    raw = re.sub(r"(?im)^\s*;\s*Material\s+PO\s+text\s*:\s*(?=\n|$)", "", raw)
    raw = expand_inline_rfq_fields(raw)
    if not raw:
        return "", ""
    # Map supplied lines to known sections without losing free-form descriptions.
    sections = {name: [] for name in SECTION_ORDER}
    active = "Basic Data Text"
    original_lines = [line.strip() for line in raw.split("\n")]
    for idx, line in enumerate(original_lines):
        next_nonempty = next((x for x in original_lines[idx + 1:] if x.strip()), "")
        if not line:
            continue
        header = match_section(line)
        if header:
            active = header
            continue
        if is_empty_field(line, next_nonempty):
            continue
        sections[active].append(line)

    # In SAP extracts PO Material Text often acts as a carrier of Basic Data.
    # Keep a separate PO section only for genuine purchase-order prose preceding
    # a populated basic-data section; otherwise move its text under Basic Data.
    po_lines = sections["PO Material Text"]
    if po_lines:
        if sections["Basic Data Text"]:
            # When both have content, preserve both separately for user review.
            pass
        else:
            sections["Basic Data Text"] = po_lines
            sections["PO Material Text"] = []

    # Correct the known duplicated PDF/Excel text-fragment extraction artifact.
    # Do not remove unrelated text: this is a highly specific repeat pattern.
    for section_lines in sections.values():
        for i, value in enumerate(section_lines):
            section_lines[i] = re.sub(
                r"(?i)(\baccompany this part\.)ompany this part\.",
                r"\1", value)

    # Preserve every MPN occurrence in its original manufacturer/OEM context.
    # The same part number can correspond to different manufacturer numbers,
    # drawings, OEM models, serial numbers, and equipment positions.
    # Deduplicating individual lines destroys these associations.
    # Any genuinely repeated blocks remain visible for human verification.

    # Keep each manufacturer and its corresponding part number together.
    # Insert a blank line only when the previous pair is complete and a
    # new manufacturer or part-number entry starts the next pair.
    mpn_lines = sections["MPN Text"]
    separated_mpn = []
    seen_part = False
    seen_manufacturer = False
    for line in mpn_lines:
        is_part_number = bool(re.match(r"^Manufacturer\s+Part\s+Number\s*:", line, re.I))
        is_manufacturer_number = bool(re.match(r"^Manufacturer\s+number\s*:", line, re.I))
        # The first entry of a subsequent pair is a repeated field type;
        # the complementary field in the first pair should never be split.
        begins_new_pair = (seen_part and seen_manufacturer and
                           ((is_part_number and seen_part) or
                            (is_manufacturer_number and seen_manufacturer)))
        if begins_new_pair:
            if separated_mpn and separated_mpn[-1] != "":
                separated_mpn.append("")
            seen_part = False
            seen_manufacturer = False
        separated_mpn.append(line)
        seen_part = seen_part or is_part_number
        seen_manufacturer = seen_manufacturer or is_manufacturer_number
    # Separate first MPN record from the MPN Details heading, and preserve
    # a blank line between different OEM/manufacturer records.
    for first, value in enumerate(separated_mpn):
        if re.match(r"^Manufacturer\s+(?:Part\s+Number|number)\s*:", value, re.I):
            if first > 0 and separated_mpn[first-1] != "":
                separated_mpn.insert(first, "")
            break
    sections["MPN Text"] = separated_mpn

    chunks = []
    for section in SECTION_ORDER:
        lines = sections[section]
        if lines:
            chunks.append(section + ":\n" + "\n".join(lines))
    # Always preserve relevant data, even when no headers were present.
    output = "\n\n".join(chunks)
    return remove_separator_lines(normalize_pdf_symbols(output)), extract_manufacturers(raw)


def process_rfq_text(text):
    return format_rfq_python(text)


def format_text(text):
    return process_rfq_text(text)[0]


def manufacture_name(text):
    return process_rfq_text(text)[1]

def pdf_text(pdf_file):
    """Extract page text with dependable physical line breaks for RFQ tables."""
    pdf_file.seek(0)
    if fitz is not None:
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        try:
            return "\n".join(page.get_text("text", sort=True) for page in doc)
        finally:
            doc.close()
    pdf_file.seek(0)
    reader = PdfReader(pdf_file)
    output = []
    for i, page in enumerate(reader.pages, start=1):
        try:
            output.append(page.extract_text() or "")
        except Exception as exc:
            raise ValueError(f"Cannot extract page {i}: {exc}") from exc
    return "\n".join(output)

def parse_pdf(body, full_rfq_text):
    """Read ADNOC Gas item blocks without relying on one-line PDF text layout.

    PDF text extractors commonly place each table cell on its own physical line.
    The old all-on-one-line regex therefore missed valid RFQ item rows.
    """
    rfq_match = re.search(r"RFQ\s+Number\s+(\d+)", full_rfq_text, re.I)
    rfq_number = rfq_match.group(1) if rfq_match else "Unknown"
    lines = [line.strip() for line in full_rfq_text.splitlines()]
    results = []
    # ADNOC item header: 00010, 12-digit material number, qty, UOM, price,
    # total, VAT, date; the PDF may wrap each cell onto a different line.
    start_pattern = re.compile(r"^\d{5}$")
    material_pattern = re.compile(r"^(?:B?12\d{10}|B?15\d{10}|B?16\d{10})$")
    date_pattern = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    candidates = []
    for i, value in enumerate(lines):
        if not start_pattern.fullmatch(value):
            continue
        following = lines[i+1:i+12]
        if len(following) < 7 or not material_pattern.fullmatch(following[0]):
            continue
        if not re.fullmatch(r"[\d,]+(?:\.\d+)?", following[1]):
            continue
        if not re.fullmatch(r"[A-Za-z]{1,8}", following[2]):
            continue
        if not any(date_pattern.fullmatch(v) for v in following[3:10]):
            continue
        candidates.append(i)
    for number, idx in enumerate(candidates):
        next_idx = candidates[number+1] if number+1 < len(candidates) else len(lines)
        # Limit to the material item, not the next page's boilerplate.
        segment = lines[idx:next_idx]
        short_pos = next((j for j, v in enumerate(segment) if re.match(r"^Short Text\s*:",v,re.I)), None)
        if short_pos is None:
            continue
        short_text = re.sub(r"^Short Text\s*:\s*", "", segment[short_pos], flags=re.I)
        # Stop at Agreement/LineNo., plant/storage, or footer content.
        stop_pos = next((j for j in range(short_pos+1,len(segment))
                         if re.match(r"^(?:Agreement\s*/\s*LineNo\.|COMMENTS OR SPECIAL INSTRUCTION|REQUEST FOR QUOTATION|INSPECTION/TESTING AND CERTIFICATION REQUIREMENTS)",segment[j],re.I)), len(segment))
        block = segment[short_pos+1:stop_pos]
        po_start = next((j for j,v in enumerate(block) if re.match(r"^PO Material Text\s*:",v,re.I)), None)
        po_lines = block[po_start+1:] if po_start is not None else block
        # Keep section boundaries, blank lines and all substantive content.
        note = "\n".join(po_lines).strip()
        results.append({"RFx Number":rfq_number,
                        "RFx Item No":str(int(lines[idx])),
                        "PR Item No":"", "Material No":lines[idx+1],
                        "Description":short_text, "PO Text":note,
                        "QTY":lines[idx+2], "UOM":lines[idx+3]})
    if results:
        return results
    # Preserve support for unusual legacy exports where item fields are inline.
    pattern = re.compile(r"(\d{5})\s+(\w?12\d{10})\s+(\d+(?:\.\d+)?)\s+(\w+)\s+.*?(\d{2}\.\d{2}\.\d{4})", re.S)
    items = pattern.findall(body)
    shorts = re.findall(r"Short Text :(.*?)\n", body, re.S)
    notes = re.findall(r"PO Material Text :(.*?)Agreement / LineNo\.", body, re.S)
    for i, item in enumerate(items):
        results.append({"RFx Number":rfq_number, "RFx Item No":str(int(item[0])),
                        "PR Item No":"", "Material No":item[1],
                        "Description":shorts[i].strip() if i<len(shorts) else "",
                        "PO Text":notes[i].strip() if i<len(notes) else "",
                        "QTY":item[2], "UOM":item[3]})
    return results

# Nawah quality requirement policy: known codes follow existing established rules;
# unfamiliar codes are carried verbatim into the Working Sheet for review.
_NAWAH_EXCLUDE_CODES = {
    "COMMERCIAL QA", "PC3 MATERIAL", "GENERAL PACKAGING",
    # This is an internal Nawah procurement-at-risk/receipt hold notice,
    # not a supplier-facing technical or certification requirement.
    "RISK",
}
_NAWAH_CERT_CODES = (
    (r"^CERT OF CONFORMANCE$", "CERTIFICATE OF CONFORMANCE"),
    (r"^CERT OF COMPLIANCE$", "CERTIFICATE OF COMPLIANCE"),
    (r"^CERT OF CONFORMITY$", "CERTIFICATE OF CONFORMITY"),
    (r"^CODE CMTR$", "CERTIFIED MATERIAL TEST REPORT (CMTR)"),
    (r"^CMTR$", "MATERIAL TEST REPORT (MTR)"),
    (r"^CODE DATA REPORT$", "CODE DATA REPORT"),
    (r"^CODE RECONCILIATION$", "CODE RECONCILIATION (IF REQUIRED)"),
)
_NAWAH_INLINE_REQUIREMENTS = (
    (r"\bCALIBRATION CERTIFICATE\b|\bCERTIFICATE OF CALIBRATION\b", "CALIBRATION CERTIFICATE"),
    (r"\bMATERIAL TEST CERTIFICATE\b|\bMTC\b", "MATERIAL TEST CERTIFICATE"),
    (r"\bTEST CERTIFICATE\b", "TEST CERTIFICATE"),
)
_NAWAH_START = re.compile(r"^\s*START OF MESSAGE CODE\s+(Z_NPE_[^\n]+?)\s*$", re.I)
_NAWAH_END = re.compile(r"^\s*END OF MESSAGE CODE\s+(Z_NPE_[^\n]+?)\s*$", re.I)


def _summarize_nawah_quality(quality):
    """Handle known message codes; preserve every unknown block unabridged."""
    lines = quality.splitlines()
    blocks = []
    current = None
    loose = []
    for line in lines:
        start = _NAWAH_START.match(line)
        end = _NAWAH_END.match(line)
        if start:
            if current is not None:  # malformed source: do not lose earlier block
                blocks.append((current[0], "\n".join(current[1])))
            elif any(x.strip() for x in loose):
                blocks.append((None, "\n".join(loose)))
            loose = []
            current = (start.group(1).strip().upper(), [line])
        elif current is not None:
            current[1].append(line)
            if end:
                # Even mismatched END codes retain the full source block.
                blocks.append((current[0], "\n".join(current[1])))
                current = None
        else:
            loose.append(line)
    if current is not None:
        blocks.append((current[0], "\n".join(current[1])))
    if any(x.strip() for x in loose):
        blocks.append((None, "\n".join(loose)))

    clauses = []
    unknown = []
    for code, block in blocks:
        payload = block.strip()
        if not payload:
            continue
        # Loose un-coded quality instructions must never be discarded merely
        # because another known requirement was detected elsewhere.
        if code is None:
            payload = re.sub(r"^\s*QUALITY REQUIREMENTS?\s*:?\s*", "", payload, flags=re.I)
            if not payload.strip():
                continue
            # Standalone known certificates can be shortened, but anything
            # with additional instructions remains intact.
            recognized = False
            for pattern, label in _NAWAH_INLINE_REQUIREMENTS:
                if re.fullmatch(pattern, payload.strip(), re.I):
                    clauses.append(label)
                    recognized = True
                    break
            if not recognized:
                unknown.append(payload.strip())
            continue
        short_code = re.sub(r"^Z_NPE_", "", code, flags=re.I).strip().upper()
        if short_code in _NAWAH_EXCLUDE_CODES:
            continue
        if short_code == "CALIBRATION":
            # Unlike a simple calibration-certificate request, Z_NPE_CALIBRATION
            # prescribes calibration methods, traceability, TAR and the report
            # contents. Preserve ALL instructions, matching the approved Nawah
            # Working Sheet; strip only message-code wrappers/dividers.
            substantive = []
            for source_line in block.splitlines():
                if (_NAWAH_START.match(source_line) or _NAWAH_END.match(source_line)
                        or re.fullmatch(r"\s*[*=_-]{4,}\s*", source_line)):
                    continue
                substantive.append(source_line.rstrip())
            calibration = "\n".join(substantive).strip()
            # Preserve section and list separation, but not stacks of blank
            # lines left by the removed message wrappers and divider lines.
            calibration = re.sub(r"\n[ \t]*\n(?:[ \t]*\n)+", "\n\n", calibration)
            if calibration:
                # An individual PDF extraction may leave a terminal full stop
                # outside the message block. It is not a requirement.
                unknown.append(calibration)
            continue
        if short_code == "QA SR MATERIAL":
            clauses.append("NUCLEAR SAFETY-RELATED MATERIAL: NAWAH-approved nuclear QA requirements apply; review full source requirements.")
            continue
        matched = False
        for pattern, label in _NAWAH_CERT_CODES:
            if re.fullmatch(pattern, short_code, re.I):
                clauses.append(label)
                matched = True
                break
        if matched:
            continue
        if re.search(r"Packing Slip or Equivalent Commercial Certification is required", payload, re.I) and short_code in {"PACKING SLIP", "COMMERCIAL CERTIFICATION"}:
            clauses.append("PACKING SLIP OR EQUIVALENT COMMERCIAL CERTIFICATION REQUIRED.")
            continue
        # An unfamiliar code may specify a mandatory certificate, test or
        # inspection. Retain its actual original wording for human verification.
        unknown.append(payload)
    unique_clauses = list(dict.fromkeys(clauses))
    parts = unique_clauses + unknown
    return "\n\n".join(parts)


def summarize_nawah_working_note(raw):
    """Follow existing rules for known QA and preserve any unknown requirements.

    The Uploading Sheet and original review text remain unmodified.
    """
    source = str(raw or "").replace("\r\n", "\n").replace("\r", "\n")
    source = re.sub(r"\*{4,}\s*;?\s*Material PO text\s*:\s*\*{4,}",
                    "Material PO Text:", source, flags=re.I)
    source = re.sub(r"^\s*\[Material PO Text\]\s*$", "Material PO Text:",
                    source, flags=re.I | re.M)
    cut = re.search(
        r"^\s*(?:QUALITY REQUIREMENTS?\s*:?\s*|START OF MESSAGE CODE\s+Z_NPE_[^\n]+)\s*$",
        source, re.I | re.M)
    if cut is None:
        cut = re.search(r"^\s*\*{8,}\s*$\n\s*START OF MESSAGE CODE", source, re.I | re.M)
    material = (source[:cut.start()] if cut else source).strip()
    material = re.sub(r"^\s*Material PO Text\s*:\s*", "", material, flags=re.I)
    quality = source[cut.start():] if cut else ""
    quality_note = _summarize_nawah_quality(quality)
    result = "Material PO Text:\n\n" + material if material else "Material PO Text:"
    if quality_note:
        result += "\n\nQUALITY REQUIREMENTS\n\n" + quality_note
    result = remove_separator_lines(normalize_pdf_symbols(result))
    result = re.sub(r"\n[ \t]*\n(?:[ \t]*\n)+", "\n\n", result)
    return result, extract_manufacturers(material)


def _nawah_excel_records(book, rfq_number):
    """Read Nawah's Ariba .xls Items grid by actual cell content.

    Avoid relying on item-row numbers or a single rigid cell arrangement.
    Header names are used when available; material/PO cells provide an
    independent fallback. Never fabricate quantities or technical text.
    """
    records = []
    diagnostics = []
    material_re = re.compile(r"^\s*(\d{8,12})\s+(.+?)\s*$", re.S)
    marker_re = re.compile(r"^\s*6[.]([0-9]+)(?:[.]0+)?\s*$")
    item_number_re = re.compile(r"^\s*(?:6\.)?([0-9]{1,4})(?:\.0+)?\s*$")
    qty_re = re.compile(r"^\d{1,3}(?:,\d{3})*(?:\.0+)?$|^\d+(?:\.0+)?$")
    units = {"piece": "EA", "pieces": "EA", "each": "EA", "ea": "EA",
             "set": "SET", "sets": "SET", "pc": "PC", "pcs": "PCS",
             "m": "M", "meter": "M", "metre": "M", "kg": "KG"}

    for sheet_name in book.sheet_names:
        table = pd.read_excel(book, sheet_name=sheet_name, header=None,
                              dtype=object, keep_default_na=False)
        rows = [[_clean(v) for v in row] for row in table.itertuples(index=False, name=None)]
        if not rows:
            continue
        headings = {}
        for ridx, row in enumerate(rows):
            for cidx, value in enumerate(row):
                h = re.sub(r"\s+", " ", value).strip().lower().lstrip("* ")
                if h in ("material code", "specifications", "quantity", "unit of measure",
                         "description", "number", "name", "supplier part id"):
                    headings.setdefault(h, cidx)
        for ridx, row in enumerate(rows):
            # The Items sheet usually has column H: '220201319 RELAY,...'
            # and J: '*****;Material PO text:*****\n...'. Search content
            # so the import survives merged headings and column movement.
            material_candidates = [(ci, material_re.fullmatch(value))
                                   for ci, value in enumerate(row) if value]
            material_candidates = [(ci, match) for ci, match in material_candidates if match]
            if not material_candidates:
                continue
            note_candidates = [(ci, value) for ci, value in enumerate(row)
                               if re.search(r"Material\s+PO\s+text\s*:", value, re.I)]
            if not note_candidates:
                continue
            # Only accept material IDs associated with the Specifications area.
            note_col, note = max(note_candidates, key=lambda x: len(x[1]))
            before_note = [c for c in material_candidates if c[0] < note_col]
            mat_col, match = (before_note or material_candidates)[-1]
            material, embedded_description = match.groups()
            # 'ARIBA Number' is the bidding grid's original item number,
            # NOT the index of the extracted rows. Some Nawah workbooks start
            # at item 4 (or later), so len(records)+1 silently corrupts it.
            marker = next((marker_re.fullmatch(v) for v in row[:mat_col]
                           if marker_re.fullmatch(v)), None)
            explicit_item = None
            for heading in ("number", "item number", "item", "lot number"):
                col = headings.get(heading)
                if col is not None and col < mat_col:
                    match_number = item_number_re.fullmatch(row[col])
                    if match_number:
                        explicit_item = int(match_number.group(1))
                        break
            if marker is not None:
                ordinal = int(marker.group(1))
            elif explicit_item is not None:
                ordinal = explicit_item
            else:
                # On older Nawah grids the numbered item cell may be an
                # unlabelled column to the left of Material Code.
                early_cells = row[:min(mat_col, 5)]
                candidate_numbers = [item_number_re.fullmatch(v) for v in early_cells]
                candidate_numbers = [m for m in candidate_numbers if m]
                if not candidate_numbers:
                    raise ValueError(
                        f"Nawah sheet {sheet_name!r}, row {ridx+1}, material {material}: "
                        "source ARIBA item number not identified. "
                        "Please check the original Number column."
                    )
                ordinal = int(candidate_numbers[0].group(1))
            if ordinal < 1:
                continue

            def from_header(name):
                col = headings.get(name)
                return row[col].strip() if col is not None and col < len(row) else ""

            description = (from_header("name") or from_header("description")
                           or embedded_description.strip())
            # Exclude header titles and generic values from description.
            if description.lower() in ("name", "description", "item or lot description"):
                description = embedded_description.strip()
            unit_raw = from_header("unit of measure")
            if not unit_raw:
                unit_raw = next((v for v in row if v.lower() in units), "")
            unit = units.get(unit_raw.lower(), unit_raw.upper())

            qty = from_header("quantity")
            if not qty or not qty_re.fullmatch(qty):
                # The real Nawah Items grid places Quantity two columns after
                # the Specifications cell (J -> L). Try this first, then a
                # narrow range to account for export variations.
                search_cols = [note_col + 2, note_col + 1, note_col + 3, note_col + 4]
                qty = next((row[c].strip() for c in search_cols
                            if 0 <= c < len(row) and qty_re.fullmatch(row[c].strip())), "")
            if not qty:
                raise ValueError(
                    f"Nawah sheet {sheet_name!r}, row {ridx+1}, material {material}: "
                    "item identified, but Quantity cannot be identified. "
                    "Check the Quantity column of the original Items worksheet."
                )
            if not unit:
                raise ValueError(
                    f"Nawah sheet {sheet_name!r}, row {ridx+1}, material {material}: "
                    "item identified, but Unit of Measure is missing."
                )
            note = re.sub(r"^\*+\s*;?\s*Material PO text\s*:\s*\*+",
                          "Material PO Text:", note.strip(), flags=re.I)
            # If a second quality heading is present it must be preserved for
            # Uploading Sheet and reduced only in the Working Sheet formatter.
            records.append({
                "RFx Number": rfq_number,
                "RFx Item No": str((len(records) + 1) * 10),
                "PR Item No": "",
                "Material No": material,
                "Description": description,
                "PO Text": note,
                "QTY": qty,
                "UOM": unit,
                "ARIBA Number": str(ordinal),
            })
        diagnostics.append(f"{sheet_name}: {len(rows)} rows")
    unique = {}
    for item in records:
        unique.setdefault((item["RFx Item No"], item["Material No"]), item)
    return list(unique.values()), diagnostics


def parse_nawah_rfq_excel(book, filename):
    rfq = re.search(r"\d{10}", filename)
    rfq_number = rfq.group(0) if rfq else "Unknown"
    records, diagnostics = _nawah_excel_records(book, rfq_number)
    if records:
        return sorted(records, key=lambda r: int(r["RFx Item No"]))
    raise ValueError(
        "Nawah Excel items not found. The importer searched for a material "
        "number and 'Material PO text:' in the same Items worksheet row. "
        "Worksheets inspected: " + "; ".join(diagnostics)
    )

def parse_nawah_rfq_pdf(pdf_file):
    """Parse NAWAH/ENEC item rows and preserve their per-item QA requirements.

    These RFQs put Material PO Text and extended QUALITY REQUIREMENTS below
    each numbered item, including requirements continuing onto later pages.
    The next numbered item, not a page break, closes the previous item.
    """
    if fitz is None:
        raise RuntimeError("Nawah RFQ extraction requires PyMuPDF (pymupdf).")
    pdf_file.seek(0)
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    try:
        pages = [page.get_text("text", sort=True) for page in doc]
    finally:
        doc.close()
    raw = "\n".join(pages)
    if not (re.search(r"\bNAWAH\s+ENERGY\b|\bENEC\s+Operations\b", raw, re.I)
            and re.search(r"\bMPN\s+NO\b", raw, re.I)
            and re.search(r"\[Material\s+PO\s+Text\]", raw, re.I)):
        return None
    rfq_match = re.search(r"\bRFQ\s+NO\s+(\d{7,12})\b", raw, re.I)
    rfq_number = rfq_match.group(1) if rfq_match else "Unknown"
    # Header includes distinct text cells separated by multiple spaces.
    # Match the item number and material id first; capture the last columns
    # from the right so product model numbers can contain spaces and commas.
    head = re.compile(
        r"^\s*(\d{1,5})\s+(\d{7,12})\s+(.+?)\s{2,}(.+?)\s{2,}"
        r"([\d,]+)\s+([A-Za-z]{1,8})\s+(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})\s*$",
        re.I)
    page_header = re.compile(r"^\s*(?:VENDOR\s+90000000|\d+\s*/\s*\d+\s*$|RFQ\s+NO\s+\d+)" ,re.I)
    entries = []
    current = None
    for page in pages:
        for raw_line in page.splitlines():
            line = raw_line.rstrip()
            found = head.fullmatch(line)
            if found:
                item_no, material, description, mpn, quantity, uom, delivery = found.groups()
                current = {"RFx Number":rfq_number, "RFx Item No":str(int(item_no)),
                           "PR Item No":"", "Material No":material,
                           "Description":description.strip(), "PO Text":"",
                           "QTY":quantity, "UOM":uom, "MPN":mpn.strip(),
                           "Delivery Date":delivery, "_lines":[], "_started":False}
                entries.append(current)
                continue
            if current is None:
                continue
            if re.search(r"^\s*INSTRUCTION\s+TO\s+BIDDERS\s*$", line, re.I):
                current = None
                continue
            if page_header.match(line):
                continue
            if not current["_started"]:
                if re.match(r"^\s*\[Material\s+PO\s+Text\]\s*$", line, re.I):
                    current["_started"] = True
                    current["_lines"].append("Material PO Text:")
                continue
            stripped = line.strip()
            if re.fullmatch(r"[.]+", stripped):
                continue
            current["_lines"].append(stripped)
    records = []
    for entry in entries:
        entry["PO Text"] = "\n".join(entry.pop("_lines")).strip()
        entry.pop("_started", None)
        if not entry["PO Text"]:
            raise ValueError(f"Nawah RFQ item {entry['RFx Item No']} has no Material PO Text")
        records.append(entry)
    if not records:
        raise ValueError("Nawah layout detected but no complete item rows extracted")
    return records


def format_nawah_pdf_note(note):
    """Keep long quality obligations verbatim; do not treat them as empty fields."""
    raw = normalize_pdf_symbols(str(note).replace("\r\n", "\n")).strip()
    # Preserve section breaks, messages, and numbered QA requirements; keep
    # technical attributes on their own lines. Human reviewer controls changes.
    return remove_separator_lines(raw), extract_manufacturers(raw)


def parse_tabular_rfq_pdf(pdf_file):
    """Extract Fertil RFQ table rows, including continuation text across pages.

    The first text line of a numbered row is its description. Remaining physical
    PDF lines are the unmodified PO Text. Pages without any new item number can
    still contribute continuation text to the preceding item.
    """
    if fitz is None:
        raise RuntimeError("PDF table support requires PyMuPDF. Add pymupdf to requirements.txt.")
    pdf_file.seek(0)
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    all_items = []
    layout_seen = False
    rfq_number = ""

    def physical_lines(words):
        rows = []
        for word in sorted(words, key=lambda w: (w[1], w[0])):
            if not rows or abs(word[1] - rows[-1][0]) > 2.5:
                rows.append((word[1], [word]))
            else:
                rows[-1][1].append(word)
        return [" ".join(str(w[4]) for w in sorted(group, key=lambda w: w[0])).strip()
                for _, group in rows]

    def column_words(words, x_low, x_high, top, bottom):
        return [w for w in words if x_low <= w[0] < x_high and top <= w[1] < bottom]

    try:
        for page in doc:
            page_text = page.get_text("text")
            if not rfq_number:
                m = re.search(r"RFQ\s*FORM(?:\s*Page\s+\d+\s*/\s*\d+)?\s*(\d{8,})", page_text, re.I)
                if m:
                    rfq_number = m.group(1)
            # Important: do not confuse ADNOC's unrelated table with Fertil.
            if not (re.search(r"\bRFQ\s+FORM\b", page_text, re.I)
                    and rfq_number):
                continue
            words = page.get_text("words", sort=True)
            headers = [w for w in words if str(w[4]).lower() == "description"
                       and 190 < w[0] < 400 and w[1] < 380]
            header_y = min((w[1] for w in headers), default=None)
            # Fertil's first page has a cover block and table starting ~350;
            # continuation pages have table at y~86, sometimes no header.
            if header_y is not None:
                text_top = header_y + 8
                layout_seen = True
            elif layout_seen:
                text_top = 88
            else:
                continue
            bottom = page.rect.height - 30
            row_heads = sorted([(w[1], int(str(w[4]))) for w in words
                                if text_top <= w[1] < bottom and 0 <= w[0] < 35
                                and re.fullmatch(r"\d{1,3}", str(w[4]))], key=lambda x: x[0])
            # A partial row may occupy the top of a new page BEFORE a new item.
            first_y = row_heads[0][0] if row_heads else bottom
            continuation = physical_lines(column_words(words, 190, 400, text_top, first_y - 1))
            if continuation and all_items:
                previous = all_items[-1]
                previous["PO Text"] = (previous["PO Text"].rstrip() + "\n" + "\n".join(continuation)).strip()

            for idx, (top_y, sequence) in enumerate(row_heads):
                lower_y = row_heads[idx + 1][0] - 1 if idx + 1 < len(row_heads) else bottom
                row_words = [w for w in words if top_y - 2 <= w[1] < lower_y]
                pr_item = " ".join(w[4] for w in column_words(row_words, 85, 119, top_y - 2, lower_y)).strip()
                material = " ".join(w[4] for w in column_words(row_words, 119, 190, top_y - 2, lower_y)).strip()
                qty = " ".join(w[4] for w in column_words(row_words, 400, 445, top_y - 2, lower_y)).strip()
                uom = " ".join(w[4] for w in column_words(row_words, 445, 490, top_y - 2, lower_y)).strip()
                description_lines = physical_lines(column_words(row_words, 190, 400, top_y - 2, lower_y))
                if not description_lines:
                    continue
                if not (re.fullmatch(r"\d{4,6}", pr_item) and re.fullmatch(r"\d{6,}", material)
                        and re.fullmatch(r"[\d,.]+", qty) and re.fullmatch(r"[A-Za-z]{1,8}", uom)):
                    continue
                all_items.append({"RFx Number": rfq_number, "RFx Item No": str(int(pr_item)),
                                  "PR Item No": pr_item, "Material No": material,
                                  "Description": description_lines[0],
                                  "PO Text": "\n".join(description_lines[1:]),
                                  "QTY": qty, "UOM": uom})
    finally:
        doc.close()
    return all_items if all_items else None

def resolve_template(uploaded, default_filename):
    path = uploaded if uploaded is not None else Path(__file__).with_name(default_filename)
    if isinstance(path, Path) and not path.is_file():
        raise FileNotFoundError(f"Missing default template: {path.name}. Please upload a template.")
    return path

def normalize_pdf_symbols(value):
    """Replace the exact unwanted token in generated PDF draft text only."""
    text = str(value)
    return text.replace("<(>&<)>", "&").replace("(>&<)", "&")


def remove_separator_lines(value):
    """Drop visual divider lines made only of separator characters.

    Keep meaningful content such as 'START OF MESSAGE...' while removing
    lines like '----------------', '************', '.............', or
    mixed visual separators containing only punctuation.
    """
    cleaned = []
    for line in str(value).splitlines():
        stripped = line.strip()
        if stripped and re.fullmatch(r"[\-*=._#~]{4,}", stripped):
            continue
        cleaned.append(line)
    return "\n".join(cleaned)


def normalize_technical_punctuation(text):
    """Make label punctuation legible without changing decimal or identifier codes."""
    # Normalize both FIELD : VALUE and FIELD:VALUE. Do not modify URI schemes,
    # clock times (12:30), numeric ratios (1:2), or Windows drive paths (C:\\).
    # Apply to each line, where a labelled technical field is expected.
    def normalize_line(line):
        match = re.match(r"^(\s*)([A-Za-z][A-Za-z0-9 /().#_+\-]{0,75}?)(\s*):(?![/\\])(\s*)(.*)$", line)
        if match:
            leading, label, _, _, value = match.groups()
            label = label.rstrip()
            if label.lower() not in ("http", "https", "ftp"):
                return f"{leading}{label}: {value}" if value else f"{leading}{label}:"
        return line
    text = "\n".join(normalize_line(line) for line in text.split("\n"))
    # Also normalize nested colon-separated technical labels such as
    # 'SET PRESSURE :6.0' while leaving number-number delimiters untouched.
    text = re.sub(r"(?<=[A-Za-z])\s+:\s*(?=[^\s/:])", ": ", text)
    text = re.sub(r"(?<=[A-Za-z]):(?=[^\s/:])", ": ", text)
    # After word-ending commas insert a space unless the next token is numeric
    # punctuation or a code delimiter. Preserve quantities like 1,000.
    text = re.sub(r"(?<=[A-Za-z]),(?=[A-Za-z])", ", ", text)
    text = re.sub(r"(?<=[0-9]),(?=[A-Za-z])", ", ", text)
    return text


def split_nested_equipment_fields(text):
    """Start nested specification fields on new lines after wrapped joins."""
    labels = ("SET PRESSURE", "BACK PRESSURE", "BLOWDOWN", "VALVE ORIFICE",
              "RELVG. CAPACITY", "CAPACITY", "BODY MATERIAL", "TRIM MATERIAL")
    result = []
    for line in text.splitlines():
        if re.match(r"^EQUIPMENT SPECS(?:\d+)?\s*:", line, re.I):
            for label in labels:
                pattern = r",\s*(" + re.escape(label) + r"\s*:)"
                line = re.sub(pattern, lambda m: "\n" + m.group(1), line, flags=re.I)
        result.append(line)
    return "\n".join(result)


def join_pdf_wrapped_lines(note):
    """Reflow wrapped Fertil technical specifications while retaining field boundaries.

    A line ending in an unfinished label/value is joined to its continuation.
    Explicit new uppercase field labels remain on separate lines; unrecognized
    fragments are kept in the previous technical paragraph where safe.
    """
    lines = [normalize_pdf_symbols(line.strip()) for line in str(note).splitlines()]
    heading = re.compile(r"^(?:BASIC DATA TEXT|CLASSIFICATION TEXT|MPN TEXT|INSPECTION TEXT|PO MATERIAL TEXT|MPN DETAILS|CHARACTERISTICS|TAG NOS)\s*:?$", re.I)
    field = re.compile(r"^[A-Z][A-Z0-9 /()_.#-]{1,75}\s*:\s*", re.I)
    narrative = re.compile(r"^(?:ADDITIONAL INFORMATION|PRE-QUOTE|PRE-SHIPMENT|PRE-ORDER CONDITION|STANDARD|SPECIFICATION|REMARKS?|NOTES?|DRAWING NUMBER|HARDNESS)\b", re.I)
    out = []
    for line in lines:
        if not line:
            if out and out[-1] != "": out.append("")
            continue
        line = normalize_technical_punctuation(line)
        if not out or heading.fullmatch(line):
            out.append(line)
            continue
        prev = out[-1]
        if not prev:
            out.append(line)
            continue
        # A named field beginning a fresh specification must start a new line.
        # Except a nested sub-field within an already wrapped EQUIPMENT SPECS.
        if field.match(line) and not (prev.upper().startswith(('EQUIPMENT SPECS', 'ADDITIONAL INFORMATION'))
                                      and re.match(r"^(?:PRESSURE|KG/CM2|VALVE ORIFICE|SET PRESSURE|BACK PRESSURE|CAPACITY|MATERIAL|TRIM MATERIAL|INLET|OUTLET)\b",line,re.I)):
            out.append(line)
            continue
        if prev.endswith((':', ',', ';', '-', '/', '&')) or re.search(r"\b(?:SET|VALVE|BODY|TRIM|SPIR|REL[VW]G\.?|INLET|OUTLET|MATERIAL|PRESSURE)\s*$",prev,re.I):
            out[-1] = prev.rstrip() + " " + line
            continue
        if narrative.match(prev) or prev.upper().startswith(('EQUIPMENT SPECS', 'ADDITIONAL INFORMATION')):
            out[-1] = prev.rstrip() + " " + line
            continue
        # Wrapped part numbers after '&' and drawing/tag values.
        if re.match(r"^[A-Z0-9][A-Z0-9./#-]*(?:\s+[A-Z0-9./#-]+)*$",line) and (
                re.search(r"(?:PART NUMBER|DRAWING NUMBER|EQUIPMENT TAG NUMBER)\s*:",prev,re.I)):
            out[-1] = prev.rstrip() + " " + line
            continue
        out.append(line)
    return "\n".join(out).strip()

def format_tabular_pdf_note(note):
    """Readable PDF review draft; keep technical fields separate, join prose wraps."""
    physical = [line.strip() for line in str(note).splitlines() if line.strip()]
    if not physical:
        return "", ""

    formatted = split_nested_equipment_fields(join_pdf_wrapped_lines(note))
    # A tabular item's first sentence is exported as Description, not as part
    # of the formatted PO Text. Strip it even if it leaked into the note.
    # (The caller passes a separate description when preparing review.)
    # Group some related fields visually, without changing PDF upload text.
    output = []
    group_start = re.compile(
        r"^(?:MAKE|MODEL NUMBER OF EQUIPMENT|ADDITIONAL INFORMATION)\s*:", re.I
    )
    for line in formatted.splitlines():
        if output and group_start.match(line) and output[-1] != "":
            output.append("")
        output.append(line)

    # Continue extracting manufacturer from the original source text, avoiding
    # inadvertent changes to manufacturer names from paragraph normalization.
    manufacturer = ""
    extra_index = next((i for i, line in enumerate(physical)
                        if re.match(r"^ADDITIONAL\s+INFORMATION\s*:", line, re.I)), None)
    if extra_index is not None:
        extra_lines = []
        for line in physical[extra_index:]:
            if extra_lines and re.match(
                r"^(?:POSITION NUMBER|ITEM TYPE|DRAWING NUMBER|MANUFACTURER PART NUMBER)\s*:",
                line, re.I
            ):
                break
            extra_lines.append(line)
        extra = " ".join(extra_lines)
        match = re.search(r"\bMAKE\s*:\s*(.*?)(?=,\s*(?:MODEL|SERIAL|TAG)\s*:?|$)", extra, re.I)
        if match:
            manufacturer = match.group(1).strip().strip(", ")
    if not manufacturer:
        manufacturer = extract_manufacturers("\n".join(physical))
    return remove_separator_lines("\n".join(output)), manufacturer


def parse_dubai_petroleum_csv(uploaded):
    """Parse the Dubai Petroleum bidding export, including empty leading columns."""
    uploaded.seek(0)
    content = uploaded.read()
    if isinstance(content, bytes):
        try:
            content = content.decode("utf-8-sig")
        except UnicodeDecodeError:
            # Some Dubai Petroleum exports use Windows-1252 (e.g. en dashes).
            content = content.decode("cp1252")
    rows = list(csv.reader(content.splitlines(keepends=True)))
    if not rows:
        raise ValueError("Dubai Petroleum CSV has no rows")
    header = [re.sub(r"\s+", " ", col.strip()).casefold() for col in rows[0]]
    required = ("line", "item, rev", "item description", "long description", "unit", "quantity")
    missing = [name for name in required if name not in header]
    if missing:
        raise ValueError("Not a supported Dubai Petroleum CSV export. Missing headings: " + ", ".join(missing))
    indices = {name: header.index(name) for name in required}
    parsed = []
    for index, row in enumerate(rows[1:], start=2):
        def get(name):
            col = indices[name]
            return row[col].strip() if len(row) > col else ""
        item_no = get("line")
        if not item_no:
            continue
        if not re.fullmatch(r"\d+", item_no):
            raise ValueError(f"Invalid line number on CSV row {index}: {item_no!r}")
        description = get("item description")
        note = get("long description").replace("\r\n", "\n").replace("\r", "\n")
        quantity = get("quantity")
        unit = get("unit")
        if not all((description, note, quantity, unit)):
            raise ValueError(f"CSV row {index} missing description, long description, quantity or unit")
        parsed.append({"RFx Item No": str(int(item_no) * 10), "Material No": get("item, rev"),
                       "Description": description, "PO Text": note, "QTY": quantity, "UOM": unit})
    if not parsed:
        raise ValueError("Dubai Petroleum CSV contains no item rows")
    return parsed


def format_dubai_petroleum_note(source):
    """Add readable breaks between logical groups, not inside manufacturer sets.

    Preserve the wording, sequence, and source lines.  Either PART NO -> MFR
    or MFR -> PART NO is one contiguous pair.
    """
    text = remove_separator_lines(normalize_pdf_symbols(str(source)))
    lines = [line.strip() for line in text.splitlines()]
    output = []
    manufacturer_re = re.compile(r"^(?:MFR\.?|MANUFACTURER(?:\s+NAME)?)\s*[:.,]", re.I)
    part_re = re.compile(r"^(?:PART\s*(?:NO\.?|NUMBER)|MFR\.?\s*PART\s*(?:NO\.?|NUMBER)|MANUFACTURER\s+PART\s+NUMBER)\s*:", re.I)
    note_re = re.compile(r"^(?:USED FOR\b|NOTE\s*:|\*+\s*NOTE\b)", re.I)

    for idx, stripped in enumerate(lines):
        if not stripped:
            if output and output[-1] != "":
                output.append("")
            continue
        previous = next((entry for entry in reversed(output) if entry), "")
        is_mfr = bool(manufacturer_re.match(stripped))
        is_part = bool(part_re.match(stripped))
        prev_mfr = bool(manufacturer_re.match(previous))
        prev_part = bool(part_re.match(previous))
        is_note = bool(note_re.match(stripped))
        prev_note = bool(note_re.match(previous))

        # Start a manufacturer/part-number block, but keep the counterpart
        # on the very next line.  Multiple different pairs remain distinct.
        start_identification = (is_mfr or is_part) and not (prev_mfr or prev_part)
        # Notes are separated from specification and identification blocks.
        start_note = is_note and not prev_note
        if output and output[-1] != "" and (start_identification or start_note):
            output.append("")
        output.append(stripped)
    # Sample Working Sheet leaves Manufacturer blank for multi-brand alternatives.
    return "\n".join(output).strip(), ""


# RFQ checks are advisory; the original source always remains authoritative.
VALIDATION_LABELS = re.compile(
    r"^(?:MFR|MFG|MANUFACTURER(?: NAME| NUMBER| PART NUMBER)?|PART(?: NUMBER| NO)?|"
    r"MPN|OEM(?: PART| MODEL| SERIAL)?(?: NUMBER)?|MODEL(?: NUMBER)?|"
    r"DRAWING(?: NUMBER| NO)?|TAG(?: NUMBER| NO| NOS)?|"
    r"MATERIAL|GRADE|STANDARD|SIZE(?: [A-Z])?|THICKNESS|DIAMETER|LENGTH|"
    r"PRESSURE(?: RATING| RANGE)?|RANGE|TEMPERATURE(?: RATING)?|VOLTAGE|"
    r"POWER(?: SUPPLY| RATING)?|CONNECTION|CLASS(?: RATING)?|"
    r"FLOW(?: RATING)?|CAPACITY|ACCURACY|RATED VOLTAGE)\s*:\s*(.+)$", re.I)
VALIDATION_EMPTY = {"", "-", "--", "---", "N/A", "NONE", "NULL"}
CRITICAL_FIELD = re.compile(r"MFR|MANUFACTURER|PART|MPN|OEM|MODEL|DRAWING|TAG", re.I)
NUMBER_TOKEN = re.compile(r"(?<![A-Za-z])(?:\d+(?:[.,]\d+)*(?:/\d+)?)(?:\s*(?:MM|CM|IN|INCH|BAR|PSI|VAC|VDC|V|HZ|RPM|DEG|°C|°F|KG|LPM))?", re.I)

def compare_normalize(value):
    value = normalize_pdf_symbols(str(value)).casefold()
    return re.sub(r"[^a-z0-9]+", "", value)

def _multiset_contains(original, formatted):
    return compare_normalize(original) in compare_normalize(formatted)

def validate_rfq(source, draft, format_style="default", description=""):
    """Check labeled technical fields, numbers and certificate deliverables.

    Do not flag omitted source sections for Nawah that are intentionally
    summarized. The warnings are review aids, never claims of full compliance.
    """
    source = str(source or "")
    draft = str(draft or "")
    if format_style == "tabular_pdf" and description:
        draft = description + "\n" + draft
    src = source.splitlines()
    cut = next((i for i, line in enumerate(src) if re.search(r"^\s*(?:QUALITY REQUIREMENTS?\s*:?|START OF MESSAGE CODE)", line, re.I)), len(src)) if format_style.startswith("nawah") else len(src)
    check_lines = src[:cut]
    warnings = []
    seen = set()
    for line in check_lines:
        match = VALIDATION_LABELS.match(line.strip())
        if not match:
            continue
        label, value = match.group(0).split(":", 1)
        value = value.strip()
        if value.upper() in VALIDATION_EMPTY:
            continue
        # Check the technical value, allowing label changes and spaces/punctuation.
        if not _multiset_contains(value, draft):
            priority = "Critical" if CRITICAL_FIELD.search(label) else "Review"
            signature = (label.upper(), value)
            if signature not in seen:
                warnings.append({"priority":priority,"field":label.strip(),"value":value,
                                 "message":"Technical value not found in formatted text"})
                seen.add(signature)
    # For complex narrative, compare digit-bearing identifiers outside recognized
    # fields. Ignore boilerplate separators and skip this check in Nawah source QA.
    number_source = "\n".join(line for line in check_lines if not VALIDATION_LABELS.match(line.strip()))
    for token in NUMBER_TOKEN.findall(number_source):
        if len(compare_normalize(token)) < 3:
            continue
        if not _multiset_contains(token, draft) and len(warnings) < 35:
            key = ("numeric",compare_normalize(token))
            if key not in seen:
                warnings.append({"priority":"Review", "field":"Number / unit", "value":token.strip(),
                                 "message":"Number or unit may be missing or changed"})
                seen.add(key)
    if format_style.startswith("nawah"):
        # Explicit certificate requirements must survive the shortened QA section.
        requirements = [
            (r"Z_NPE_CERT OF CONFORMANCE", "CERTIFICATE OF CONFORMANCE"),
            (r"Z_NPE_CERT OF COMPLIANCE", "CERTIFICATE OF COMPLIANCE"),
            (r"Z_NPE_CERT OF CONFORMITY", "CERTIFICATE OF CONFORMITY"),
            (r"Z_NPE_CODE CMTR", "CERTIFIED MATERIAL TEST REPORT"),
            (r"Z_NPE_CMTR", "MATERIAL TEST REPORT"),
            (r"Z_NPE_CODE DATA REPORT", "CODE DATA REPORT"),
            (r"Z_NPE_CODE RECONCILIATION", "CODE RECONCILIATION"),
        ]
        for pattern, name in requirements:
            if re.search(pattern, source, re.I) and name.casefold() not in draft.casefold():
                warnings.append({"priority":"Critical", "field":"Quality requirement", "value":name,
                                 "message":"Required documentation may be missing"})
    return warnings


def _classification(line, draft, format_style):
    val = line.strip()
    if not val:
        return "normal"
    if re.fullmatch(r"[-*=_#. ~]{4,}", val):
        return "excluded"
    if format_style.startswith("nawah") and re.search(r"COMMERCIAL QA|AUGMENTED QUALITY|GENERAL PACKAGING|START OF MESSAGE CODE|END OF MESSAGE CODE", val, re.I):
        return "excluded"
    # Empty placeholders are intentionally removed, but --- markers should
    # be RED in the source so reviewers can verify their removal explicitly.
    if re.match(r"^[\w /().-]+\s*:\s*-{3,}\s*$", val, re.I):
        return "missing"
    if re.match(r"^[\w /().-]+\s*:\s*(?:-+|N/?A)?$", val, re.I) and val.casefold() not in {"characteristics :", "mpn details"}:
        return "excluded"
    if _multiset_contains(val, plain_review_text(draft)):
        return "matched"
    return "missing"


def comparison_html(original, formatted, style="default", sync=True, height=510):
    """Both previews live in one iframe, allowing reliable scroll linking."""
    original_lines = original.splitlines()
    rendered = []
    stats = Counter()
    # A source line occurring twice must also occur twice in the draft.
    # Otherwise, highlight the missing occurrence even when the first matches.
    available_counts = Counter(compare_normalize(line) for line in plain_review_text(formatted).splitlines() if line.strip())
    in_quality = False
    for line in original_lines:
        if style.startswith("nawah") and re.match(r"^\s*QUALITY REQUIREMENTS?\s*:?\s*$", line, re.I):
            in_quality = True
        cls = _classification(line, formatted, style)
        key = compare_normalize(line)
        if cls == "matched" and key:
            if available_counts[key] > 0:
                available_counts[key] -= 1
            else:
                cls = "missing"
        if in_quality and style.startswith("nawah") and cls == "missing":
            cls = "excluded"  # Nawah quality boilerplate is summarized intentionally.
        stats[cls] += 1
        rendered.append('<span class="' + cls + '">' + html.escape(line) + '</span>')
    left = "\n".join(rendered)
    right_lines = []
    for line in formatted.splitlines():
        display = html.escape(line)
        parts = []
        offset = 0
        for m in BOLD_MARKERS.finditer(line):
            parts.append(html.escape(line[offset:m.start()]))
            parts.append('<strong>' + html.escape(m.group(1)) + '</strong>')
            offset = m.end()
        parts.append(html.escape(line[offset:]))
        display = ''.join(parts)
        if BOLD_TECHNICAL_LINE.match(plain_review_text(line)):
            display = '<strong>' + display + '</strong>'
        right_lines.append(display)
    right = "\n".join(right_lines)
    document = """<!doctype html><html><head><meta charset="utf-8"><style>
    html,body{margin:0;background:transparent;color:#172033;font:13px ui-monospace,Consolas,monospace}
    *{box-sizing:border-box}
    .grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:12px;padding:2px;width:100%;max-width:100%;overflow:hidden}
    .grid>div{min-width:0;max-width:100%;overflow:hidden}
    .heading{font:600 13px system-ui;margin:0 0 9px;color:#dae2f0}
    .panel{height:HEIGHTpx;overflow:auto;scrollbar-gutter:stable;background:#ffffff;
      border:1px solid #aebccd;border-radius:12px;padding:14px;box-sizing:border-box;min-width:0;width:100%}
    pre{white-space:pre-wrap;overflow-wrap:anywhere;margin:0;line-height:1.65;font:12.5px ui-monospace,Consolas,monospace}
    .missing{color:#bd1e32;font-weight:600}.excluded{color:#0868a8}
    .matched{color:#087735}.normal{color:#172033}
    .legend{font:12px system-ui;color:#a6b6cb;margin:8px 0 0}
    @media(max-width:620px){.grid{grid-template-columns:1fr}.panel{height:240px}}
    </style></head><body><div class="grid">
    <div><div class="heading">Original RFQ · selectable</div><div id="src" class="panel"><pre>LEFT</pre></div></div>
    <div><div class="heading">Formatted output · preview</div><div id="dst" class="panel"><pre>RIGHT</pre></div></div>
    </div><div class="legend">Red: review for omission · Light blue: intentionally excluded / empty · Bright green: retained · MODE</div>
    <script>const a=document.getElementById('src'),b=document.getElementById('dst');
    let lock=false;const synced=SYNC;
    function follow(from,to){if(!synced||lock)return;lock=true;
       let max=Math.max(1,from.scrollHeight-from.clientHeight), target=Math.max(0,to.scrollHeight-to.clientHeight);
       to.scrollTop=(from.scrollTop/max)*target;to.scrollLeft=from.scrollLeft;lock=false;}
    a.addEventListener('scroll',()=>follow(a,b));b.addEventListener('scroll',()=>follow(b,a));
    </script></body></html>"""
    # Replace template placeholders before inserting RFQ text. Global replacement
    # afterwards corrupts source words such as MODEL (MODE -> scroll label).
    document = document.replace('height:HEIGHTpx', f'height:{height}px')
    document = document.replace('const synced=SYNC;',
                                'const synced=' + ('true' if sync else 'false') + ';')
    document = document.replace('· MODE</div>',
                                '· ' + ('Synchronized scroll' if sync else 'Independent scroll') + '</div>')
    document = document.replace('<pre>LEFT</pre>', '<pre>' + left + '</pre>')
    document = document.replace('<pre>RIGHT</pre>', '<pre>' + right + '</pre>')
    components.html(document, height=height + 62, scrolling=False, width=None)
    return stats


def repair_split_bold_markers(text):
    """Rejoin an orphan closing ** with the preceding bolded line.

    Ctrl+B selections can include a trailing newline. This normalizes older
    drafts without joining independent technical specification lines.
    """
    text = str(text or "")
    return re.sub(
        r"(?m)^(\*\*[^\n]*?\S)[ \t]*\n[ \t]*\*\*[ \t]*(?=\n|$)",
        r"\1**", text,
    )


def editor_bold_markers(text):
    """Expose automatically bold technical lines using editable **...** markers.

    Marker delimiters are inserted inside each line, never across a newline.
    Already marked lines remain untouched to preserve manually toggled styling.
    """
    output = []
    for line in repair_split_bold_markers(text).split("\n"):
        if BOLD_TECHNICAL_LINE.match(plain_review_text(line)) and "**" not in line:
            output.append("**" + line + "**")
        else:
            output.append(line)
    return "\n".join(output)


def _review_enter_edit(prefix):
    """Initialize each editor from the last nonempty approved/saved draft."""
    review = st.session_state[f"{prefix}_review"]
    pos = review["position"]
    item = review["queue"][pos]
    current = st.session_state.get(f"formatted_{prefix}_{pos}") or item["formatted"]
    if not str(current).strip():
        current = item.get("initial_formatted", "")
    st.session_state[f"editable_{prefix}_{pos}"] = editor_bold_markers(current)
    st.session_state[f"editablemaker_{prefix}_{pos}"] = st.session_state.get(
        f"maker_{prefix}_{pos}", item["manufacturer"])
    item["editing"] = True


def _save_current(review, prefix):
    """Save edits without ever replacing a nonempty item by an empty widget state."""
    pos = review["position"]
    item = review["queue"][pos]
    if item.get("editing"):
        value = st.session_state.get(f"editable_{prefix}_{pos}")
        if value is None or (not str(value).strip() and str(item["source"]).strip()):
            st.error("The formatted editor is empty. Restore or enter text before leaving this item.")
            return False
        item["formatted"] = repair_split_bold_markers(value)
        item["manufacturer"] = st.session_state.get(f"editablemaker_{prefix}_{pos}", item["manufacturer"])
    else:
        value = st.session_state.get(f"formatted_{prefix}_{pos}", item["formatted"])
        if value is None or (not str(value).strip() and str(item["source"]).strip()):
            st.error("Formatted text is empty; cannot navigate without restoring it.")
            return False
        item["formatted"] = value
        item["manufacturer"] = st.session_state.get(f"maker_{prefix}_{pos}", item["manufacturer"])
    if item.get("tabular_pdf"):
        item["desc"] = st.session_state.get(f"desc_{prefix}_{pos}", item["desc"])
    st.session_state[f"formatted_{prefix}_{pos}"] = item["formatted"]
    st.session_state[f"maker_{prefix}_{pos}"] = item["manufacturer"]
    return True

def _navigate_review(review, prefix, new_position):
    if not _save_current(review, prefix):
        return
    review["position"] = new_position
    st.rerun(scope="fragment")


def space_after_initial_section_header(text):
    """Use one blank line after an initial PO/Basic/Classification/MPN/Inspection header.

    Only the generated Working Sheet draft is changed. Original RFQ and
    Uploading Sheet text remain untouched.
    """
    text = str(text or "")
    return re.sub(
        r"\A(\s*(?:(?:Basic Data|Classification|MPN|Inspection|Material PO|PO Material|Purchase Order)\s+Text)\s*:)\s*\n+",
        lambda m: m.group(1) + "\n\n",
        text,
        count=1,
        flags=re.I,
    )


def prepare_review(records, template, suffix, session_prefix, upload_wb=None, format_style="default", no_format=False):
    """Pre-format all rows instantly, queue for mandatory per-row verification."""
    if not records:
        raise ValueError("No RFQ items to process")
    # Validate the template before initializing a review session.
    template_wb = load_workbook(template)
    template_data = workbook_bytes(template_wb)
    queue = []
    for line, desc, qty, uom, note in records:
        formatted, manufacturer = ((str(note) if note is not None else ""), "") if no_format else (format_dubai_petroleum_note(note)
                                   if format_style == "dubai_csv"
                                   else format_tabular_pdf_note(note)
                                   if format_style == "tabular_pdf"
                                   else summarize_nawah_working_note(note) if format_style in ("nawah_pdf", "nawah_excel")
                                   else process_rfq_text(note))
        # Formatting-only normalization across all RFQ formats. The original
        # source and Uploading Sheet PO Text are deliberately not altered.
        if not no_format:
            formatted = normalize_technical_punctuation(formatted)
            formatted = space_after_initial_section_header(formatted)
        is_table = format_style == "tabular_pdf"
        if is_table and not no_format:
            desc_clean = str(desc).strip()
            lines = formatted.splitlines()
            if lines and lines[0].strip().casefold() == desc_clean.casefold():
                formatted = "\n".join(lines[1:]).lstrip("\n")
        # On this PDF layout the first line is the description. Show that full
        # original item in the left reference, but keep PO Text separate for export.
        source = (str(desc).strip() + "\n" + _clean(note)).strip() if is_table else _clean(note)
        queue.append({"line": line, "desc": desc, "qty": qty, "uom": uom,
                      "source": source, "formatted": formatted,
                      "tabular_pdf": is_table, "manufacturer": manufacturer,
                      "verified": False, "editing": False, "style":format_style,
                      "initial_formatted": formatted, "initial_manufacturer": manufacturer,
                      "initial_desc": desc, "history": []})
    for old_key in list(st.session_state):
        if old_key.startswith((f"original_{session_prefix}_", f"formatted_{session_prefix}_", f"maker_{session_prefix}_", f"editable_{session_prefix}_", f"editablemaker_{session_prefix}_", f"desc_{session_prefix}_")):
            del st.session_state[old_key]
    st.session_state[f"{session_prefix}_review"] = {
        "queue": queue, "position": 0, "suffix": suffix, "format_label": {
            "dubai_csv":"Dubai Petroleum CSV", "tabular_pdf":"Fertil table PDF",
            "nawah_pdf":"Nawah PDF", "nawah_excel":"Nawah Excel",
            "default":"Standard Ariba / ADNOC"}.get(format_style,format_style),
        "template": template_data,
        "upload_bytes": workbook_bytes(upload_wb) if upload_wb is not None else None,
        "no_format": no_format,
    }
    # Uploading Sheet has no AI formatting and can be downloaded immediately.
    st.session_state.pop(f"{session_prefix}_final_bytes", None)
    st.session_state[f"{session_prefix}_suffix"] = suffix
    if upload_wb is not None:
        st.session_state[f"{session_prefix}_upload_bytes"] = workbook_bytes(upload_wb)
    else:
        st.session_state.pop(f"{session_prefix}_upload_bytes", None)


# Bold emphasis in Working Sheet only. **markers** entered via Ctrl+B are
# translated to actual Excel rich text; no literal Markdown is exported.
BOLD_TECHNICAL_LINE = re.compile(
    r"^\s*(?:manufacturer(?:\s+part)?\s+(?:number|name)|"
    r"manufacturer\s*[:：]|mfr\.?\s*[:：]|"
    r"mfg\s*[:：]|make\s*[:：]|"
    r"part\s*(?:no\.?|number)\s*[:：]|"
    r"item\s*[-–]?\s*manufacturer\s*part\s*(?:no\.?|number)\s*[:：]|"
    r"(?:mfr|oem|smg|supplier)\s+part\s*(?:no\.?|number)\s*[:：])",
    re.I,
)
BOLD_MARKERS = re.compile(r"\*\*(.+?)\*\*", re.S)


def plain_review_text(text):
    """Strip explicit bold markers only; leave source values untouched."""
    return repair_split_bold_markers(text).replace("**", "")


def working_sheet_rich_text(text):
    """Turn manual emphasis and manufacturer/part-number lines into Excel runs."""
    source = repair_split_bold_markers(text)
    segments = []
    for line in source.splitlines(keepends=True):
        auto_bold = bool(BOLD_TECHNICAL_LINE.match(line))
        cursor = 0
        for match in BOLD_MARKERS.finditer(line):
            if match.start() > cursor:
                segments.append((line[cursor:match.start()], auto_bold))
            segments.append((match.group(1), True))
            cursor = match.end()
        if cursor < len(line):
            segments.append((line[cursor:], auto_bold))
    if not segments:
        return source
    # Preserve the template's base font; only set bold on the rich-text runs.
    result = CellRichText()
    for value, is_bold in segments:
        if not value:
            continue
        if is_bold:
            result.append(TextBlock(InlineFont(b=True, rFont="Aptos Narrow", sz=11), value))
        else:
            result.append(value)
    return result


def make_rich_text_shared_strings(data, items=None):
    """Insert Excel-standard shared rich strings generated by XlsxWriter.

    Do not hand-author `<rPr>`: Excel has repeatedly repaired the old handmade
    rich-string XML even when openpyxl could load it. Keep the original
    openpyxl workbook, styles, fills, and column widths intact, and transplant
    only XlsxWriter-created shared string entries for Working Sheet column E.
    """
    if not items:
        return data
    import zipfile
    import xlsxwriter
    from xml.etree import ElementTree as etree

    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    content_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    relation_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    q = lambda tag: f"{{{ns}}}{tag}"

    # XlsxWriter generates the rich string properties and their XML ordering.
    scratch = BytesIO()
    scratch_book = xlsxwriter.Workbook(scratch, {"in_memory": True})
    temp_sheet = scratch_book.add_worksheet("Rich Text")
    base_font = scratch_book.add_format({"font_name": "Aptos Narrow", "font_size": 11})
    bold_font = scratch_book.add_format({"font_name": "Aptos Narrow", "font_size": 11, "bold": True})
    indices = {}
    for idx, item in enumerate(items):
        plain = plain_review_text(item["formatted"])
        runs = working_sheet_runs(item["formatted"])
        # XlsxWriter requires at least two fragments for write_rich_string.
        if len(runs) >= 2:
            parts = []
            for value, bold in runs:
                parts.extend((bold_font if bold else base_font, value))
            temp_sheet.write_rich_string(idx, 0, *parts)
        elif len(runs) == 1:
            value, bold = runs[0]
            temp_sheet.write_string(idx, 0, value, bold_font if bold else base_font)
        else:
            temp_sheet.write_string(idx, 0, plain, base_font)
    scratch_book.close()
    with zipfile.ZipFile(BytesIO(scratch.getvalue())) as temp:
        generated_sst = etree.fromstring(temp.read("xl/sharedStrings.xml"))
        temp_sheet_xml = etree.fromstring(temp.read("xl/worksheets/sheet1.xml"))
        for cell in temp_sheet_xml.iter(q("c")):
            address = cell.get("r", "")
            match = re.fullmatch(r"A(\d+)", address)
            v = cell.find(q("v"))
            if match and v is not None and cell.get("t") == "s":
                indices[int(match.group(1)) + 1] = v.text
    with zipfile.ZipFile(BytesIO(data)) as archive:
        source_files = {name: archive.read(name) for name in archive.namelist()}
    sheet_name = "xl/worksheets/sheet1.xml"
    sheet_root = etree.fromstring(source_files[sheet_name])
    updated = 0
    for cell in sheet_root.iter(q("c")):
        match = re.fullmatch(r"E(\d+)", cell.get("r", ""))
        if not match:
            continue
        value = indices.get(int(match.group(1)))
        if value is None:
            continue
        for child in list(cell):
            if child.tag in {q("is"), q("v"), q("f")}:
                cell.remove(child)
        cell.set("t", "s")
        etree.SubElement(cell, q("v")).text = value
        updated += 1
    if updated != len(items):
        raise ValueError(f"Rich text export covered {updated}/{len(items)} Working Sheet items")
    source_files[sheet_name] = etree.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
    source_files["xl/sharedStrings.xml"] = etree.tostring(generated_sst, encoding="utf-8", xml_declaration=True)
    ct = etree.fromstring(source_files["[Content_Types].xml"])
    if not any(node.get("PartName") == "/xl/sharedStrings.xml" for node in ct):
        etree.SubElement(ct, f"{{{content_ns}}}Override", PartName="/xl/sharedStrings.xml",
                         ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
    source_files["[Content_Types].xml"] = etree.tostring(ct, encoding="utf-8", xml_declaration=True)
    relationship_file = "xl/_rels/workbook.xml.rels"
    rel_root = etree.fromstring(source_files[relationship_file])
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    if not any(el.get("Type") == rel_type for el in rel_root):
        used = [int(el.get("Id")[3:]) for el in rel_root if re.fullmatch(r"rId\d+", el.get("Id", ""))]
        etree.SubElement(rel_root, f"{{{relation_ns}}}Relationship",
                         Id=f"rId{max(used, default=0)+1}", Type=rel_type,
                         Target="sharedStrings.xml")
    source_files[relationship_file] = etree.tostring(rel_root, encoding="utf-8", xml_declaration=True)
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, value in source_files.items():
            archive.writestr(name, value)
    return out.getvalue()


def working_sheet_runs(value):
    """Merge adjacent rich text runs so an entire MPN line stays bold."""
    rich = working_sheet_rich_text(value)
    if not isinstance(rich, CellRichText):
        return [(str(rich), False)] if rich else []
    pieces = []
    for segment in rich:
        if isinstance(segment, TextBlock):
            part, bold = segment.text, bool(segment.font.b)
        else:
            part, bold = str(segment), False
        if not part:
            continue
        if pieces and pieces[-1][1] == bold:
            pieces[-1] = (pieces[-1][0] + part, bold)
        else:
            pieces.append((part, bold))
    return pieces


def finalize_verified(session_prefix):
    review = st.session_state[f"{session_prefix}_review"]
    queue = review["queue"]
    if not queue or not all(item["verified"] for item in queue):
        raise ValueError("All RFQ items must be verified before exporting")
    wb = load_workbook(BytesIO(review["template"]))
    sheet = wb.active
    clear_template_rows(sheet)
    # Some replacement Working Sheet templates style only row 1. Carry its
    # line-item blue fill into every populated data row (without recoloring
    # other columns or overwriting fonts, borders or alignment).
    from openpyxl.styles import PatternFill
    line_fill = copy(sheet["A1"].fill)
    if line_fill.patternType != "solid":
        line_fill = PatternFill(fill_type="solid", fgColor="B4C6E7")
    for i, item in enumerate(queue, 2):
        for col, value in {
            "A": item["line"], "B": item["desc"], "C": item["qty"],
            "D": item["uom"], "E": (item["formatted"] if review.get("no_format") else plain_review_text(item["formatted"])),
        }.items():
            sheet[f"{col}{i}"] = value
        sheet[f"A{i}"].fill = copy(line_fill)
    # Keep the blue Line item number column continuous below the last item.
    # Replacement templates sometimes have unstyled blank rows between the
    # populated area and their next preformatted row; fill those gaps too.
    blue_fill_end_row = max(sheet.max_row, len(queue) + 26)
    for row_number in range(2, blue_fill_end_row + 1):
        sheet[f"A{row_number}"].fill = copy(line_fill)
    sheet["E1"] = "Internal Note"
    # The Working Sheet contains only A:E. Remove any legacy Manufacturer,
    # Subject line, mail ID or unused template columns rather than exporting
    # a sixth field.
    if sheet.max_column > 5:
        sheet.delete_cols(6, sheet.max_column - 5)
    for other in list(wb.worksheets):
        if other is not sheet:
            wb.remove(other)
    wrap_all(sheet)
    output_bytes = (workbook_bytes(wb) if review.get("no_format")
                    else make_rich_text_shared_strings(workbook_bytes(wb), queue))
    # Reopen the exact generated file to catch malformed OOXML before delivery.
    try:
        reopened = load_workbook(BytesIO(output_bytes), read_only=False, rich_text=True)
        reopened.close()
    except Exception as exc:
        raise ValueError(f"Working Sheet failed XLSX validation: {exc}") from exc
    st.session_state[f"{session_prefix}_final_bytes"] = output_bytes
    if review["upload_bytes"] is not None:
        st.session_state[f"{session_prefix}_upload_bytes"] = review["upload_bytes"]
    st.session_state[f"{session_prefix}_suffix"] = review["suffix"]
    review["completed"] = True


def _comparison_line(line):
    """Ignore whitespace and colon spacing when comparing original and draft."""
    return re.sub(r"\s+", "", line).casefold()


REVIEW_PANEL_HEIGHT = 335


def _review_panel(body, *, empty_label="(No text)", panel_height=None):
    """Matching, independent, fixed-height panels for both review columns."""
    content = body if body else html.escape(empty_label)
    actual_height = panel_height or REVIEW_PANEL_HEIGHT
    st.html(
        '<div style="box-sizing:border-box;height:' + str(actual_height) + 'px;'
        'overflow-y:auto;overflow-x:hidden;background:#ffffff;color:#172033;'
        'border:1px solid #aebccd;border-radius:10px;padding:16px;'
        'user-select:text;scrollbar-gutter:stable">'
        '<pre style="white-space:pre-wrap;overflow-wrap:anywhere;'
        'font-family:ui-monospace,SFMono-Regular,Consolas,monospace;'
        'font-size:13px;line-height:1.6;margin:0;color:inherit">'
        + content + '</pre></div>'
    )


def colored_reference(original, formatted, panel_height=None, format_style="default"):
    """Use the same source-line classification as non-editing preview."""
    highlighted_lines = []
    in_quality = False
    available_counts = Counter(compare_normalize(line) for line in plain_review_text(formatted).splitlines() if line.strip())
    for line in original.splitlines():
        if format_style.startswith("nawah") and re.match(
                r"^\s*QUALITY REQUIREMENTS?\s*:?\s*$", line, re.I):
            in_quality = True
        cls = _classification(line, formatted, format_style)
        key = compare_normalize(line)
        if cls == "matched" and key:
            if available_counts[key] > 0:
                available_counts[key] -= 1
            else:
                cls = "missing"
        if in_quality and format_style.startswith("nawah") and cls == "missing":
            cls = "excluded"
        style = {
            "missing": "color:#bd1e32;font-weight:600",
            "excluded": "color:#0868a8",
            "matched": "color:#087735",
            "normal": "color:#172033",
        }.get(cls, "color:#172033")
        highlighted_lines.append(
            f'<span style="{style}">{html.escape(line)}</span>')
    content = "\n".join(highlighted_lines)
    if original.endswith("\n"):
        content += "\n"
    _review_panel(content, empty_label="(No original Internal Note)", panel_height=panel_height)


def formatted_reference(formatted):
    """Use precisely the same visual container as the source reference."""
    _review_panel(html.escape(formatted), empty_label="(Empty formatted text)")


@st.dialog("Confirm Verify All", width="small")
def confirm_verify_all_popup(session_prefix):
    review = st.session_state.get(f"{session_prefix}_review")
    if not review or review.get("completed"):
        return
    st.markdown('<div id="rfq-confirm-modal"></div>', unsafe_allow_html=True)
    st.warning("Confirm to skip the verification process at your own risk")
    st.caption("This will approve every remaining line without individual review.")
    yes_col, no_col = st.columns([1, 1], gap="small")
    with yes_col:
        if st.button("Confirm", type="primary", key=f"confirm_all_{session_prefix}", use_container_width=True):
            queue = review["queue"]
            pos = review["position"]
            item = queue[pos]
            if item.get("tabular_pdf"):
                item["desc"] = st.session_state.get(f"desc_{session_prefix}_{pos}", item["desc"]).strip()
            if item["editing"]:
                item["formatted"] = st.session_state.get(f"editable_{session_prefix}_{pos}", item["formatted"])
                item["manufacturer"] = st.session_state.get(f"editablemaker_{session_prefix}_{pos}", item["manufacturer"])
            for record in queue:
                record["verified"] = True
                record["editing"] = False
            review["confirm_all"] = False
            review["verification_skipped"] = True
            finalize_verified(session_prefix)
            st.rerun()
    with no_col:
        if st.button("Cancel", key=f"cancel_all_{session_prefix}", use_container_width=True):
            review["confirm_all"] = False
            review["reopen_review"] = True
            st.rerun()


def _review_undo(prefix):
    review = st.session_state[f"{prefix}_review"]
    pos = review["position"]
    item = review["queue"][pos]
    if not item["history"]:
        return
    text, maker, desc = item["history"].pop()
    item.update(formatted=text, manufacturer=maker, desc=desc, verified=False)
    st.session_state[f"formatted_{prefix}_{pos}"] = text
    st.session_state[f"maker_{prefix}_{pos}"] = maker
    st.session_state[f"editable_{prefix}_{pos}"] = editor_bold_markers(text)
    st.session_state[f"editablemaker_{prefix}_{pos}"] = maker
    if item.get("tabular_pdf"):
        st.session_state[f"desc_{prefix}_{pos}"] = desc


def _review_reset(prefix):
    review = st.session_state[f"{prefix}_review"]
    pos = review["position"]
    item = review["queue"][pos]
    current = (st.session_state.get(f"editable_{prefix}_{pos}",item["formatted"]) if item["editing"] else item["formatted"])
    item["history"].append((current, item["manufacturer"],item["desc"]))
    item.update(formatted=item["initial_formatted"],manufacturer=item["initial_manufacturer"],
                desc=item["initial_desc"],verified=False)
    st.session_state[f"formatted_{prefix}_{pos}"] = item["formatted"]
    st.session_state[f"maker_{prefix}_{pos}"] = item["manufacturer"]
    st.session_state[f"editable_{prefix}_{pos}"] = editor_bold_markers(item["formatted"])
    st.session_state[f"editablemaker_{prefix}_{pos}"] = item["manufacturer"]
    if item.get("tabular_pdf"):
        st.session_state[f"desc_{prefix}_{pos}"] = item["desc"]


@st.dialog("RFQ Quality Review", width="large")
def review_popup(session_prefix):
    review = st.session_state.get(f"{session_prefix}_review")
    if not review or review.get("completed"):
        return
    queue = review["queue"]
    position = review["position"]
    item = queue[position]
    done = sum(q["verified"] for q in queue)
    st.markdown(f"#### Line {item['line']}  ·  Item {position + 1} of {len(queue)}")
    st.caption(f"**Detected RFQ format:** {review.get('format_label', 'Unknown')}  ·  {done}/{len(queue)} verified")
    st.progress(done / len(queue))
    text_key = f"formatted_{session_prefix}_{position}"
    maker_key = f"maker_{session_prefix}_{position}"
    edit_text_key = f"editable_{session_prefix}_{position}"
    edit_maker_key = f"editablemaker_{session_prefix}_{position}"
    desc_key = f"desc_{session_prefix}_{position}"
    st.session_state.setdefault(text_key, item["formatted"])
    st.session_state.setdefault(maker_key, item["manufacturer"])
    if item.get("tabular_pdf"):
        st.session_state.setdefault(desc_key, str(item["desc"]))
    draft = (st.session_state.get(edit_text_key, item["formatted"]) if item["editing"]
             else st.session_state.get(text_key, item["formatted"]))
    desc = st.session_state.get(desc_key, item["desc"]) if item.get("tabular_pdf") else ""
    compared = plain_review_text(draft)
    issues = validate_rfq(item["source"], plain_review_text(draft), item.get("style","default"), desc)
    critical = sum(x["priority"] == "Critical" for x in issues)
    a,b,c = st.columns([1,1,1])
    with a:
        st.markdown(f'<div class="rfq-review-stat"><div class="rfq-stat-label rfq-orange">Technical warnings</div><div class="rfq-stat-count rfq-orange">{len(issues)}</div></div>', unsafe_allow_html=True)
    with b:
        st.markdown(f'<div class="rfq-review-stat"><div class="rfq-stat-label rfq-red">Critical warnings</div><div class="rfq-stat-count rfq-red">{critical}</div></div>', unsafe_allow_html=True)
    with c:
        st.markdown(f'<div class="rfq-review-stat"><div class="rfq-stat-label">Remaining</div><div class="rfq-stat-count">{sum(not q["verified"] for q in queue)}</div></div>', unsafe_allow_html=True)
    # Separate the warning cards from the scrolling toggle. Streamlit Markdown
    # blocks may size slightly shorter than their internal card on some browsers.
    st.markdown('<div style="height:16px;clear:both" aria-hidden="true"></div>', unsafe_allow_html=True)
    scroll_key = f"scrollsync_{session_prefix}"
    st.session_state.setdefault(scroll_key, True)
    sync = st.toggle("Synchronize preview scrolling", key=scroll_key,
                     help="Applies to the read-only two-panel comparison. During editing, the editor scrolls independently.")
    action_help = {
        "edit": "Ctrl + E",
        "undo": "Undo last saved change",
        "reset": "Restore the original Python-formatted draft",
        "prev": "Alt + ←",
        "next": "Alt + →",
        "verify": "Ctrl + Enter",
        "verify_all": "Skip remaining items after confirmation",
    }
    # Keep controls above the RIGHT panel only. The empty left toolbar cell
    # balances vertical space so both preview panels begin at one baseline.
    toolbar_left, toolbar_right = st.columns([1, 1], gap="medium")
    with toolbar_left:
        st.markdown("<div style='height:38px'></div>", unsafe_allow_html=True)
    with toolbar_right:
        tool_edit, tool_undo, tool_reset, tool_space = st.columns([1, 1, 1, 1.6], gap="small")
        with tool_edit:
            if st.button("Edit" if not item["editing"] else "Editing",
                         disabled=item["editing"], key=f"edit_{session_prefix}",
                         help=action_help["edit"], use_container_width=True):
                _review_enter_edit(session_prefix)
                st.rerun(scope="fragment")
        with tool_undo:
            st.button("Undo", key=f"undo_{session_prefix}", disabled=not item["history"],
                      on_click=_review_undo, args=(session_prefix,),
                      help=action_help["undo"], use_container_width=True)
        with tool_reset:
            st.button("Reset", key=f"reset_{session_prefix}",
                      on_click=_review_reset, args=(session_prefix,),
                      help=action_help["reset"], use_container_width=True)
    if item["editing"]:
        # Ensure the widget has the complete saved draft before Streamlit creates it.
        if edit_text_key not in st.session_state or (
            not str(st.session_state[edit_text_key]).strip() and str(item["formatted"]).strip()
        ):
            st.session_state[edit_text_key] = editor_bold_markers(item["formatted"])
        st.session_state[edit_text_key] = repair_split_bold_markers(st.session_state[edit_text_key])
        st.session_state.setdefault(edit_maker_key, item["manufacturer"])
        left_panel, right_panel = st.columns([1, 1], gap="medium")
        with left_panel:
            st.markdown("**Original RFQ** · selectable")
            colored_reference(item["source"], compared,
                              panel_height=REVIEW_PANEL_HEIGHT + (90 if item.get("tabular_pdf") else 0),
                              format_style=item.get("style", "default"))
        with right_panel:
            st.markdown("**Formatted output** · editing")
            if item.get("tabular_pdf"):
                st.text_input("Description", key=desc_key)
            st.text_area("Formatted Internal Note", key=edit_text_key,
                         height=REVIEW_PANEL_HEIGHT + (90 if item.get("tabular_pdf") else 0),
                         label_visibility="collapsed",
                         help="Select any text and press Ctrl+B to toggle bold in the Working Sheet. **markers** are converted into Excel formatting.")
            st.caption("Ctrl+B toggles ** markers around the selected text. Markers become bold in the Working Sheet and are not exported literally.")
    else:
        if item.get("tabular_pdf"):
            st.text_input("Description (first sentence of PDF)", key=desc_key)
        comparison_html(item["source"], compared, item.get("style", "default"), sync,
                        REVIEW_PANEL_HEIGHT + (65 if item.get("tabular_pdf") else 0))
    if issues:
        with st.expander(f"Quality checks · {len(issues)} potential discrepancies", expanded=bool(critical)):
            st.caption("These are rule-based alerts, not confirmed errors. Source exceptions and shortened certificate summaries can trigger false positives.")
            for issue in issues[:20]:
                mark = "🔴" if issue["priority"] == "Critical" else "🟠"
                st.markdown(f"{mark} **{html.escape(issue['field'])}** — `{html.escape(issue['value'])}` · {issue['message']}")
            if len(issues)>20:
                st.caption(f"... plus {len(issues)-20} more checks")
    else:
        st.success("No missing labeled technical values detected by the configured rules.")
    # Navigation controls must share a single row: separate Streamlit column
    # rows can differ in height and misalign the dropdown and buttons.
    nav_col, previous_col, next_col = st.columns([2.3, 1.1, 1.1], gap="small")
    def item_option(i):
        q = queue[i]
        current_desc = st.session_state.get(f"editable_{session_prefix}_{i}", q["formatted"]) if q.get("editing") else q["formatted"]
        warnings = validate_rfq(q["source"], plain_review_text(current_desc), q.get("style", "default"), q.get("desc", ""))
        severity = (" (🔴 Critical Warning)" if any(w["priority"] == "Critical" for w in warnings)
                    else " (🟠 Technical Warning)" if warnings else "")
        status = "" if q["verified"] else " · Pending"
        if q["verified"]:
            severity = ""
        return f"Line {q['line']}{status}{severity}"
    with nav_col:
        st.markdown("<div class='rfq-nav-label'>Jump to line item</div>", unsafe_allow_html=True)
        picked = st.selectbox("Jump to line item", options=list(range(len(queue))), index=position,
             format_func=item_option, label_visibility="collapsed",
             key=f"jump_{session_prefix}_{position}")
        if picked != position:
            _navigate_review(review, session_prefix, picked)
    with previous_col:
        st.markdown("<div class='rfq-nav-label'>&nbsp;</div>", unsafe_allow_html=True)
        if st.button("← Previous", disabled=position == 0, key=f"prev_{session_prefix}",
                     use_container_width=True, help=action_help["prev"]):
            _navigate_review(review, session_prefix, position - 1)
    with next_col:
        st.markdown("<div class='rfq-nav-label'>&nbsp;</div>", unsafe_allow_html=True)
        if st.button("Next →", disabled=position == len(queue) - 1, key=f"next_{session_prefix}",
                     use_container_width=True, help=action_help["next"]):
            _navigate_review(review, session_prefix, position + 1)
    verify_col, skip_col = st.columns([2.1,1])
    with verify_col:
        if st.button("Verify & Next" if position+1<len(queue) else "Verify & Finish",type="primary",
                     key=f"verify_{session_prefix}",use_container_width=True, help=action_help["verify"]):
            updated = (st.session_state.get(edit_text_key,draft) if item["editing"] else st.session_state[text_key]).strip()
            new_desc = st.session_state.get(desc_key,item["desc"]).strip() if item.get("tabular_pdf") else item["desc"]
            if item.get("tabular_pdf") and not new_desc:
                st.error("Description cannot be empty.")
            elif item["source"] and not updated:
                st.error("Formatted text cannot be empty.")
            else:
                new_mfr = (st.session_state.get(edit_maker_key,item["manufacturer"]) if item["editing"]
                           else st.session_state[maker_key]).strip()
                if (updated,new_mfr,new_desc)!=(item["formatted"],item["manufacturer"],item["desc"]):
                    item["history"].append((item["formatted"],item["manufacturer"],item["desc"]))
                item.update(formatted=updated,manufacturer=new_mfr,desc=new_desc,verified=True,editing=False)
                st.session_state[text_key]=updated
                st.session_state[maker_key]=new_mfr
                if all(q["verified"] for q in queue):
                    finalize_verified(session_prefix)
                    st.rerun()
                else:
                    pending=[i for i,q in enumerate(queue) if not q["verified"]]
                    review["position"]=next((i for i in pending if i>position),pending[0])
                    st.rerun(scope="fragment")
    with skip_col:
        if st.button("Verify All",key=f"verify_all_{session_prefix}", use_container_width=True, help=action_help["verify_all"]):
            if _save_current(review,session_prefix):
                review["confirm_all"]=True
                st.rerun()
    st.caption("Shortcuts: when the comparison panel is focused, Alt+←/→ navigates and Ctrl+Enter verifies; Ctrl+E edits. Browser shortcuts can take precedence. Buttons always work.")
    # Register shortcuts in Streamlit's document, not inside a sandboxed iframe.
    # Requires Streamlit 1.52+ for st.html(unsafe_allow_javascript=True).
    shortcut_js = r"""<script>
    (() => {
      if (window.__rfqShortcutHandler) {
        document.removeEventListener('keydown', window.__rfqShortcutHandler, true);
      }
      const handler = (ev) => {
        const dialog = document.querySelector('[role="dialog"]');
        if (!dialog || !dialog.contains(ev.target)) return;
        const node = ev.target;
        const textarea = node instanceof HTMLTextAreaElement ? node : null;
        const ctrl = ev.ctrlKey || ev.metaKey;
        if (ctrl && !ev.altKey && ev.key.toLowerCase() === 'b' && textarea) {
          const start = textarea.selectionStart, end = textarea.selectionEnd;
          if (start === end) return;
          ev.preventDefault();
          ev.stopPropagation();
          let before = textarea.value.slice(0, start);
          let selected = textarea.value.slice(start, end);
          let after = textarea.value.slice(end);
          // Keep trailing whitespace/newlines outside the bold span.
          const tail = (selected.match(/[\s]+$/) || [''])[0];
          if (tail) { selected = selected.slice(0, -tail.length); after = tail + after; }
          const lead = (selected.match(/^[ \t]+/) || [''])[0];
          if (lead) { before += lead; selected = selected.slice(lead.length); }
          if (!selected) return;
          let replacement;
          if (selected.startsWith('**') && selected.endsWith('**') && selected.length >= 4) {
            replacement = selected.slice(2, -2);
          } else if (before.endsWith('**') && after.startsWith('**')) {
            before = before.slice(0, -2);
            after = after.slice(2);
            replacement = selected;
          } else {
            replacement = '**' + selected + '**';
          }
          const value = before + replacement + after;
          // React/Streamlit expects the native setter + a bubbling input event.
          const nativeSetter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set;
          nativeSetter.call(textarea, value);
          textarea.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
          textarea.focus();
          textarea.setSelectionRange(before.length, before.length + replacement.length);
          return;
        }
        const typing = textarea || node instanceof HTMLInputElement || node.isContentEditable;
        if (typing) return;
        let label;
        if (ctrl && ev.key === 'Enter') label = 'Verify &';
        else if (ctrl && ev.key.toLowerCase() === 'e') label = 'Edit';
        else if (ev.altKey && ev.key === 'ArrowLeft') label = 'Previous';
        else if (ev.altKey && ev.key === 'ArrowRight') label = 'Next';
        if (!label) return;
        const button = [...dialog.querySelectorAll('button')]
          .find(el => !el.disabled && el.innerText.trim().includes(label));
        if (button) { ev.preventDefault(); ev.stopPropagation(); button.click(); }
      };
      window.__rfqShortcutHandler = handler;
      document.addEventListener('keydown', handler, true);
    })();
    </script>"""
    st.html(shortcut_js, unsafe_allow_javascript=True)

def review_status(session_prefix):
    review = st.session_state.get(f"{session_prefix}_review")
    if not review:
        return
    if review.get("completed"):
        if review.get("verification_skipped"):
            st.warning("Verification skipped by user confirmation. Please review the Working Sheet before sending.")
        else:
            st.success("All line items verified. Your output files are ready.")
        return
    queue = review["queue"]
    done = sum(1 for row in queue if row["verified"])
    st.info(f"Detected format: {review.get('format_label', 'Unknown')} | Prepared {len(queue)} line items locally. Verified {done}/{len(queue)}. "
            "The Working Sheet becomes available after every line is verified.")
    if st.button("Open verification window / resume", key=f"resume_{session_prefix}"):
        review_popup(session_prefix)
    if st.button("Cancel current verification", key=f"cancel_{session_prefix}"):
        st.session_state.pop(f"{session_prefix}_review", None)
        st.rerun()


def download_xlsx(label, data, filename, key):
    st.download_button(label, data, file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=key)

@st.dialog("Confirm No-Formatting Export", width="small")
def confirm_excel_no_format():
    st.markdown('<span id="rfq-noformat-modal"></span>', unsafe_allow_html=True)
    st.warning("Copy PO Text unchanged and skip quality review?")
    st.caption("Review the exported Working Sheet before use.")
    yes, no = st.columns(2)
    with yes:
        if st.button("Continue", type="primary", key="noformat_continue", use_container_width=True):
            st.session_state["excel_noformat_confirmed"] = True
            st.rerun()
    with no:
        if st.button("Cancel", key="noformat_cancel", use_container_width=True):
            st.rerun()

st.caption("Processing engine: Python rules only | AI/API: Disabled | Every Working Sheet line requires approval")

col1, col2 = st.columns([1, 1], gap="medium")

with col1:
    st.subheader("Excel Data Processor")
    techno = st.file_uploader("RFQ file (.xls, .xlsx, .csv)", type=["xls", "xlsx", "csv"], key="techno")
    with st.expander("Excel templates", expanded=False):
        upload_tpl = st.file_uploader("Upload template (.xlsx)", type=["xlsx"], key="tpl_upl")
        final_tpl = st.file_uploader("Final Sheet template (.xlsx)", type=["xlsx"], key="tpl_fin")
    suffix = st.text_input("Output name suffix", key="suffix_excel")
    regular_excel_col, raw_excel_col = st.columns(2)
    with regular_excel_col:
        process_excel_clicked = st.button("Process Excel", key="btn_excel", use_container_width=True)
    with raw_excel_col:
        no_format_clicked = st.button("Process Excel (No Formatting)", key="btn_excel_raw", use_container_width=True)
    if no_format_clicked:
        if techno is None or not suffix.strip():
            st.warning("Upload an RFQ file and enter an output name suffix first.")
        else:
            confirm_excel_no_format()
    no_format_export = st.session_state.pop("excel_noformat_confirmed", False)
    if process_excel_clicked or no_format_export:
        if techno is None or not suffix.strip():
            st.warning("Upload the Techno-Commercial Envelope and enter an output suffix.")
        else:
            try:
                rfq_match = re.search(r"\d+", techno.name)
                rfq_number = rfq_match.group() if rfq_match else "Unknown"
                if techno.name.lower().endswith(".csv"):
                    csv_records = parse_dubai_petroleum_csv(techno)
                    upload_wb = load_workbook(resolve_template(upload_tpl, "Uploading Sheet.xlsx"))
                    upload_sheet = upload_wb.active
                    clear_template_rows(upload_sheet)
                    records = []
                    for row_index, rec in enumerate(csv_records, start=2):
                        values = {"A": "", "B": rec["RFx Item No"], "C": "",
                                  "D": rec["Material No"], "E": rec["Description"],
                                  "F": rec["PO Text"], "G": normalize_qty(rec["QTY"]),
                                  "H": rec["UOM"], "I": ""}
                        for column, value in values.items():
                            upload_sheet[f"{column}{row_index}"] = value
                        records.append((rec["RFx Item No"], rec["Description"],
                                        rec["QTY"], rec["UOM"], rec["PO Text"]))
                    selected_style = "dubai_csv"
                else:
                    book = pd.ExcelFile(techno)
                    is_nawah_excel = False
                    # Ariba uses Description/InternalNote/Quantity/Unit of Measure.
                    # Nawah uses a different table layout and long PO text cells.
                    required = {"Description", "InternalNote", "Quantity", "Unit of Measure"}
                    matching = None
                    for name in book.sheet_names:
                        headings = set(map(str, pd.read_excel(book, sheet_name=name, nrows=1).columns))
                        if required.issubset(headings):
                            matching = name
                            break
                    if matching is None:
                        is_nawah_excel = True
                        nawah_records = parse_nawah_rfq_excel(book, techno.name)
                        upload_wb = load_workbook(resolve_template(upload_tpl, "Uploading Sheet.xlsx"))
                        upload_sheet = upload_wb.active
                        clear_template_rows(upload_sheet)
                        records = []
                        for i, rec in enumerate(nawah_records, 2):
                            for col, val in {"A":rec['RFx Number'], "B":rec['RFx Item No'], "C":rec['PR Item No'],
                                             "D":rec['Material No'], "E":rec['Description'], "F":rec['PO Text'],
                                             "G":normalize_qty(rec['QTY']), "H":rec['UOM'], "I":rec.get('ARIBA Number', '')}.items():
                                upload_sheet[f"{col}{i}"] = val
                            records.append((rec['RFx Item No'],rec['Description'],rec['QTY'],rec['UOM'],rec['PO Text']))
                    else:
                        df = pd.read_excel(book, sheet_name=matching, keep_default_na=False)
                        description = df["Description"].astype(str).str.strip().str.lower()
                        qty = df["Quantity"].astype(str).str.strip()
                        uom = df["Unit of Measure"].astype(str).str.strip()
                        valid = df[description.ne("item or lot description") & qty.ne("") & uom.ne("") & uom.str.lower().ne("unit of measure")]
                        if valid.empty:
                            raise ValueError("No valid RFQ line items found")
                        upload_wb = load_workbook(resolve_template(upload_tpl, "Uploading Sheet.xlsx"))
                        upload_sheet = upload_wb.active
                        clear_template_rows(upload_sheet)
                        records = []
                        for i, (_, rec) in enumerate(valid.iterrows()):
                            line = (i+1)*10
                            row = i+2
                            values = {"A":rfq_number, "B":line, "E":rec["Description"],
                                      "F":rec["InternalNote"], "G":normalize_qty(rec["Quantity"]),
                                      "H":rec["Unit of Measure"], "I":rec.get("Number", "")}
                            for c, value in values.items():
                                upload_sheet[f"{c}{row}"] = value
                            records.append((line, rec["Description"], rec["Quantity"], rec["Unit of Measure"], rec["InternalNote"]))
                    selected_style = "nawah_excel" if is_nawah_excel else "default"
                wrap_all(upload_sheet)
                prepare_review(records, resolve_template(final_tpl, "Working Sheet.xlsx"), suffix.strip(), "excel", upload_wb, format_style=selected_style, no_format=no_format_export)
                if no_format_export:
                    for item in st.session_state["excel_review"]["queue"]:
                        item["verified"] = True
                    st.session_state["excel_review"]["verification_skipped"] = True
                    finalize_verified("excel")
                    st.success("Excel files generated without formatting or line-by-line review. Please check the Working Sheet before use.")
                else:
                    st.success("Excel RFQ prepared for verification")
                    review_popup("excel")
            except Exception as exc:
                show_error("Excel processing failed (no incomplete files were saved)", exc)
    review_status("excel")
    if "excel_upload_bytes" in st.session_state:
        s = st.session_state["excel_suffix"]
        download_xlsx("Download Upload file", st.session_state["excel_upload_bytes"], f"Uploading Sheet_{s}.xlsx", "dl_excel_up")
        if "excel_final_bytes" in st.session_state:
            download_xlsx("Download Working Sheet", st.session_state["excel_final_bytes"], f"Working Sheet_{s}.xlsx", "dl_excel_fin")

with col2:
    st.subheader("PDF Data Processor")
    pdf = st.file_uploader("RFQ PDF", type=["pdf"], key="pdf")
    with st.expander("Excel templates", expanded=False):
        raw_tpl = st.file_uploader("Raw template", type=["xlsx"], key="tpl_raw")
        hts_tpl = st.file_uploader("HTS template", type=["xlsx"], key="tpl_hts")
        pdf_final_tpl = st.file_uploader("Final Sheet template", type=["xlsx"], key="tpl_final_pdf")
    hts_number = st.text_input("HTS number", key="hts_no")
    if st.button("Process PDF", key="btn_pdf"):
        if pdf is None or not hts_number.strip():
            st.warning("Upload an RFQ PDF and enter HTS number")
        else:
            try:
                # First handle table-based RFQs (description embedded in first line).
                # Fallback retains support for older Ariba-style PDF layouts.
                data = parse_nawah_rfq_pdf(pdf)
                is_nawah_pdf = data is not None
                if data is None:
                    data = parse_tabular_rfq_pdf(pdf)
                is_tabular_pdf = data is not None and not is_nawah_pdf
                if data is None:
                    pdf.seek(0)
                    full = pdf_text(pdf)
                    body = re.sub(r"(REQUEST FOR QUOTATION[\s\S]*?RFQ Number \d+)", "", full)
                    data = parse_pdf(body, full)
                if not data:
                    raise ValueError("No RFQ items parsed. The PDF may not match the supported RFQ layout or may be scanned.")
                upload_wb = load_workbook(resolve_template(hts_tpl, "Uploading Sheet.xlsx"))
                sheet = upload_wb.active
                clear_template_rows(sheet)
                records = []
                mapping = [("RFx Number", "A"), ("RFx Item No", "B"), ("PR Item No", "C"),
                           ("Material No", "D"), ("Description", "E"), ("PO Text", "F"),
                           ("QTY", "G"), ("UOM", "H")]
                for i, item in enumerate(data, start=2):
                    for field, col in mapping:
                        sheet[f"{col}{i}"] = normalize_qty(item[field]) if field == "QTY" else item[field]
                    records.append((item["RFx Item No"], item["Description"], item["QTY"], item["UOM"], item["PO Text"]))
                wrap_all(sheet)
                prepare_review(records, resolve_template(pdf_final_tpl, "Working Sheet.xlsx"), hts_number.strip(), "pdf", upload_wb,
                               format_style="nawah_pdf" if is_nawah_pdf else "tabular_pdf" if is_tabular_pdf else "default")
                st.success("PDF RFQ prepared for verification")
                review_popup("pdf")
            except Exception as exc:
                show_error("PDF processing failed (no incomplete files were saved)", exc)
    review_status("pdf")
    if "pdf_upload_bytes" in st.session_state:
        s = st.session_state["pdf_suffix"]
        download_xlsx("Download Upload file", st.session_state["pdf_upload_bytes"], f"Uploading Sheet_{s}.xlsx", "dl_pdf_up")
        if "pdf_final_bytes" in st.session_state:
            download_xlsx("Download Working Sheet", st.session_state["pdf_final_bytes"], f"Working Sheet_{s}.xlsx", "dl_pdf_fin")

# Display at most one modal per Streamlit run. Separate dialogs cannot be nested.
for _prefix in ("excel", "pdf"):
    _review = st.session_state.get(f"{_prefix}_review")
    if not _review or _review.get("completed"):
        continue
    if _review.get("confirm_all"):
        confirm_verify_all_popup(_prefix)
        break
    if _review.pop("reopen_review", False):
        review_popup(_prefix)
        break
