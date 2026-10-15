"""RFQ Data Processor - Python-only verification edition.

Preserves the Excel/PDF/HTS/List Maker workflows without external AI APIs.
"""
import re
import html
import traceback
from collections import defaultdict
from copy import copy
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from PyPDF2 import PdfReader
try:
    import fitz  # PyMuPDF: layout-aware extraction for tabular RFQ PDFs
except ImportError:
    fitz = None

st.set_page_config(page_title="RFQ Data Processor", layout="wide", initial_sidebar_state="auto")
st.markdown("""<style>
.stButton button {background:#ff914d;color:white;border-radius:8px;padding:10px 16px;margin-top:10px}
.stExpander {border-radius:10px}
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
PRESERVED_LABELS = {"characteristics"}
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


def format_rfq_python(text):
    """Preservation-first deterministic formatter. Never fabricate or paraphrase values.

    The original text is always retained in the review pane, and the user must
    verify every line. Unrecognized content is included instead of discarded.
    """
    raw = _clean(text).replace("\r\n", "\n").replace("\r", "\n")
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

    # Collapse exact adjacent/nonadjacent duplicates ONLY in a single section.
    # Preserve duplicate PO, basic and inspection lines because they can convey
    # distinct requirements despite identical wording; MPN repeats are commonly
    # duplicated by SAP and are safe to consolidate exactly.
    mpn = sections["MPN Text"]
    seen = set()
    deduped = []
    for line in mpn:
        if line not in seen:
            deduped.append(line)
            seen.add(line)
    sections["MPN Text"] = deduped

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
    sections["MPN Text"] = separated_mpn

    chunks = []
    for section in SECTION_ORDER:
        lines = sections[section]
        if lines:
            chunks.append(section + ":\n" + "\n".join(lines))
    # Always preserve relevant data, even when no headers were present.
    output = "\n\n".join(chunks)
    return normalize_pdf_symbols(output), extract_manufacturers(raw)


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

def summarize_nawah_working_note(raw):
    """Preserve Nawah material details and concisely list explicit QA deliverables.

    General commercial QA/packaging boilerplate is omitted from Working Sheet,
    but nuclear safety status and code documents are not silently discarded.
    The unmodified original remains in the Uploading Sheet/review reference.
    """
    source = str(raw or "").replace("\r\n", "\n").replace("\r", "\n")
    source = re.sub(r"\*{4,}\s*;?\s*Material PO text\s*:\s*\*{4,}",
                    "Material PO Text:", source, flags=re.I)
    source = re.sub(r"^\s*\[Material PO Text\]\s*$", "Material PO Text:",
                    source, flags=re.I | re.M)
    # Some Nawah RFQs omit the literal QUALITY REQUIREMENTS header and
    # begin directly with their first START OF MESSAGE CODE.
    cut = re.search(
        r"^\s*(?:QUALITY REQUIREMENTS?\s*:?\s*|START OF MESSAGE CODE\s+Z_NPE_[^\n]+)\s*$",
        source, re.I | re.M)
    if cut is None:
        marker = re.search(r"^\s*\*{8,}\s*$\n\s*START OF MESSAGE CODE", source, re.I | re.M)
        if marker:
            cut = marker
    material = (source[:cut.start()] if cut else source).strip()
    material = re.sub(r"^\s*Material PO Text\s*:\s*", "", material, flags=re.I)
    quality = source[cut.start():] if cut else ""

    # Do not infer a certificate from generic boilerplate mentions such as
    # traceability of sub-tier certifications. Detect explicit named requirements.
    definitions = [
        (r"Z_NPE_CERT OF CONFORMANCE|^\s*CERTIFICATE OF CONFORMANCE\s*$", "CERTIFICATE OF CONFORMANCE"),
        (r"Z_NPE_CERT OF COMPLIANCE|^\s*CERTIFICATE OF COMPLIANCE\s*$", "CERTIFICATE OF COMPLIANCE"),
        (r"Z_NPE_CERT OF CONFORMITY|^\s*CERTIFICATE OF CONFORMITY\s*$", "CERTIFICATE OF CONFORMITY"),
        (r"Z_NPE_CODE CMTR|CODE CERTIFIED MATERIAL TEST REPORT", "CERTIFIED MATERIAL TEST REPORT (CMTR)"),
        (r"Z_NPE_CMTR|MATERIAL TEST REPORT \(MTR\) REQUIREMENTS", "MATERIAL TEST REPORT (MTR)"),
        (r"Z_NPE_CODE DATA REPORT|^\s*CODE DATA REPORT\s*$", "CODE DATA REPORT"),
        (r"Z_NPE_CODE RECONCILIATION|CODE RECONCILIATION REQUIREMENTS", "CODE RECONCILIATION (IF REQUIRED)"),
        (r"CALIBRATION CERTIFICATE|CERTIFICATE OF CALIBRATION", "CALIBRATION CERTIFICATE"),
        (r"MATERIAL TEST CERTIFICATE|\bMTC\b", "MATERIAL TEST CERTIFICATE"),
        (r"\bTEST CERTIFICATE\b", "TEST CERTIFICATE"),
    ]
    certificates = [label for pattern, label in definitions if re.search(pattern, quality, re.I | re.M)]
    nuclear_safety = bool(re.search(r"Z_NPE_QA\s+SR\s+MATERIAL|\bNUCLEAR SAFETY RELATED\b", quality, re.I))
    packing_slip = bool(re.search(r"Packing Slip or Equivalent Commercial Certification is required", quality, re.I))
    clauses = []
    if nuclear_safety:
        clauses.append("NUCLEAR SAFETY-RELATED MATERIAL: NAWAH-approved nuclear QA requirements apply; review full source requirements.")
    if packing_slip:
        clauses.append("PACKING SLIP OR EQUIVALENT COMMERCIAL CERTIFICATION REQUIRED.")
    clauses.extend(dict.fromkeys(certificates))
    result = "Material PO Text:\n\n" + material if material else "Material PO Text:"
    if clauses:
        result += "\n\nQUALITY REQUIREMENTS\n\n" + "\n".join(clauses)
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
            marker = next((marker_re.fullmatch(v) for v in row if marker_re.fullmatch(v)), None)
            ordinal = int(marker.group(1)) if marker else len(records) + 1
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
                "RFx Item No": str(ordinal * 10),
                "PR Item No": "",
                "Material No": material,
                "Description": description,
                "PO Text": note,
                "QTY": qty,
                "UOM": unit,
                "ARIBA Number": f"6.{ordinal}" if marker else str(ordinal),
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
    return raw, extract_manufacturers(raw)


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


def join_pdf_wrapped_lines(note):
    """Join wrapped narrative values, retaining one line per named attribute.

    Never silently rewrite a technical value. In particular, broken words such
    as 'MA' + 'TERIAL' remain explicit for human verification unless the split
    is unambiguous and configured separately.
    """
    lines = [normalize_pdf_symbols(line.strip()) for line in str(note).splitlines()]
    heading = re.compile(r"^(?:BASIC DATA TEXT|CLASSIFICATION TEXT|MPN TEXT|INSPECTION TEXT|PO MATERIAL TEXT|MPN DETAILS|CHARACTERISTICS|TAG NOS)\s*:?$", re.I)
    attribute = re.compile(r"^[A-Z][A-Z0-9 /()_.#-]{1,70}\s*:\s*\S", re.I)
    narrative_label = re.compile(
        r"^(?:ADDITIONAL INFORMATION(?:\s*\(CONTD\d+\))?|PRE-QUOTE\d*|PRE-SHIPMENT\d*|PRE-ORDER CONDITION\d*|STANDARD|SPECIFICATION|REMARKS?|NOTES?|DRAWING NUMBER|HARDNESS)\s*:", re.I)
    out = []
    joining = False
    for line in lines:
        if not line:
            if out and out[-1] != "":
                out.append("")
            joining = False
        elif heading.fullmatch(line):
            out.append(line)
            joining = False
        elif attribute.match(line):
            out.append(line)
            joining = bool(narrative_label.match(line))
        elif out and out[-1] and joining:
            out[-1] = out[-1].rstrip() + " " + line
        else:
            out.append(line)
    return "\n".join(out).strip()

def format_tabular_pdf_note(note):
    """Readable PDF review draft; keep technical fields separate, join prose wraps."""
    physical = [line.strip() for line in str(note).splitlines() if line.strip()]
    if not physical:
        return "", ""

    formatted = join_pdf_wrapped_lines(note)
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
    return "\n".join(output), manufacturer


def prepare_review(records, template, suffix, session_prefix, upload_wb=None, format_style="default"):
    """Pre-format all rows instantly, queue for mandatory per-row verification."""
    if not records:
        raise ValueError("No RFQ items to process")
    # Validate the template before initializing a review session.
    template_wb = load_workbook(template)
    template_data = workbook_bytes(template_wb)
    queue = []
    for line, desc, qty, uom, note in records:
        formatted, manufacturer = (format_tabular_pdf_note(note)
                                   if format_style == "tabular_pdf"
                                   else summarize_nawah_working_note(note) if format_style in ("nawah_pdf", "nawah_excel")
                                   else process_rfq_text(note))
        is_table = format_style == "tabular_pdf"
        # On this PDF layout the first line is the description. Show that full
        # original item in the left reference, but keep PO Text separate for export.
        source = (str(desc).strip() + "\n" + _clean(note)).strip() if is_table else _clean(note)
        queue.append({"line": line, "desc": desc, "qty": qty, "uom": uom,
                      "source": source, "formatted": formatted,
                      "tabular_pdf": is_table, "manufacturer": manufacturer,
                      "verified": False, "editing": False})
    for old_key in list(st.session_state):
        if old_key.startswith((f"original_{session_prefix}_", f"formatted_{session_prefix}_", f"maker_{session_prefix}_", f"editable_{session_prefix}_", f"editablemaker_{session_prefix}_", f"desc_{session_prefix}_")):
            del st.session_state[old_key]
    st.session_state[f"{session_prefix}_review"] = {
        "queue": queue, "position": 0, "suffix": suffix,
        "template": template_data,
        "upload_bytes": workbook_bytes(upload_wb) if upload_wb is not None else None,
    }
    # Uploading Sheet has no AI formatting and can be downloaded immediately.
    st.session_state.pop(f"{session_prefix}_final_bytes", None)
    st.session_state[f"{session_prefix}_suffix"] = suffix
    if upload_wb is not None:
        st.session_state[f"{session_prefix}_upload_bytes"] = workbook_bytes(upload_wb)
    else:
        st.session_state.pop(f"{session_prefix}_upload_bytes", None)


def finalize_verified(session_prefix):
    review = st.session_state[f"{session_prefix}_review"]
    queue = review["queue"]
    if not queue or not all(item["verified"] for item in queue):
        raise ValueError("All RFQ items must be verified before exporting")
    wb = load_workbook(BytesIO(review["template"]))
    sheet = wb.active
    clear_template_rows(sheet)
    for i, item in enumerate(queue, 2):
        for col, value in {
            "A": item["line"], "B": item["desc"], "C": item["qty"],
            "D": item["uom"], "E": item["formatted"],
            "G": item["manufacturer"],
        }.items():
            sheet[f"{col}{i}"] = value
    sheet["E1"] = "Internal Note"
    sheet.delete_cols(8, 1)  # mail id
    sheet.delete_cols(6, 1)  # Subject line; Manufacturer moves to F
    for other in list(wb.worksheets):
        if other is not sheet:
            wb.remove(other)
    wrap_all(sheet)
    st.session_state[f"{session_prefix}_final_bytes"] = workbook_bytes(wb)
    if review["upload_bytes"] is not None:
        st.session_state[f"{session_prefix}_upload_bytes"] = review["upload_bytes"]
    st.session_state[f"{session_prefix}_suffix"] = review["suffix"]
    review["completed"] = True


def _comparison_line(line):
    """Ignore whitespace and colon spacing when comparing original and draft."""
    return re.sub(r"\s+", "", line).casefold()


REVIEW_PANEL_HEIGHT = 510


def _review_panel(body, *, empty_label="(No text)", panel_height=None):
    """Matching, independent, fixed-height panels for both review columns."""
    content = body if body else html.escape(empty_label)
    actual_height = panel_height or REVIEW_PANEL_HEIGHT
    st.html(
        '<div style="box-sizing:border-box;height:' + str(actual_height) + 'px;'
        'overflow-y:auto;overflow-x:hidden;background:#1a1d26;color:#f4f4f5;'
        'border:1px solid #363b48;border-radius:10px;padding:16px;'
        'user-select:text;scrollbar-gutter:stable">'
        '<pre style="white-space:pre-wrap;overflow-wrap:anywhere;'
        'font-family:ui-monospace,SFMono-Regular,Consolas,monospace;'
        'font-size:13px;line-height:1.6;margin:0;color:inherit">'
        + content + '</pre></div>'
    )


def colored_reference(original, formatted, panel_height=None):
    """Original text with preserved whitespace and red unmatched lines."""
    formatted_lines = [_comparison_line(x) for x in formatted.splitlines() if x.strip()]
    highlighted_lines = []
    for line in original.splitlines():
        token = _comparison_line(line)
        matched = not token or any(
            token == candidate or (len(token) > 16 and token in candidate)
            for candidate in formatted_lines
        )
        safe_line = html.escape(line)
        if not matched:
            safe_line = f'<span style="color:#ff7373;font-weight:600">{safe_line}</span>'
        highlighted_lines.append(safe_line)
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
    st.warning("Confirm to skip the verification process at your own risk")
    st.caption("This will approve every remaining line without individual review.")
    yes_col, no_col = st.columns(2)
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


@st.dialog("Verify RFQ line item", width="large")
def review_popup(session_prefix):
    review = st.session_state.get(f"{session_prefix}_review")
    if not review or review.get("completed"):
        return
    queue = review["queue"]
    position = review["position"]
    item = queue[position]
    st.write(f"**Line {item['line']} — item {position + 1} of {len(queue)}**")
    st.progress(sum(1 for record in queue if record["verified"]) / len(queue))
    text_key = f"formatted_{session_prefix}_{position}"
    maker_key = f"maker_{session_prefix}_{position}"
    if text_key not in st.session_state:
        st.session_state[text_key] = item["formatted"]
    if maker_key not in st.session_state:
        st.session_state[maker_key] = item["manufacturer"]
    edit_text_key = f"editable_{session_prefix}_{position}"
    edit_maker_key = f"editablemaker_{session_prefix}_{position}"
    desc_key = f"desc_{session_prefix}_{position}"
    if item.get("tabular_pdf") and desc_key not in st.session_state:
        st.session_state[desc_key] = str(item["desc"])
    left, right = st.columns([1, 1], gap="medium", vertical_alignment="top")
    with left:
        st.markdown("**Original text** · :gray[Read-only / selectable]")
        active_draft = (st.session_state.get(edit_text_key, st.session_state[text_key])
                        if item["editing"] else st.session_state[text_key])
        comparison_text = ((st.session_state.get(desc_key, item["desc"]) + "\n")
                           if item.get("tabular_pdf") else "") + active_draft
        colored_reference(item["source"], comparison_text,
                          panel_height=REVIEW_PANEL_HEIGHT + 98 if item.get("tabular_pdf") else REVIEW_PANEL_HEIGHT)
        st.caption("Red = source line not found in the formatted text (check before approval).")
    with right:
        if item.get("tabular_pdf"):
            st.text_input("Description (first line of PDF item)", key=desc_key,
                          help="Only for table-style PDFs. Edit this value to correct the Working Sheet Description column.")
            st.caption("The original description remains visible on the left.")
        st.markdown("**Formatted PO Text** · :gray[" + ("Editing" if item["editing"] else "Preview") + "]")
        if item["editing"]:
            # Separate widget keys prevent Streamlit from resetting the stored
            # draft when switching a read-only code block into an editor.
            if edit_text_key not in st.session_state:
                st.session_state[edit_text_key] = st.session_state[text_key]
            if edit_maker_key not in st.session_state:
                st.session_state[edit_maker_key] = st.session_state[maker_key]
            st.text_area("Formatted Internal Note", height=REVIEW_PANEL_HEIGHT, key=edit_text_key,
                         label_visibility="collapsed")
            st.text_input("Manufacturer", key=edit_maker_key)
        else:
            formatted_reference(st.session_state[text_key])
            st.caption("Manufacturer: " + (st.session_state[maker_key] or "(Not identified)"))
    if not item["source"]:
        st.info("This item has an empty source note. Verify it if the blank note is intentional.")
    st.caption("Click Edit to correct the right-hand text or manufacturer, then click Verify. Every line requires explicit approval.")
    back_col, skip_col, edit_col, confirm_col, all_col = st.columns([1, 1.5, 1, 2, 1.2])
    with back_col:
        if st.button("Previous", disabled=position == 0, key=f"back_{session_prefix}"):
            item["formatted"] = st.session_state.get(edit_text_key, st.session_state[text_key]) if item["editing"] else st.session_state[text_key]
            item["manufacturer"] = st.session_state.get(edit_maker_key, st.session_state[maker_key]) if item["editing"] else st.session_state[maker_key]
            if item.get("tabular_pdf"):
                item["desc"] = st.session_state.get(desc_key, item["desc"])
            review["position"] -= 1
            st.rerun(scope="fragment")
    with skip_col:
        if st.button("Close / resume later", key=f"close_{session_prefix}"):
            item["formatted"] = st.session_state.get(edit_text_key, st.session_state[text_key]) if item["editing"] else st.session_state[text_key]
            item["manufacturer"] = st.session_state.get(edit_maker_key, st.session_state[maker_key]) if item["editing"] else st.session_state[maker_key]
            if item.get("tabular_pdf"):
                item["desc"] = st.session_state.get(desc_key, item["desc"])
            st.rerun()
    with edit_col:
        if st.button("Edit" if not item["editing"] else "Editing", key=f"edit_{session_prefix}",
                     disabled=item["editing"]):
            # Copy the visible draft into dedicated edit fields before
            # switching widgets, so the editor never starts out empty.
            st.session_state[edit_text_key] = st.session_state[text_key]
            st.session_state[edit_maker_key] = st.session_state[maker_key]
            item["editing"] = True
            # Re-render only this dialog, not the full app, so it stays open.
            st.rerun(scope="fragment")
    with confirm_col:
        if st.button("Verify & Next" if position + 1 < len(queue) else "Verify & Finish",
                     type="primary", key=f"verify_{session_prefix}", use_container_width=True):
            updated = (st.session_state.get(edit_text_key, st.session_state[text_key]) if item["editing"] else st.session_state[text_key]).strip()
            desc_value = st.session_state.get(desc_key, item["desc"]).strip() if item.get("tabular_pdf") else item["desc"]
            if item.get("tabular_pdf") and not desc_value:
                st.error("Description cannot be empty for this PDF layout.")
            elif item["source"] and not updated:
                st.error("Formatted text cannot be empty when the original text has content.")
            else:
                item["desc"] = desc_value
                item["formatted"] = updated
                item["manufacturer"] = (st.session_state.get(edit_maker_key, st.session_state[maker_key]) if item["editing"] else st.session_state[maker_key]).strip()
                item["verified"] = True
                item["editing"] = False
                if all(row["verified"] for row in queue):
                    finalize_verified(session_prefix)
                    st.rerun()
                else:
                    pending = [i for i, row in enumerate(queue) if not row["verified"]]
                    review["position"] = next((i for i in pending if i > position), pending[0])
                    st.rerun(scope="fragment")
    with all_col:
        if st.button("Verify All", key=f"verify_all_{session_prefix}"):
            review["confirm_all"] = True
            # Full rerun closes review dialog before opening a separate dialog.
            st.rerun()


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
    st.info(f"Prepared {len(queue)} line items locally. Verified {done}/{len(queue)}. "
            "The Working Sheet becomes available after every line is verified.")
    if st.button("Open verification window / resume", key=f"resume_{session_prefix}"):
        review_popup(session_prefix)
    if st.button("Cancel current verification", key=f"cancel_{session_prefix}"):
        st.session_state.pop(f"{session_prefix}_review", None)
        st.rerun()


def download_xlsx(label, data, filename, key):
    st.download_button(label, data, file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=key)

st.caption("Processing engine: Python rules only | AI/API: Disabled | Every Working Sheet line requires approval")

col1, col2, col3 = st.columns([2, 2, 1])

with col1:
    st.subheader("Excel Data Processor")
    techno = st.file_uploader("Techno-Commercial Envelope (.xls)", type=["xls"], key="techno")
    with st.expander("Excel templates", expanded=True):
        upload_tpl = st.file_uploader("Upload template (.xlsx)", type=["xlsx"], key="tpl_upl")
        final_tpl = st.file_uploader("Final Sheet template (.xlsx)", type=["xlsx"], key="tpl_fin")
    suffix = st.text_input("Output name suffix", key="suffix_excel")
    if st.button("Process Excel", key="btn_excel"):
        if techno is None or not suffix.strip():
            st.warning("Upload the Techno-Commercial Envelope and enter an output suffix.")
        else:
            try:
                rfq_match = re.search(r"\d+", techno.name)
                rfq_number = rfq_match.group() if rfq_match else "Unknown"
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
                    upload_wb = load_workbook(resolve_template(upload_tpl, "upload file - HTS.xlsx"))
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
                    upload_wb = load_workbook(resolve_template(upload_tpl, "upload file - HTS.xlsx"))
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
                wrap_all(upload_sheet)
                prepare_review(records, resolve_template(final_tpl, "FINAL SHEET.xlsx"), suffix.strip(), "excel", upload_wb, format_style="nawah_excel" if is_nawah_excel else "default")
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
    with st.expander("Excel templates", expanded=True):
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
                upload_wb = load_workbook(resolve_template(hts_tpl, "upload file - HTS.xlsx"))
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
                prepare_review(records, resolve_template(pdf_final_tpl, "FINAL SHEET.xlsx"), hts_number.strip(), "pdf", upload_wb,
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

with col3:
    st.subheader("HTS Cleaner")
    hts_upload = st.file_uploader("upload file - HTS.xlsx", type=["xlsx"], key="hts_clean")
    clean_final_tpl = st.file_uploader("Final Sheet template (optional)", type=["xlsx"], key="tpl_clean_fin")
    if st.button("Clean HTS", key="btn_clean_hts"):
        if hts_upload is None:
            st.warning("Upload the HTS workbook")
        else:
            try:
                source = load_workbook(hts_upload).active
                rows = [r for r in source.iter_rows(min_row=2) if any(c.value is not None for c in r)]
                if not rows:
                    raise ValueError("The uploaded HTS workbook contains no data rows")
                records = []
                for row in rows:
                    # Existing HTS layout: B=line, E=description, F=PO text, G=qty, H=UOM.
                    if len(row) < 8:
                        raise ValueError("HTS workbook must have at least eight columns")
                    records.append((row[1].value, row[4].value, row[6].value, row[7].value, row[5].value or ""))
                prepare_review(records, resolve_template(clean_final_tpl, "FINAL SHEET.xlsx"), "cleaned", "clean")
                st.success("HTS RFQ prepared for verification")
                review_popup("clean")
            except Exception as exc:
                show_error("HTS cleaning failed (no incomplete files were saved)", exc)
    review_status("clean")
    if "clean_final_bytes" in st.session_state:
        download_xlsx("Download cleaned FINAL SHEET", st.session_state["clean_final_bytes"], "Working Sheet_cleaned.xlsx", "dl_clean")

    st.subheader("List Maker")
    final_file = st.file_uploader("FINAL SHEET for manufacturers", type=["xlsx"], key="manuf")
    if st.button("Build list", key="btn_list"):
        if final_file is None:
            st.warning("Upload the FINAL SHEET workbook")
        else:
            try:
                df = pd.read_excel(final_file)
                groups = defaultdict(lambda: {"items":[], "emails":[]})
                for _, row in df.iterrows():
                    manufacturers = _clean(row.get("Manufacturer", ""))
                    if not manufacturers:
                        continue
                    item = row.get("Line item number", "")
                    emails = [row[col] for col in df.columns if ("mail" in str(col).lower() or "unnamed" in str(col).lower()) and pd.notna(row[col])]
                    for name in (x.strip() for x in manufacturers.split("-") if x.strip()):
                        groups[name]["items"].append(item)
                        groups[name]["emails"].extend(emails)
                lines = []
                for manufacturer, values in groups.items():
                    items = sorted({_clean(x) for x in values["items"] if _clean(x)})
                    lines.append(f"Item {', '.join(items)}: {manufacturer}")
                    lines.extend(dict.fromkeys(_clean(x) for x in values["emails"] if _clean(x)))
                    lines.append("")
                st.session_state["manufacturer_list"] = "\n".join(lines)
            except Exception as exc:
                show_error("List creation failed", exc)
    if "manufacturer_list" in st.session_state:
        st.text_area("Output", st.session_state["manufacturer_list"], height=300)
        try:
            from st_copy_to_clipboard import st_copy_to_clipboard
            st_copy_to_clipboard(st.session_state["manufacturer_list"])
        except ImportError:
            st.caption("Select the text above to copy it.")

# Display at most one modal per Streamlit run. Separate dialogs cannot be nested.
for _prefix in ("excel", "pdf", "clean"):
    _review = st.session_state.get(f"{_prefix}_review")
    if not _review or _review.get("completed"):
        continue
    if _review.get("confirm_all"):
        confirm_verify_all_popup(_prefix)
        break
    if _review.pop("reopen_review", False):
        review_popup(_prefix)
        break
