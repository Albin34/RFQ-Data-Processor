"""RFQ Data Processor - Python-only verification edition.

Preserves the Excel/PDF/HTS/List Maker workflows without external AI APIs.
"""
import re
import traceback
from collections import defaultdict
from copy import copy
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from PyPDF2 import PdfReader

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
META_LINES = {"mpn details"}
PRESERVED_LABELS = {"characteristics"}


def is_empty_field(line):
    """Drop only genuinely empty attribute fields; preserve prose and technical values."""
    line = line.strip()
    if not line:
        return True
    if line.casefold() in META_LINES:
        return True
    if ":" in line:
        left, right = line.split(":", 1)
        # Empty field label, including extraction placeholders.
        if left.strip().casefold() in PRESERVED_LABELS:
            return False
        if left.strip() and right.strip().casefold() in EMPTY_MARKERS:
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
    for line in original_lines:
        if not line:
            continue
        header = match_section(line)
        if header:
            active = header
            continue
        if is_empty_field(line):
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

    # Separate different manufacturer/part-number sets with a blank line.
    # A new Manufacturer Part Number following an earlier manufacturer-number
    # line begins another set; never change the original specification values.
    mpn_lines = sections["MPN Text"]
    separated_mpn = []
    previous_set_has_manufacturer = False
    for line in mpn_lines:
        is_part_number = bool(re.match(r"^Manufacturer\s+Part\s+Number\s*:", line, re.I))
        is_manufacturer_number = bool(re.match(r"^Manufacturer\s+number\s*:", line, re.I))
        if is_part_number and previous_set_has_manufacturer and separated_mpn:
            separated_mpn.append("")
            previous_set_has_manufacturer = False
        separated_mpn.append(line)
        if is_manufacturer_number:
            previous_set_has_manufacturer = True
    sections["MPN Text"] = separated_mpn

    chunks = []
    for section in SECTION_ORDER:
        lines = sections[section]
        if lines:
            chunks.append(section + ":\n" + "\n".join(lines))
    # Always preserve relevant data, even when no headers were present.
    output = "\n\n".join(chunks)
    return output, extract_manufacturers(raw)


def process_rfq_text(text):
    return format_rfq_python(text)


def format_text(text):
    return process_rfq_text(text)[0]


def manufacture_name(text):
    return process_rfq_text(text)[1]

def pdf_text(pdf_file):
    reader = PdfReader(pdf_file)
    output = []
    for i, page in enumerate(reader.pages, start=1):
        try:
            output.append(page.extract_text() or "")
        except Exception as exc:
            raise ValueError(f"Cannot extract page {i}: {exc}") from exc
    return "".join(output)

def parse_pdf(body, full_rfq_text):
    rfq_match = re.search(r"RFQ Number (\d+)", full_rfq_text)
    rfq_number = rfq_match.group(1) if rfq_match else "Unknown"
    item_pattern = re.compile(r"(\d{5}) (\w?12\d{10}) (\d+(?:\.\d+)?)\s*(\w+) .*?(\d{2}\.\d{2}\.\d{4})", re.S)
    short_pattern = re.compile(r"Short Text :(.*?)\n", re.S)
    po_pattern = re.compile(r"PO Material Text :(.*?)Agreement / LineNo.", re.S)
    items = item_pattern.findall(body)
    shorts = short_pattern.findall(body)
    pos = po_pattern.findall(body)
    results = []
    for i, item in enumerate(items):
        material = item[1] if item[1].startswith(("B12", "12", "B16", "15")) else ""
        results.append({"RFx Number":rfq_number, "RFx Item No":item[0], "PR Item No":"",
                        "Material No":material, "Description":shorts[i] if i<len(shorts) else "",
                        "PO Text":pos[i] if i<len(pos) else "", "QTY":item[2], "UOM":item[3]})
    return results

def resolve_template(uploaded, default_filename):
    path = uploaded if uploaded is not None else Path(__file__).with_name(default_filename)
    if isinstance(path, Path) and not path.is_file():
        raise FileNotFoundError(f"Missing default template: {path.name}. Please upload a template.")
    return path

def prepare_review(records, template, suffix, session_prefix, upload_wb=None):
    """Pre-format all rows instantly, queue for mandatory per-row verification."""
    if not records:
        raise ValueError("No RFQ items to process")
    # Validate the template before initializing a review session.
    template_wb = load_workbook(template)
    template_data = workbook_bytes(template_wb)
    queue = []
    for line, desc, qty, uom, note in records:
        formatted, manufacturer = process_rfq_text(note)
        queue.append({"line": line, "desc": desc, "qty": qty, "uom": uom,
                      "source": _clean(note), "formatted": formatted,
                      "manufacturer": manufacturer, "verified": False, "editing": False})
    for old_key in list(st.session_state):
        if old_key.startswith((f"original_{session_prefix}_", f"formatted_{session_prefix}_", f"maker_{session_prefix}_", f"editable_{session_prefix}_", f"editablemaker_{session_prefix}_")):
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
    left, right = st.columns(2, gap="medium")
    with left:
        st.markdown("**Original text (reference — selectable, not editable)**")
        # st.code is selectable and does not show the forbidden cursor used by
        # disabled Streamlit text inputs. It also cannot change source text.
        st.code(item["source"] or "(No original Internal Note)", language=None, wrap_lines=True)
    with right:
        st.markdown("**Formatted text**")
        if item["editing"]:
            # Separate widget keys prevent Streamlit from resetting the stored
            # draft when switching a read-only code block into an editor.
            if edit_text_key not in st.session_state:
                st.session_state[edit_text_key] = st.session_state[text_key]
            if edit_maker_key not in st.session_state:
                st.session_state[edit_maker_key] = st.session_state[maker_key]
            st.text_area("Formatted Internal Note", height=390, key=edit_text_key)
            st.text_input("Manufacturer", key=edit_maker_key)
        else:
            st.code(st.session_state[text_key] or "(Empty formatted text)", language=None, wrap_lines=True)
            st.caption("Manufacturer: " + (st.session_state[maker_key] or "(Not identified)"))
    if not item["source"]:
        st.info("This item has an empty source note. Verify it if the blank note is intentional.")
    st.caption("Click Edit to correct the right-hand text or manufacturer, then click Verify. Every line requires explicit approval.")
    back_col, skip_col, edit_col, confirm_col, all_col = st.columns([1, 1.5, 1, 2, 1.2])
    with back_col:
        if st.button("Previous", disabled=position == 0, key=f"back_{session_prefix}"):
            item["formatted"] = st.session_state.get(edit_text_key, st.session_state[text_key]) if item["editing"] else st.session_state[text_key]
            item["manufacturer"] = st.session_state.get(edit_maker_key, st.session_state[maker_key]) if item["editing"] else st.session_state[maker_key]
            review["position"] -= 1
            st.rerun(scope="fragment")
    with skip_col:
        if st.button("Close / resume later", key=f"close_{session_prefix}"):
            item["formatted"] = st.session_state.get(edit_text_key, st.session_state[text_key]) if item["editing"] else st.session_state[text_key]
            item["manufacturer"] = st.session_state.get(edit_maker_key, st.session_state[maker_key]) if item["editing"] else st.session_state[maker_key]
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
            if item["source"] and not updated:
                st.error("Formatted text cannot be empty when the original text has content.")
            else:
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
            st.rerun(scope="fragment")
    if review.get("confirm_all"):
        st.warning("Confirm to skip the verification process at your own risk")
        yes_col, no_col = st.columns(2)
        with yes_col:
            if st.button("Confirm", type="primary", key=f"confirm_all_{session_prefix}"):
                # Preserve any current edits before confirming the entire queue.
                if item["editing"]:
                    item["formatted"] = st.session_state.get(edit_text_key, item["formatted"])
                    item["manufacturer"] = st.session_state.get(edit_maker_key, item["manufacturer"])
                for record in queue:
                    record["verified"] = True
                    record["editing"] = False
                review["confirm_all"] = False
                review["verification_skipped"] = True
                finalize_verified(session_prefix)
                st.rerun()
        with no_col:
            if st.button("Cancel", key=f"cancel_all_{session_prefix}"):
                review["confirm_all"] = False
                st.rerun(scope="fragment")


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
                required = {"Description", "InternalNote", "Quantity", "Unit of Measure"}
                matching = None
                for name in book.sheet_names:
                    if required.issubset(set(pd.read_excel(book, sheet_name=name, nrows=1).columns)):
                        matching = name
                        break
                if matching is None:
                    raise ValueError(f"No sheet contains the required columns: {', '.join(sorted(required))}")
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
                prepare_review(records, resolve_template(final_tpl, "FINAL SHEET.xlsx"), suffix.strip(), "excel", upload_wb)
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
                prepare_review(records, resolve_template(pdf_final_tpl, "FINAL SHEET.xlsx"), hts_number.strip(), "pdf", upload_wb)
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
