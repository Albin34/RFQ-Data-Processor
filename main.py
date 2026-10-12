"""RFQ Data Processor - Groq edition.

Requires GROQ_API_KEY in Streamlit secrets or environment variables.
Preserves the repository's Excel/PDF/HTS/List Maker workflows.
"""
import functools
import json
import math
import os
import re
import threading
import time
import traceback
from collections import defaultdict
from copy import copy
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openai import OpenAI, APIConnectionError, APITimeoutError, RateLimitError, InternalServerError
from openpyxl import load_workbook
from PyPDF2 import PdfReader
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

st.set_page_config(page_title="RFQ Data Processor", layout="wide", initial_sidebar_state="auto")
st.markdown("""<style>
.stButton button {background:#ff914d;color:white;border-radius:8px;padding:10px 16px;margin-top:10px}
.stExpander {border-radius:10px}
</style>""", unsafe_allow_html=True)
if "show_diagnostics" not in st.session_state:
    st.session_state.show_diagnostics = False

# Configuration (never commit an API key to GitHub).
try:
    API_KEY = st.secrets.get("GROQ_API_KEY", os.environ.get("GROQ_API_KEY", ""))
    MODEL = st.secrets.get("GROQ_MODEL", os.environ.get("GROQ_MODEL", "openai/gpt-oss-20b"))
except Exception:
    API_KEY = os.environ.get("GROQ_API_KEY", "")
    MODEL = os.environ.get("GROQ_MODEL", "openai/gpt-oss-20b")
MIN_INTERVAL_SECONDS = 3.0
PROMPT_PATH = Path(__file__).with_name("prompt.txt")
try:
    FORMAT_PROMPT = PROMPT_PATH.read_text(encoding="utf-8").strip()
except OSError as exc:
    st.error(f"Unable to load prompt.txt: {exc}")
    st.stop()
if not FORMAT_PROMPT:
    st.error("prompt.txt is empty.")
    st.stop()
if not API_KEY.strip():
    st.error("Missing GROQ_API_KEY. Add it in Streamlit Cloud app settings > Secrets.")
    st.stop()

client = OpenAI(api_key=API_KEY.strip(), base_url="https://api.groq.com/openai/v1", timeout=90, max_retries=0)

@st.cache_resource
def rate_limiter():
    return {"lock": threading.Lock(), "last": 0.0}

LIMITER = rate_limiter()

def throttled_chat_complete(**kwargs):
    # Serializes calls across app sessions and spaces request start times.
    with LIMITER["lock"]:
        wait = MIN_INTERVAL_SECONDS - (time.monotonic() - LIMITER["last"])
        if wait > 0:
            time.sleep(wait)
        LIMITER["last"] = time.monotonic()
        return client.chat.completions.create(**kwargs)

def _clean(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (ValueError, TypeError):
        pass
    return str(value).strip()

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

def message_content(response):
    if not response.choices or response.choices[0].message.content is None:
        raise ValueError("Groq returned an empty response")
    return response.choices[0].message.content.strip()

COMBINED_OUTPUT_INSTRUCTIONS = """
Extract manufacturer name(s) explicitly stated in the source text, while formatting
all RFQ text according to the system instructions. Preserve technical details exactly.
Do not guess manufacturers, part numbers, sizes, quantities or standards.
Preferred reply: a single valid JSON object with keys "formatted_text" and
"manufacturer". The formatted_text value must be a string; the manufacturer
value must be a string (empty if no explicitly identified manufacturer).
If JSON is difficult, instead reply as:
<<<FORMATTED_TEXT>>>
[complete formatted RFQ text]
<<<MANUFACTURER>>>
[manufacturer names separated by ' - ', or leave blank]
No markdown fences, no commentary.
""".strip()


def parse_groq_response(raw):
    raw = raw.strip()
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I).strip()
    candidates = [cleaned]
    first, last = cleaned.find("{"), cleaned.rfind("}")
    if first >= 0 and last > first:
        candidates.append(cleaned[first:last + 1])
    for candidate in candidates:
        try:
            result = json.loads(candidate)
        except (json.JSONDecodeError, ValueError):
            continue
        if isinstance(result, dict):
            formatted = result.get("formatted_text", "")
            maker = result.get("manufacturer", "")
            if isinstance(maker, list):
                maker = " - ".join(str(x).strip() for x in maker if str(x).strip())
            if isinstance(formatted, str) and formatted.strip() and isinstance(maker, str):
                return formatted.strip(), maker.strip()
    token1, token2 = "<<<FORMATTED_TEXT>>>", "<<<MANUFACTURER>>>"
    if token1 in cleaned and token2 in cleaned:
        before, after = cleaned.split(token1, 1)[1].split(token2, 1)
        if before.strip():
            return before.strip(), after.strip().splitlines()[0].strip() if after.strip() else ""
    raise ValueError("Groq response could not be parsed into formatted_text and manufacturer. Use API Diagnostics to inspect model output.")


# Only retry network failures and rate limits. A 400 error should be shown directly.
@functools.lru_cache(maxsize=1024)
@retry(wait=wait_exponential(multiplier=2, min=2, max=20), stop=stop_after_attempt(3),
       retry=retry_if_exception_type((RateLimitError, APIConnectionError, APITimeoutError, InternalServerError)), reraise=True)
def process_rfq_text(text):
    cleaned = _clean(text)
    if not cleaned:
        return "", ""
    response = throttled_chat_complete(
        model=MODEL,
        messages=[
            {"role": "system", "content": FORMAT_PROMPT + "\n\n" + COMBINED_OUTPUT_INSTRUCTIONS},
            {"role": "user", "content": cleaned},
        ],
        temperature=0,
        # Do not set response_format=json_object: some Groq model outputs cause
        # server-side json_validate_failed (HTTP 400) before Python sees the text.
    )
    formatted, manufacturer = parse_groq_response(message_content(response))
    return formatted.replace("`", ""), manufacturer.replace("`", "")

# Compatibility with old callers.
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

def process_to_final(records, template, suffix, session_prefix, upload_wb=None):
    """Records: (line, desc, qty, uom, source_note) tuples."""
    wb = load_workbook(template)
    sheet = wb.active
    clear_template_rows(sheet)
    count = len(records)
    progress = st.progress(0)
    status = st.empty()
    for i, (line, desc, qty, uom, note) in enumerate(records, 1):
        status.write(f"Processing item {i} of {count}")
        formatted, manufacturer = process_rfq_text(note)
        target = i + 1
        for col, value in {"A":line, "B":desc, "C":qty, "D":uom, "E":formatted, "G":manufacturer}.items():
            sheet[f"{col}{target}"] = value
        progress.progress(i/count)
    # Customize the FINAL SHEET output (all processing modes):
    # E: PO Text -> Internal Note; drop H: mail id and F: Subject line.
    # Delete from right to left so G: Manufacturer moves to F correctly.
    sheet["E1"] = "Internal Note"
    sheet.delete_cols(8, 1)  # H = mail id
    sheet.delete_cols(6, 1)  # F = Subject line
    wrap_all(sheet)
    st.session_state[f"{session_prefix}_final_bytes"] = workbook_bytes(wb)
    if upload_wb is not None:
        st.session_state[f"{session_prefix}_upload_bytes"] = workbook_bytes(upload_wb)
    st.session_state[f"{session_prefix}_suffix"] = suffix
    status.empty()
    progress.empty()

def download_xlsx(label, data, filename, key):
    st.download_button(label, data, file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=key)

if st.button("API Diagnostics", key="toggle_diagnostics"):
    st.session_state.show_diagnostics = not st.session_state.show_diagnostics
if st.session_state.show_diagnostics:
    with st.container(border=True):
        st.subheader("Groq API Diagnostics")
        st.caption(f"Provider: Groq Cloud | Model: {MODEL} | Key configured: {'Yes' if API_KEY else 'No'}")
        st.caption(f"Request spacing: {MIN_INTERVAL_SECONDS} seconds; actual account rate limits also apply.")
        test_col1, test_col2 = st.columns(2)
        with test_col1:
            if st.button("1. Test standard chat API", key="diag_chat"):
                try:
                    reply = throttled_chat_complete(
                        model=MODEL,
                        messages=[{"role": "user", "content": "Reply with exactly: CHAT API OK"}],
                        temperature=0,
                    )
                    st.success("Chat API responded")
                    st.code(message_content(reply), language="text")
                except Exception as exc:
                    show_error("Chat API test failed", exc)
        with test_col2:
            if st.button("2. Test RFQ processing", key="diag_rfq"):
                try:
                    sample = "Basic Data Text:\nVALVE, SOLENOID\nPRESSURE RATING: 320 BAR\nMANUFACTURER: DANA CORPORATION"
                    formatted, manufacturer = process_rfq_text(sample)
                    st.success("Combined RFQ processing succeeded")
                    st.code(formatted, language="text")
                    st.write("Manufacturer:", manufacturer)
                except Exception as exc:
                    show_error("RFQ processing test failed", exc)

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
                              "F":rec["InternalNote"], "G":rec["Quantity"],
                              "H":rec["Unit of Measure"], "I":rec.get("Number", "")}
                    for c, value in values.items():
                        upload_sheet[f"{c}{row}"] = value
                    records.append((line, rec["Description"], rec["Quantity"], rec["Unit of Measure"], rec["InternalNote"]))
                wrap_all(upload_sheet)
                process_to_final(records, resolve_template(final_tpl, "FINAL SHEET.xlsx"), suffix.strip(), "excel", upload_wb)
                st.success("Excel processed")
            except Exception as exc:
                show_error("Excel processing failed (no incomplete files were saved)", exc)
    if "excel_upload_bytes" in st.session_state:
        s = st.session_state["excel_suffix"]
        download_xlsx("Download Upload file", st.session_state["excel_upload_bytes"], f"upload file - {s}.xlsx", "dl_excel_up")
        download_xlsx("Download FINAL SHEET", st.session_state["excel_final_bytes"], f"FINAL SHEET - {s}.xlsx", "dl_excel_fin")

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
                        sheet[f"{col}{i}"] = item[field]
                    records.append((item["RFx Item No"], item["Description"], item["QTY"], item["UOM"], item["PO Text"]))
                wrap_all(sheet)
                process_to_final(records, resolve_template(pdf_final_tpl, "FINAL SHEET.xlsx"), hts_number.strip(), "pdf", upload_wb)
                st.success("PDF processed")
            except Exception as exc:
                show_error("PDF processing failed (no incomplete files were saved)", exc)
    if "pdf_upload_bytes" in st.session_state:
        s = st.session_state["pdf_suffix"]
        download_xlsx("Download Upload file", st.session_state["pdf_upload_bytes"], f"upload file - {s}.xlsx", "dl_pdf_up")
        download_xlsx("Download FINAL SHEET", st.session_state["pdf_final_bytes"], f"FINAL SHEET - {s}.xlsx", "dl_pdf_fin")

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
                process_to_final(records, resolve_template(clean_final_tpl, "FINAL SHEET.xlsx"), "cleaned", "clean")
                st.success("HTS cleaned")
            except Exception as exc:
                show_error("HTS cleaning failed (no incomplete files were saved)", exc)
    if "clean_final_bytes" in st.session_state:
        download_xlsx("Download cleaned FINAL SHEET", st.session_state["clean_final_bytes"], "FINAL SHEET - cleaned.xlsx", "dl_clean")

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
