# streamlit_app.py
import os
import math
import time
import functools
import tempfile
from io import BytesIO
from collections import defaultdict

import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from mistralai import Mistral
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# ----------------------------
# 🔑  API & MODEL INITIALISATION
# ----------------------------
API_KEY = "MoYnS046nk9Z8WvGs2f057o27ZdP5TO9"
MODEL   = "mistral-large-latest"
client  = Mistral(api_key=API_KEY)

# ----------------------------
# 🛠️  HELPER FUNCTIONS
# ----------------------------

def as_clean_str(x: object) -> str:
    """Return a safe string ('' if None/NaN)."""
    if x is None:
        return ""
    if isinstance(x, float) and math.isnan(x):
        return ""
    return str(x)

# A very small in-memory cache for duplicate PO texts
@functools.lru_cache(maxsize=1024)
def _cached_format_text(po_text: str) -> str:
    return _format_text_uncached(po_text)

@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=1, min=2, max=20),
    retry=retry_if_exception_type(Exception),
)
def _format_text_uncached(po_text: str) -> str:
    """Call the Mistral agent once (with retries)."""
    chat_resp = client.agents.complete(
        agent_id="ag:9d0568a2:20250612:cleaner:12c5f2da",
        messages=[{"role": "user", "content": po_text}],
    )
    import re
    return re.sub(r"[`]+", "", chat_resp.choices[0].message.content)

def format_text(po_text: str) -> str:
    try:
        return _cached_format_text(as_clean_str(po_text))
    except Exception as e:
        st.error(f"Error formatting text → {e}")
        return as_clean_str(po_text)

@functools.lru_cache(maxsize=1024)
def _cached_manu(po_text: str) -> str:
    return _manu_uncached(po_text)

@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=1, min=2, max=20),
    retry=retry_if_exception_type(Exception),
)
def _manu_uncached(po_text: str) -> str:
    resp = client.chat.complete(
        model=MODEL,
        messages=[
            {
                "role": "user",
                "content": (
                    "Extract the manufacturer or maker names separated by hyphen - "
                    "mentioned in the PO text as a list in plain text. Output must "
                    "contain the list of manufacturer names only.\ncontent: "
                    + po_text
                ),
            }
        ],
    )
    return resp.choices[0].message.content

def manufacture_name(po_text: str) -> str:
    try:
        return _cached_manu(as_clean_str(po_text))
    except Exception as e:
        st.error(f"Error extracting manufacturer name → {e}")
        return ""

# ---------- PDF utilities ----------

def extract_text_from_pdf(pdf_bytes):
    reader = PdfReader(pdf_bytes)
    import re
    full = "".join(page.extract_text() for page in reader.pages)
    return re.sub(r"(REQUEST FOR QUOTATION[\s\S]*?RFQ Number \d+)", "", full)

def extract_rfq_from_pdf(pdf_bytes):
    reader = PdfReader(pdf_bytes)
    return "".join(page.extract_text() for page in reader.pages)

def parse_text(text: str, rfq_text: str):
    import re
    rfx_match = re.search(r"RFQ Number (\d+)", rfq_text)
    rfx_no    = rfx_match.group(1) if rfx_match else "Unknown"

    item_pat  = re.compile(r"(\d{5}) (\w?12\d{10}) (\d+(?:\.\d+)?)(\s*)(\w+) .*?(\d{2}\.\d{2}\.\d{4})", re.DOTALL)
    short_pat = re.compile(r"Short Text :(.*?)\n", re.DOTALL)
    po_pat    = re.compile(r"PO Material Text :(.*?)Agreement / LineNo.", re.DOTALL)

    items      = item_pat.findall(text)
    short_txts = short_pat.findall(text)
    po_txts    = po_pat.findall(text)

    data = []
    for i, itm in enumerate(items):
        mat_no = itm[1] if itm[1].startswith(("B12", "12", "B16", "15")) else ""
        data.append(
            {
                "RFx Number":  rfx_no,
                "RFx Item No": itm[0],
                "PR Item No":  "",
                "Material No": mat_no,
                "Description": short_txts[i] if i < len(short_txts) else "",
                "PO Text":     po_txts[i]   if i < len(po_txts)   else "",
                "QTY":         itm[2],
                "UOM":         itm[4],
            }
        )
    return data

# ---------- Excel helpers ----------

def workbook_to_bytes(wb):
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

cols_order = [
    "RFx Number", "RFx Item No", "PR Item No", "Material No",
    "Description", "PO Text", "QTY", "UOM",
]

def build_upload_wb(data: list[dict]):
    df = pd.DataFrame(data, columns=cols_order)
    buf = BytesIO(); df.to_excel(buf, index=False); buf.seek(0)
    wb = load_workbook(buf)
    ws = wb.active
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    return wb

def merge_into_template(template_path: str, upload_wb):
    df = pd.read_excel(BytesIO(workbook_to_bytes(upload_wb)))
    wb = load_workbook(template_path)
    ws = wb.active
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in r: c.value = None

    mapping = {"RFx Number": "A", "RFx Item No": "B", "PR Item No": "C", "Material No": "D",
               "Description": "E", "PO Text": "F", "QTY": "G", "UOM": "H"}
    for col_name, col_letter in mapping.items():
        for i, val in enumerate(df[col_name], start=2):
            ws[f"{col_letter}{i}"] = val

    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    return wb

def hts_to_final_sheet(upload_wb, final_template_path: str) -> bytes:
    final_wb = load_workbook(final_template_path)
    up_ws    = upload_wb.active
    fi_ws    = final_wb.active

    for r in fi_ws.iter_rows(min_row=2, max_row=fi_ws.max_row):
        for c in r: c.value = None

    paste = 2
    for r in up_ws.iter_rows(min_row=2, max_row=up_ws.max_row):
        if not any(cell.value for cell in r):
            continue
        fi_ws[f"A{paste}"] = r[1].value   # RFx Item No
        fi_ws[f"B{paste}"] = r[4].value   # Description
        fi_ws[f"C{paste}"] = r[6].value   # QTY
        fi_ws[f"D{paste}"] = r[7].value   # UOM
        po_text            = r[5].value or ""
        fi_ws[f"E{paste}"] = format_text(po_text)
        fi_ws[f"G{paste}"] = manufacture_name(po_text)
        paste += 1

    for row in fi_ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    return workbook_to_bytes(final_wb)

# ----------------------------
# 🎈  STREAMLIT PAGE CONFIG
# ----------------------------
st.set_page_config(page_title="Data Processor", layout="wide", initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
    .stButton button{background:#ff914d;color:#fff;border-radius:8px;padding:10px 16px;margin-top:10px;}
    .stExpander{background:#333;border-radius:10px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------
# ░█▀▀░█░█░█▀▀░█░█░█▀▄  COLUMN LAYOUT
# --------------------------------------
col1, col2, col3, col4 = st.columns([2, 2, 1.5, 1.5])

# ------------------------------------------------------
# 1️⃣  EXCEL DATA PROCESSOR
# ------------------------------------------------------
with col1:
    st.subheader("🗃️ Excel Data Processor")
    techno_file = st.file_uploader("Techno Commercial Envelope (.xls)", type=["xls"], key="techno_xls")
    with st.expander("Upload Excel Templates", expanded=True):
        upload_tpl = st.file_uploader("Upload File template (.xlsx)", type=["xlsx"], key="upl_tpl")
        final_tpl  = st.file_uploader("Final Sheet template (.xlsx)", type=["xlsx"], key="fin_tpl")
    upload_tpl  = upload_tpl or "upload file - HTS.xlsx"
    final_tpl   = final_tpl  or "FINAL SHEET.xlsx"

    if techno_file:
        custom_name = st.text_input("Custom name for results", key="cust_name_excel")
        if st.button("🚀 Process Excel", key="btn_excel") and custom_name:
            try:
                import re
                rfx_no   = re.search(r"\d+", techno_file.name).group()
                # keep_default_na=False to avoid NaNs
                xls      = pd.ExcelFile(techno_file)
                sheet_ok = next(
                    (
                        s
                        for s in xls.sheet_names
                        if all(
                            c in pd.read_excel(xls, sheet_name=s).columns
                            for c in ["Description", "InternalNote", "Quantity", "Unit of Measure"]
                        )
                    ),
                    None,
                )
                if not sheet_ok:
                    st.error("Template columns missing in uploaded XLS")
                    st.stop()
                df = pd.read_excel(techno_file, sheet_name=sheet_ok, keep_default_na=False)

                # Build Upload workbook
                wb_upl = load_workbook(upload_tpl)
                ws_upl = wb_upl.active
                for r in ws_upl.iter_rows(min_row=2, max_row=ws_upl.max_row):
                    for c in r: c.value = None
                row = 2; item = 10
                for _, rec in df.iterrows():
                    if rec["Description"]:
                        ws_upl[f"A{row}"] = rfx_no
                        ws_upl[f"B{row}"] = item
                        ws_upl[f"E{row}"] = rec["Description"]
                        ws_upl[f"H{row}"] = rec["Unit of Measure"]
                        ws_upl[f"G{row}"] = rec["Quantity"]
                        ws_upl[f"F{row}"] = rec["InternalNote"]
                        item += 10; row += 1
                for r in ws_upl.iter_rows():
                    for c in r: c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                st.download_button(
                    "📥 Download Upload File",
                    workbook_to_bytes(wb_upl),
                    file_name=f"upload file - {custom_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # Build FINAL SHEET
                wb_fin = load_workbook(final_tpl)
                ws_fin = wb_fin.active
                for r in ws_fin.iter_rows(min_row=2, max_row=ws_fin.max_row):
                    for c in r: c.value = None

                row = 2; item = 10
                for _, rec in df.iterrows():
                    if rec["Description"]:
                        ws_fin[f"A{row}"] = item
                        ws_fin[f"B{row}"] = rec["Description"]
                        ws_fin[f"C{row}"] = rec["Quantity"]
                        ws_fin[f"D{row}"] = rec["Unit of Measure"]
                        po = rec["InternalNote"]
                        ws_fin[f"E{row}"] = format_text(po)
                        ws_fin[f"G{row}"] = manufacture_name(po)
                        item += 10; row += 1
                for r in ws_fin.iter_rows():
                    for c in r: c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
                st.download_button(
                    "📥 Download FINAL SHEET",
                    workbook_to_bytes(wb_fin),
                    file_name=f"FINAL SHEET - {custom_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success("Excel processed ✔️")
            except Exception as e:
                st.error(f"❌ Error: {e}")

# --------------------------------------------------
# 2️⃣  PDF DATA PROCESSOR
# --------------------------------------------------
with col2:
    st.subheader("📑 PDF Data Processor")
    pdf_file = st.file_uploader("RFQ PDF", type=["pdf"], key="pdf_main")
    with st.expander("Upload Excel templates", expanded=True):
        raw_tpl   = st.file_uploader("Raw template (.xlsx)", type=["xlsx"], key="raw_tpl_pdf")
        hts_tpl   = st.file_uploader("HTS template (.xlsx)", type=["xlsx"], key="hts_tpl_pdf")
        fin_tpl_p = st.file_uploader("Final Sheet template (.xlsx)", type=["xlsx"], key="fin_tpl_pdf")
    raw_tpl   = raw_tpl   or "raw_template.xlsx"
    hts_tpl   = hts_tpl   or "upload file - HTS.xlsx"
    fin_tpl_p = fin_tpl_p or "FINAL SHEET.xlsx"

    if pdf_file:
        hts_no = st.text_input("HTS number", key="hts_num_pdf")
        if st.button("🚀 Process PDF", key="btn_pdf") and hts_no:
            try:
                # temp paths
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as t_upl, tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as t_fin:
                    t_upl_path, t_fin_path = t_upl.name, t_fin.name

                # ---------------- PDF ➜ Upload file ----------------
                rfq_text = extract_rfq_from_pdf(pdf_file)
                data     = parse_text(extract_text_from_pdf(pdf_file), rfq_text)
                wb_upload = build_upload_wb(data)
                wb_upload.save(t_upl_path)

                # ---------------- Upload ➜ Final sheet ----------------
                wb_final = merge_into_template(fin_tpl_p, wb_upload)
                wb_final.save(t_fin_path)

                upl_bytes = open(t_upl_path, "rb").read()
                fin_bytes = open(t_fin_path, "rb").read()

                st.download_button(
                    "📥 Download Upload File",
                    upl_bytes,
                    file_name=f"upload file - {hts_no}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.download_button(
                    "📥 Download FINAL SHEET",
                    fin_bytes,
                    file_name=f"FINAL SHEET - {hts_no}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success("PDF processed ✔️")
            except Exception as e:
                st.error(f"❌ Error: {e}")

# --------------------------------------------------
# 3️⃣  HTS CLEANER
# --------------------------------------------------
with col3:
    st.subheader("🧹 HTS Cleaner")
    hts_upload   = st.file_uploader("Upload *upload file – HTS.xlsx*", type=["xlsx"], key="hts_clean_upload")
    fin_tpl_opt  = st.file_uploader("Final Sheet template (.xlsx) – optional", type=["xlsx"], key="hts_clean_fin_tpl")
    fin_tpl_opt  = fin_tpl_opt or "FINAL SHEET.xlsx"

    if hts_upload:
        clean_name = st.text_input("Output file name suffix", value="Cleaned", key="clean_name")
        if st.button("🚀 Clean HTS", key="btn_clean"):
            try:
                wb_up  = load_workbook(hts_upload)
                final_bytes = hts_to_final_sheet(wb_up, fin_tpl_opt)
                st.download_button(
                    "📥 Download FINAL SHEET",
                    final_bytes,
                    file_name=f"FINAL SHEET - {clean_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success("HTS cleaned ✔️")
            except Exception as e:
                st.error(f"❌ Error: {e}")

# --------------------------------------------------
# 4️⃣  LIST MAKER (Manufacturer summary)
# --------------------------------------------------
# (You can continue your additional features here)
