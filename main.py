# -------------------------------------------------
#  📦  DATA PROCESSOR · unified & updated
# -------------------------------------------------
import streamlit as st
from st_copy_to_clipboard import st_copy_to_clipboard
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import re, os, time
from collections import defaultdict
from io import BytesIO
from mistralai import Mistral

# ──────────────────────────────────────────────
# 🔑  MISTRAL API CONFIG
# ──────────────────────────────────────────────
API_KEY = "YOUR_MISTRAL_API_KEY"          # ← put your real key here
MODEL   = "mistral-large-latest"

client = Mistral(api_key=API_KEY)

# ──────────────────────────────────────────────
# 🛠️  SMALL UTILITIES
# ──────────────────────────────────────────────
def workbook_to_bytes(wb) -> bytes:
    buf = BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.getvalue()

def _wrap(ws):
    for r in ws.iter_rows():
        for c in r:
            c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")

# ──────────────────────────────────────────────
# 🧽  TEXT-CLEAN / MANUFACTURER HELPERS
# ──────────────────────────────────────────────
def format_text(po_text: str) -> str:
    """
    Cleans & prettifies the PO text with a dedicated Mistral agent.
    Falls back to the raw text on any error.
    """
    try:
        resp = client.agents.complete(
            agent_id="ag:9d0568a2:20250612:cleaner:12c5f2da",   # <- your agent ID
            messages=[{"role": "user", "content": po_text}],
        )
        cleaned = re.sub(r"[`]+", "", resp.choices[0].message.content)
        return cleaned
    except Exception as e:
        st.error(f"Error formatting text ⇒ {e}")
        return po_text

def manufacture_name(po_text: str) -> str:
    """
    Extract manufacturer names separated by hyphen from PO text.
    """
    try:
        resp = client.chat.complete(
            model=MODEL,
            messages=[{
                "role": "user",
                "content": "Extract the manufacturer or maker names separated by "
                           "hyphen - mentioned in the PO text as a list in plain text. "
                           "Return only that list.\ncontent: " + po_text
            }],
        )
        return resp.choices[0].message.content
    except Exception as e:
        st.error(f"Error extracting manufacturer ⇒ {e}")
        return ""

# ──────────────────────────────────────────────
# 📑  PDF PARSING + EXCEL BUILD HELPERS
# ──────────────────────────────────────────────
def extract_text_from_pdf(pdf_bytes):
    reader = PdfReader(pdf_bytes)
    full = "".join(p.extract_text() for p in reader.pages)
    return re.sub(r"(REQUEST FOR QUOTATION[\s\S]*?RFQ Number \d+)", "", full)

def extract_rfq_from_pdf(pdf_bytes):
    reader = PdfReader(pdf_bytes)
    return "".join(p.extract_text() for p in reader.pages)

def parse_text(text: str, rfq_text: str):
    rfx_no = re.search(r"RFQ Number (\d+)", rfq_text)
    rfx_no = rfx_no.group(1) if rfx_no else "Unknown"

    item_pat  = re.compile(r'(\d{5}) (\w?12\d{10}) (\d+(?:\.\d+)?)(\s*)(\w+) .*?(\d{2}\.\d{2}\.\d{4})', re.DOTALL)
    short_pat = re.compile(r'Short Text :(.*?)\n', re.DOTALL)
    po_pat    = re.compile(r'PO Material Text :(.*?)Agreement / LineNo.', re.DOTALL)

    items      = item_pat.findall(text)
    short_txts = short_pat.findall(text)
    po_txts    = po_pat.findall(text)

    data = []
    for i, itm in enumerate(items):
        mat_no = itm[1] if itm[1].startswith(("B12", "12", "B16", "15")) else ""
        data.append({
            "RFx Number":  rfx_no,
            "RFx Item No": itm[0],
            "PR Item No":  "",
            "Material No": mat_no,
            "Description": short_txts[i] if i < len(short_txts) else "",
            "PO Text":     po_txts[i]   if i < len(po_txts)   else "",
            "QTY":         itm[2],
            "UOM":         itm[4],
        })
    return data

def insert_data_to_excel(data: list[dict], excel_path: str):
    order = ["RFx Number","RFx Item No","PR Item No","Material No",
             "Description","PO Text","QTY","UOM"]
    df = pd.DataFrame(data, columns=order)
    df.to_excel(excel_path, index=False)
    wb = load_workbook(excel_path); _wrap(wb.active); wb.save(excel_path)

def merge_into_template(template_path: str, created_path: str, out_path: str):
    df = pd.read_excel(created_path)
    wb = load_workbook(template_path); ws = wb.active

    # clear existing rows (keep header)
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in r: c.value = None

    map_cols = {"RFx Number":"A","RFx Item No":"B","PR Item No":"C","Material No":"D",
                "Description":"E","PO Text":"F","QTY":"G","UOM":"H"}

    for col, letter in map_cols.items():
        for i, val in enumerate(df[col], start=2):
            ws[f"{letter}{i}"] = val
    _wrap(ws); wb.save(out_path)

def process_pdf_to_final_excel(pdf_file, raw_tpl, hts_tpl, upload_path):
    rfq_text = extract_rfq_from_pdf(pdf_file)
    data     = parse_text(extract_text_from_pdf(pdf_file), rfq_text)
    insert_data_to_excel(data, raw_tpl)          # drop into raw template first
    merge_into_template(hts_tpl, raw_tpl, upload_path)

def process_final_sheet_from_pdf(final_tpl, upload_path, final_path):
    df = pd.read_excel(upload_path)
    wb = load_workbook(final_tpl); ws = wb.active

    # clear rows
    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for c in r: c.value = None

    row = 2
    for _, rec in df.iterrows():
        ws[f"A{row}"] = rec["RFx Item No"]
        ws[f"B{row}"] = rec["Description"]
        ws[f"C{row}"] = rec["QTY"]
        ws[f"D{row}"] = rec["UOM"]
        po_text       = rec["PO Text"] or ""
        ws[f"E{row}"] = format_text(po_text)
        ws[f"G{row}"] = manufacture_name(po_text)
        row += 1
    _wrap(ws); wb.save(final_path)

# ──────────────────────────────────────────────
# 🗒️  FINAL-SHEET → MANUFACTURER LIST
# ──────────────────────────────────────────────
def process_final_sheet_for_manufacturer(path: str) -> str:
    df = pd.read_excel(path)
    out = defaultdict(lambda: {"items": [], "emails": []})
    for _, r in df.iterrows():
        mans = r.get("Manufacturer")
        if pd.isna(mans): continue
        items = r["Line item number"]
        emails = [r[c] for c in df.columns if ("mail" in c.lower() or "unnamed" in c.lower()) and pd.notna(r[c])]
        for m in [m.strip() for m in str(mans).split("-") if m.strip()]:
            out[m]["items"].append(items)
            out[m]["emails"].extend(emails)

    chunks = []
    for man, det in out.items():
        items_str  = ", ".join(map(str, sorted(set(det["items"]))))
        emails_str = "\n".join(det["emails"])
        chunks.append(f"Item {items_str}: {man}\n{emails_str}\n")
    return "\n".join(chunks)

# ──────────────────────────────────────────────
# 🎈  STREAMLIT LAYOUT
# ──────────────────────────────────────────────
st.set_page_config(page_title="Data Processor", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.stButton button{background:#ff914d;color:#fff;border-radius:8px;padding:10px 16px;margin-top:10px;}
.stExpander{background:#333;border-radius:10px;}
.stTextInput>div>input{background:#2d2d2d;color:#ddd;border-radius:5px;border:none;padding:10px;}
</style>""", unsafe_allow_html=True)

col1, col2, col3 = st.columns([2, 2, 1])

# ──────────────────────────────────────────────
# 1️⃣  EXCEL DATA PROCESSOR
# ──────────────────────────────────────────────
with col1:
    st.subheader("🗃️ Excel Data Processor")
    techno_file = st.file_uploader("Techno-Commercial Envelope (.xls)", type=["xls"])
    with st.expander("Upload Excel Templates", expanded=True):
        upload_tpl = st.file_uploader("Upload File template (.xlsx)", type=["xlsx"])
        final_tpl  = st.file_uploader("Final Sheet template (.xlsx)", type=["xlsx"])

    # fallbacks
    upload_tpl = upload_tpl or "upload file - HTS.xlsx"
    final_tpl  = final_tpl  or "FINAL SHEET.xlsx"

    if techno_file:
        cust_name = st.text_input("Custom name for results")
        save_dir  = st.text_input("Save Path for outputs")
        if st.button("🚀 Process Excel"):
            if not (cust_name and save_dir):
                st.warning("Provide both a result name and a save path.")
            else:
                try:
                    rfx_no = re.search(r"\d+", techno_file.name).group()
                    xls = pd.ExcelFile(techno_file)
                    req_cols = {"Description","InternalNote","Quantity","Unit of Measure"}
                    sheet_ok = next((s for s in xls.sheet_names
                                     if req_cols.issubset(set(pd.read_excel(xls, sheet_name=s).columns))), None)
                    if not sheet_ok:
                        st.error("Template columns missing in uploaded XLS."); st.stop()

                    df = pd.read_excel(techno_file, sheet_name=sheet_ok)

                    # build UPLOAD file
                    wb = load_workbook(upload_tpl); ws = wb.active
                    for r in ws.iter_rows(min_row=2, max_row=ws.max_row):
                        for c in r: c.value = None
                    row = 2; item = 10
                    for _, rec in df.iterrows():
                        if pd.notna(rec["Description"]):
                            ws[f"A{row}"] = rfx_no
                            ws[f"B{row}"] = item
                            ws[f"E{row}"] = rec["Description"]
                            ws[f"H{row}"] = rec["Unit of Measure"]
                            ws[f"G{row}"] = rec["Quantity"]
                            ws[f"F{row}"] = rec["InternalNote"]
                            item += 10; row += 1
                    _wrap(ws)
                    out_upload = os.path.join(save_dir, f"upload file - {cust_name}.xlsx")
                    wb.save(out_upload)

                    # build FINAL SHEET
                    wb_fin = load_workbook(final_tpl); ws_fin = wb_fin.active
                    for r in ws_fin.iter_rows(min_row=2, max_row=ws_fin.max_row):
                        for c in r: c.value = None
                    row = 2; item = 10
                    for _, rec in df.iterrows():
                        if pd.notna(rec["Description"]):
                            ws_fin[f"A{row}"] = item
                            ws_fin[f"B{row}"] = rec["Description"]
                            ws_fin[f"C{row}"] = rec["Quantity"]
                            ws_fin[f"D{row}"] = rec["Unit of Measure"]
                            po = rec["InternalNote"] or ""
                            ws_fin[f"E{row}"] = format_text(po)
                            ws_fin[f"G{row}"] = manufacture_name(po)
                            item += 10; row += 1
                    _wrap(ws_fin)
                    out_final = os.path.join(save_dir, f"FINAL SHEET - {cust_name}.xlsx")
                    wb_fin.save(out_final)

                    st.success(f"Saved:\n• {out_upload}\n• {out_final}")
                except Exception as e:
                    st.error(f"❌ Error: {e}")

# ──────────────────────────────────────────────
# 2️⃣  UPDATED PDF DATA PROCESSOR  ← NEW LOGIC
# ──────────────────────────────────────────────
with col2:
    st.subheader("📑 PDF Data Processor")
    pdf_file = st.file_uploader("RFQ PDF", type=["pdf"])
    with st.expander("Upload Excel templates", expanded=True):
        raw_tpl   = st.file_uploader("Raw template (.xlsx)", type=["xlsx"])
        hts_tpl   = st.file_uploader("HTS template (.xlsx)", type=["xlsx"])
        fin_tpl   = st.file_uploader("Final Sheet template (.xlsx)", type=["xlsx"])

    # defaults
    raw_tpl = raw_tpl or "raw_template.xlsx"
    hts_tpl = hts_tpl or "upload file - HTS.xlsx"
    fin_tpl = fin_tpl or "FINAL SHEET.xlsx"

    if pdf_file:
        hts_no    = st.text_input("HTS Number")
        save_path = st.text_input("Save Path for outputs")

        if st.button("🚀 Process PDF"):
            if not (hts_no and save_path):
                st.warning("Fill both HTS number and save path.")
            else:
                try:
                    upload_path = os.path.join(save_path, f"upload file - {hts_no}.xlsx")
                    final_path  = os.path.join(save_path, f"FINAL SHEET - {hts_no}.xlsx")

                    process_pdf_to_final_excel(
                        pdf_file, raw_tpl, hts_tpl, upload_path
                    )
                    process_final_sheet_from_pdf(
                        fin_tpl, upload_path, final_path
                    )

                    st.success(f"Generated:\n• {upload_path}\n• {final_path}")
                except Exception as e:
                    st.error(f"❌ Error processing PDF ⇒ {e}")

# ──────────────────────────────────────────────
# 3️⃣  HTS CLEANER
# ──────────────────────────────────────────────
with col3:
    st.subheader("🧹 HTS Cleaner")
    hts_upload  = st.file_uploader("upload file – HTS.xlsx", type=["xlsx"])
    fin_tpl_alt = st.file_uploader("Final Sheet template (optional)", type=["xlsx"])

    fin_tpl_alt = fin_tpl_alt or "FINAL SHEET.xlsx"

    if hts_upload:
        suffix = st.text_input("Output name suffix", value="Cleaned")
        if st.button("🚀 Clean HTS"):
            try:
                wb = load_workbook(hts_upload)
                out_bytes = process_final_sheet_from_pdf(fin_tpl_alt, hts_upload,
                                                         f"temp_{time.time()}.xlsx")  # direct bytes not needed here
                st.download_button("Download FINAL SHEET",
                                   data=open(f"temp_{time.time()}.xlsx","rb").read(),
                                   file_name=f"FINAL SHEET - {suffix}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ Error: {e}")

# ──────────────────────────────────────────────
# 4️⃣  LIST MAKER – MANUFACTURER SUMMARY
# ──────────────────────────────────────────────
with col3:
    st.subheader("📝 List Maker")
    manu_file = st.file_uploader("Final Sheet for Manufacturer list", type=["xlsx"])

    if manu_file and st.button("🚀 Generate List"):
        try:
            out = process_final_sheet_for_manufacturer(manu_file)
            st.text_area("Formatted Output", out, height=300)
            st_copy_to_clipboard(out)
        except Exception as e:
            st.error(f"❌ Error: {e}")
