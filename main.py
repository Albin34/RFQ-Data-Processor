# main.py
import streamlit as st
from st_copy_to_clipboard import st_copy_to_clipboard
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment
import re
from collections import defaultdict
import os
import io
import time
from mistralai import Mistral

# ───────────────
# (1) Configuration / API‐client setup
# ───────────────
api_key = "RpBwVWJePMZCSS6cEDWROC4PTCNDl5sz"
model = "mistral-large-latest"
client = Mistral(api_key=api_key)

# ───────────────
# (2) Read system prompt (but don’t stop if missing)
# ───────────────
try:
    with open("prompt.txt", "r") as file:
        system_prompt = file.read().strip()
except FileNotFoundError:
    system_prompt = ""
    st.warning("Warning: prompt.txt not found; continuing with default prompt.")


def format_text(po_text: str) -> str:
    """
    Send po_text to the Mistral agent, strip backticks, return cleaned string.
    """
    try:
        chat_response = client.agents.complete(
            agent_id="ag:c04901dd:20241009:untitled-agent:4d5d10d7",
            messages=[{"role": "user", "content": po_text}],
        )
        time.sleep(3)
        return re.sub(r"[`']", "", chat_response.choices[0].message.content)
    except Exception as e:
        st.error(f"Error formatting text: {e}")
        return po_text  # fallback


def manufacture_name(po_text: str) -> str:
    """
    Ask Mistral to extract manufacturer names (hyphen-separated) from PO text.
    """
    try:
        chat_response = client.chat.complete(
            model=model,
            messages=[
                {
                    "role": "user",
                    "content": (
                        "Extract the manufacturer or maker names separated by hyphen '-' "
                        f"mentioned in the PO text as a list in plain text. Output should strictly "
                        f"contain the list of manufacturer names only.\ncontent: {po_text}"
                    ),
                }
            ],
        )
        time.sleep(3)
        return chat_response.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"Error extracting manufacturer name: {e}")
        return ""


def extract_text_from_pdf(pdf_file) -> str:
    """
    Read every page of a PDF (uploaded via Streamlit) and return concatenated text,
    after removing the “REQUEST FOR QUOTATION … RFQ Number ###” block.
    """
    reader = PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    # remove the repetitive “REQUEST FOR QUOTATION … RFQ Number ###” block
    cleaned_text = re.sub(r"(REQUEST FOR QUOTATION[\s\S]*?RFQ Number \d+)", "", text)
    return cleaned_text


def extract_rfq_from_pdf(pdf_file) -> str:
    """
    Read raw text of every page of a PDF (for the RFQ Number extraction).
    """
    reader = PdfReader(pdf_file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text


def parse_text(text: str, rfq_text: str) -> list[dict]:
    """
    Given cleaned text + raw rfq_text, extract line items. Returns a list of dicts with keys:
    ["RFx Number","RFx Item No","PR Item No","Material No","Description","PO Text","QTY","UOM"]
    """
    # 1) Find RFQ Number:
    rfx_number_match = re.search(r"RFQ Number (\d+)", rfq_text)
    rfx_number = rfx_number_match.group(1) if rfx_number_match else "Unknown"

    # 2) Regex for line items:
    item_pattern = re.compile(
        r"(\d{5}) (\w?12\d{10}) (\d+(?:\.\d+)?)(\s*)(\w+) .*?(\d{2}\.\d{2}\.\d{4})",
        re.DOTALL,
    )
    short_text_pattern = re.compile(r"Short Text :(.*?)\n", re.DOTALL)
    po_text_pattern = re.compile(r"PO Material Text :(.*?)Agreement / LineNo.", re.DOTALL)

    items = item_pattern.findall(text)
    short_texts = short_text_pattern.findall(text)
    po_texts = po_text_pattern.findall(text)

    data = []
    for i in range(len(items)):
        mat_candidate = items[i][1]
        if mat_candidate.startswith(("B12", "12", "B16", "15")):
            material_no = mat_candidate
        else:
            material_no = ""
        data.append(
            {
                "RFx Number": rfx_number,
                "RFx Item No": items[i][0],
                "PR Item No": "",
                "Material No": material_no,
                "Description": short_texts[i] if i < len(short_texts) else "",
                "PO Text": po_texts[i] if i < len(po_texts) else "",
                "QTY": items[i][2],
                "UOM": items[i][4],
            }
        )
    return data


def insert_data_to_workbook(data: list[dict], template_path: str) -> Workbook:
    """
    Given a list of dicts (from parse_text) and a template file path,
    return a new openpyxl.Workbook object whose cell A2:H… is populated.
    This does not save to disk; it returns the in‐memory Workbook.
    """
    # 1) Load the template:
    wb = load_workbook(template_path)
    ws = wb.active

    # 2) Clear row 2 onward:
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    # 3) Map columns:
    col_map = {
        "RFx Number": "A",
        "RFx Item No": "B",
        "PR Item No": "C",
        "Material No": "D",
        "Description": "E",
        "PO Text": "F",
        "QTY": "G",
        "UOM": "H",
    }

    # 4) Write each row of data into ws starting at row 2:
    for idx, rowdict in enumerate(data, start=2):
        for key, col_letter in col_map.items():
            ws[f"{col_letter}{idx}"] = rowdict[key]

    # 5) Wrap text for all cells:
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")

    return wb


def generate_upload_and_final_workbooks(pdf_file, raw_template_path, hts_number: str, hts_template_path) -> tuple[io.BytesIO, io.BytesIO]:
    """
    - Step 1: extract & clean text from PDF → parse_text → data list
    - Step 2: insert data into the “raw_template” to produce the “upload_file” workbook
    - Step 3: insert data + formatted PO Text + manufacturer into the “hts_template” to produce “final_sheet” workbook
    Returns: two BytesIO buffers: (upload_file_buffer, final_sheet_buffer)
    """
    # 1) Extract texts:
    cleaned_text = extract_text_from_pdf(pdf_file)
    rfq_text = extract_rfq_from_pdf(pdf_file)
    data = parse_text(cleaned_text, rfq_text)

    # 2) Build “upload_file – HTS.xlsx” in memory:
    #    Use raw_template_path (e.g. "data\\raw_template.xlsx") as the base.
    upload_wb = insert_data_to_workbook(data, raw_template_path)

    # 3) Now build “Final Sheet – HTS.xlsx”:
    #    We start from hts_template_path (e.g. "data\\upload file - HTS.xlsx") and then fill columns A–D from “upload_wb”
    #    plus format PO Text → column E and manufacturer → column G.
    final_wb = load_workbook(hts_template_path)
    final_ws = final_wb.active

    # 3a) Clear row 2 onward in final_ws:
    for row in final_ws.iter_rows(min_row=2, max_row=final_ws.max_row, min_col=1, max_col=final_ws.max_column):
        for cell in row:
            cell.value = None

    # 3b) Loop over each item in “data” (which matches upload_wb’s A:H columns), and write:
    paste_row = 2
    for rowdict in data:
        final_ws[f"A{paste_row}"] = rowdict["RFx Item No"]
        final_ws[f"B{paste_row}"] = rowdict["Description"]
        final_ws[f"C{paste_row}"] = rowdict["QTY"]
        final_ws[f"D{paste_row}"] = rowdict["UOM"]

        # Column E = formatted PO Text via Mistral
        formatted_po = format_text(rowdict["PO Text"])
        final_ws[f"E{paste_row}"] = formatted_po
        time.sleep(1)

        # Column G = manufacturer_name via Mistral
        mfr = manufacture_name(rowdict["PO Text"])
        final_ws[f"G{paste_row}"] = mfr
        time.sleep(1)

        paste_row += 1

    # 3c) Wrap text for entire final_ws:
    for row in final_ws.iter_rows(min_row=1, max_row=final_ws.max_row, min_col=1, max_col=final_ws.max_column):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")

    # 4) Save both workbooks to BytesIO buffers:
    upload_buffer = io.BytesIO()
    upload_wb.save(upload_buffer)
    upload_buffer.seek(0)

    final_buffer = io.BytesIO()
    final_wb.save(final_buffer)
    final_buffer.seek(0)

    return upload_buffer, final_buffer


def process_final_sheet_for_manufacturer(input_excel_path: str) -> str:
    """
    Read the “Final Sheet” into a pandas DataFrame, then produce the manufacturer summary text.
    """
    df = pd.read_excel(input_excel_path)
    output_dict = defaultdict(lambda: {"items": [], "emails": []})
    for _, row in df.iterrows():
        manufacturers = row.get("Manufacturer", "")
        if pd.notna(manufacturers):
            mfr_list = [m.strip() for m in manufacturers.split("-")]
            item_number = row.get("Line item number", "")
            emails = [
                row[col]
                for col in df.columns
                if "mail" in col.lower() or "unnamed" in col.lower()
            ]
            emails = [e for e in emails if pd.notna(e)]
            email_str = "\n".join(emails) if emails else None

            for mfr in mfr_list:
                output_dict[mfr]["items"].append(item_number)
                if email_str:
                    output_dict[mfr]["emails"].append(email_str)

    formatted_output_list = []
    for mfr, details in output_dict.items():
        items_str = ", ".join(map(str, sorted(set(details["items"]))))
        emails_combined = "\n".join(details["emails"])
        formatted_output_list.append(f"Item {items_str}: {mfr}\n{emails_combined}\n")

    return "\n".join(formatted_output_list)


# ───────────────
# (3) Build the Streamlit UI
# ───────────────
st.set_page_config(
    page_title="Data Processor", layout="wide", initial_sidebar_state="collapsed"
)

st.markdown(
    """
    <style>
    .stButton button {
        background-color: #ff914d;
        color: white;
        border-radius: 8px;
        padding: 10px 16px;
        margin-top: 10px;
    }
    .stTextInput>div>input {
        background-color: #2d2d2d;
        color: #ddd;
        border-radius: 5px;
        border: none;
        padding: 10px;
    }
    .stExpander {
        background-color: #333;
        border-radius: 10px;
    }
    .stHeader {
        color: #f1f1f1;
        font-size: 24px;
        font-weight: bold;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

col1, col2, col3 = st.columns([2, 2, 1])

# ---- Column 1: Excel Data Processor (unchanged) ----
with col1:
    st.subheader("🗃️ Excel Data Processor")
    techno_commercial_file = st.file_uploader(
        "Upload Techno Commercial Envelope File (.xls)", type=["xls"]
    )
    with st.expander("Upload Excel Files", expanded=True):
        upload_file = st.file_uploader("Upload File (.xlsx)", type=["xlsx"])
        final_sheet_file = st.file_uploader("Final Sheet File (.xlsx)", type=["xlsx"])

    if not upload_file:
        upload_file = os.path.join(os.getcwd(), "upload file - HTS.xlsx")
    if not final_sheet_file:
        final_sheet_file = os.path.join(os.getcwd(), "FINAL SHEET.xlsx")

    if techno_commercial_file:
        custom_name_excel = st.text_input("Custom Name for 'Upload HTS'")
        # We no longer ask for “Save Path” here; we will offer a download.
        if st.button("🚀 Process Excel Files"):
            if custom_name_excel:
                try:
                    # Exactly the same logic as before, except we save into memory and then download.
                    rfx_number = re.search(r"\d+", techno_commercial_file.name).group()
                    xls = pd.ExcelFile(techno_commercial_file)
                    required_columns = ["Description", "InternalNote", "Quantity", "Unit of Measure"]
                    correct_sheet_name = None
                    for sheet_name in xls.sheet_names:
                        df_ = pd.read_excel(xls, sheet_name=sheet_name)
                        if all(column in df_.columns for column in required_columns):
                            correct_sheet_name = sheet_name
                            break
                    if correct_sheet_name is None:
                        st.error("Could not find a sheet with the required columns.")
                        raise ValueError("Missing required columns")
                    techno_df = pd.read_excel(techno_commercial_file, sheet_name=correct_sheet_name)

                    # Load the “upload_file” template from disk:
                    wb_upload = load_workbook(upload_file)
                    ws_upload = wb_upload.active
                    # Clear row 2 onward
                    for row in ws_upload.iter_rows(min_row=2, max_row=ws_upload.max_row):
                        for cell in row:
                            cell.value = None

                    paste_row = 2
                    rfx_item_no = 10
                    for i in range(len(techno_df)):
                        if pd.notna(techno_df["Description"].iloc[i]) and i != 1:
                            ws_upload[f"A{paste_row}"] = rfx_number
                            ws_upload[f"B{paste_row}"] = rfx_item_no
                            ws_upload[f"E{paste_row}"] = techno_df["Description"].iloc[i]
                            ws_upload[f"H{paste_row}"] = techno_df["Unit of Measure"].iloc[i]
                            ws_upload[f"G{paste_row}"] = techno_df["Quantity"].iloc[i]
                            ws_upload[f"F{paste_row}"] = techno_df["InternalNote"].iloc[i]
                            ws_upload[f"I{paste_row}"] = techno_df["Number"].iloc[i]
                            paste_row += 1
                            rfx_item_no += 10

                    # Wrap text for the upload workbook:
                    for row in ws_upload.iter_rows():
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")

                    # Now create the “final worksheet” from the final_sheet_file template:
                    wb_final = load_workbook(final_sheet_file)
                    ws_final = wb_final.active
                    for row in ws_final.iter_rows(min_row=2, max_row=ws_final.max_row):
                        for cell in row:
                            cell.value = None

                    paste_row1 = 2
                    rfx_item_no1 = 10
                    for i in range(len(techno_df)):
                        if pd.notna(techno_df["Description"].iloc[i]) and i != 1:
                            ws_final[f"A{paste_row1}"] = rfx_item_no1
                            ws_final[f"B{paste_row1}"] = techno_df["Description"].iloc[i]
                            ws_final[f"C{paste_row1}"] = techno_df["Quantity"].iloc[i]
                            ws_final[f"D{paste_row1}"] = techno_df["Unit of Measure"].iloc[i]
                            po_text_ = techno_df["InternalNote"].iloc[i]
                            formatted_po = format_text(po_text_)
                            ws_final[f"E{paste_row1}"] = formatted_po
                            time.sleep(1)
                            mfr_ = manufacture_name(po_text_)
                            ws_final[f"G{paste_row1}"] = mfr_
                            time.sleep(1)
                            paste_row1 += 1
                            rfx_item_no1 += 10

                    for row in ws_final.iter_rows():
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")

                    # Now save both workbooks in memory:
                    buf_upload = io.BytesIO()
                    wb_upload.save(buf_upload)
                    buf_upload.seek(0)

                    buf_final = io.BytesIO()
                    wb_final.save(buf_final)
                    buf_final.seek(0)

                    # Offer download buttons in Streamlit:
                    download_name1 = f"upload file - {custom_name_excel}.xlsx"
                    st.download_button(
                        label="⬇️ Download Upload‐File",
                        data=buf_upload,
                        file_name=download_name1,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                    download_name2 = f"FINAL SHEET - {custom_name_excel}.xlsx"
                    st.download_button(
                        label="⬇️ Download Final‐Sheet",
                        data=buf_final,
                        file_name=download_name2,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                except Exception as e:
                    st.error(f"An error occurred: {e}")
            else:
                st.warning("Please provide a custom name.")


# ---- Column 2: PDF Data Processor (replaced Save Path with download) ----
with col2:
    st.subheader("📑 PDF Data Processor")
    pdf_file = st.file_uploader("Upload PDF File", type=["pdf"])
    with st.expander("Upload Excel Files", expanded=True):
        created_excel_template = st.file_uploader("Raw Template File (.xlsx)", type=["xlsx"])
        template_excel_path = st.file_uploader("HTS Template File (.xlsx)", type=["xlsx"])
        pdf_final_sheet = st.file_uploader("Final Sheet Template (.xlsx)", type=["xlsx"])

    if not created_excel_template:
        created_excel_template = os.path.join(os.getcwd(), "raw_template.xlsx")
    if not template_excel_path:
        template_excel_path = os.path.join(os.getcwd(), "upload file - HTS.xlsx")
    if not pdf_final_sheet:
        pdf_final_sheet = os.path.join(os.getcwd(), "FINAL SHEET.xlsx")

    htsnum = st.text_input("HTS Number")
    if pdf_file and htsnum:
        if st.button("🚀 Process PDF Files"):
            try:
                # Generate two in‐memory Excel files:
                buf_upload_pdf, buf_final_pdf = generate_upload_and_final_workbooks(
                    pdf_file, 
                    created_excel_template, 
                    htsnum, 
                    template_excel_path
                )

                # Offer Download buttons (names include HTSNum):
                download_name_pdf1 = f"upload file - {htsnum}.xlsx"
                st.download_button(
                    label="⬇️ Download Upload‐File (from PDF)",
                    data=buf_upload_pdf,
                    file_name=download_name_pdf1,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                download_name_pdf2 = f"FINAL SHEET - {htsnum}.xlsx"
                st.download_button(
                    label="⬇️ Download Final‐Sheet (from PDF)",
                    data=buf_final_pdf,
                    file_name=download_name_pdf2,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.success("Your files are ready for download!")
            except Exception as e:
                st.error(f"An error occurred while processing the PDF: {e}")
    else:
        if not pdf_file:
            st.info("Please upload a PDF file.")
        elif not htsnum:
            st.info("Please enter an HTS Number.")


# ---- Column 3: List Maker (unchanged) ----
with col3:
    st.subheader("📝 List Maker")
    final_sheet_for_manufacturer = st.file_uploader("Final Sheet File for Manufacturer", type=["xlsx"])

    if final_sheet_for_manufacturer:
        if st.button("🚀 Process Uploaded File"):
            try:
                final_output = process_final_sheet_for_manufacturer(final_sheet_for_manufacturer)
                st.text_area("Formatted Output", final_output, height=300)
                st_copy_to_clipboard(final_output)
            except Exception as e:
                st.error(f"An error occurred: {e}")

    if st.button("🚀 Process Default File"):
        try:
            default_path = os.path.join(os.getcwd(), "FINAL SHEET.xlsx")
            final_output = process_final_sheet_for_manufacturer(default_path)
            st.text_area("Formatted Output", final_output, height=300)
            st_copy_to_clipboard(final_output)
        except Exception as e:
            st.error(f"An error occurred: {e}")
