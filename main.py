import streamlit as st
from st_copy_to_clipboard import st_copy_to_clipboard
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import re
from collections import defaultdict
import os
import openai
import time
import io
from mistralai import Mistral

api_key = "RpBwVWJePMZCSS6cEDWROC4PTCNDl5sz"
model = "mistral-large-latest"

client = Mistral(api_key=api_key)

# Read the system prompt from a text file
try:
    with open("prompt.txt", "r") as file:
        system_prompt = file.read().strip()
except FileNotFoundError:
    st.error("System prompt file not found.")
    st.stop()

def format_text(po_text):
    try:
        chat_response = client.agents.complete(
        agent_id="ag:c04901dd:20241009:untitled-agent:4d5d10d7",
        messages = [
            {
                "role": "user",
                "content": po_text,
            },
        ]
        )
        time.sleep(3)
        chars_to_exclude = "```"  # Added " and '
        for_potext = re.sub(f"[{re.escape(chars_to_exclude)}]", "", chat_response.choices[0].message.content)
        return for_potext
        
    except Exception as e:
        st.error(f"Error formatting text: {e}")
        return po_text

def manufacture_name(po_text):
    try:
        chat_response = client.chat.complete(
        model = model,
        messages = [
            {
                "role": "user",
                "content": f"Extract the manufacturer or maker names seperated by hyphen -  mentioned in the po text as a list in plain text. Output should strictly contain the list of manufacturer names only  \n content: {po_text}",
            },
        ]
        )
        time.sleep(3)
        return chat_response.choices[0].message.content
    except Exception as e:
        st.error(f"Error extracting manufacturer name: {e}")
        return ""

def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ''
    for page in reader.pages:
        text += page.extract_text()
    
    # Remove the repetitive "REQUEST FOR QUOTATION" block
    cleaned_text = re.sub(r'(REQUEST FOR QUOTATION[\s\S]*?RFQ Number \d+)', '', text)
    return cleaned_text

def extract_rfq_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ''
    for page in reader.pages:
        text += page.extract_text()
    return text

def parse_text(text, rfq_text):
    rfx_number_match = re.search(r'RFQ Number (\d+)', rfq_text)
    rfx_number = rfx_number_match.group(1) if rfx_number_match else "Unknown"
    
    item_pattern = re.compile(r'(\d{5}) (\w?12\d{10}) (\d+(?:\.\d+)?)(\s*)(\w+) .*?(\d{2}\.\d{2}\.\d{4})', re.DOTALL)
    short_text_pattern = re.compile(r'Short Text :(.*?)\n', re.DOTALL)
    po_text_pattern = re.compile(r'PO Material Text :(.*?)Agreement / LineNo.', re.DOTALL)
    agreement_pattern = re.compile(r'Agreement / LineNo. Plant Description / Storage Location Description(.*?)(?=000\d{2}|$)', re.DOTALL)
    
    items = item_pattern.findall(text)
    short_texts = short_text_pattern.findall(text)
    po_texts = po_text_pattern.findall(text)
    agreements = agreement_pattern.findall(text)

    data = []
    for i in range(len(items)):
        material_no = items[i][1] if items[i][1].startswith(('B12', '12', 'B16', '15')) else ''
        data.append({
            "RFx Number": rfx_number,
            "RFx Item No": items[i][0],
            "PR Item No": "",
            "Material No": material_no,
            "Description": short_texts[i] if i < len(short_texts) else "",
            "PO Text": po_texts[i] if i < len(po_texts) else "",
            "QTY": items[i][2],
            "UOM": items[i][4],
        })
    return data

def create_excel_buffer(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

def insert_data_to_new_excel(data):
    columns_order = ["RFx Number", "RFx Item No", "PR Item No", "Material No", "Description", "PO Text", "QTY", "UOM"]
    df = pd.DataFrame(data, columns=columns_order)
    if not df.empty:
        wb = load_workbook(io.BytesIO())
        sheet = wb.active
        for row_idx, row in enumerate(data, start=1):
            for col_idx, col_name in enumerate(columns_order, start=1):
                sheet.cell(row=row_idx, column=col_idx, value=row[col_name])
        
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output
    return None

def merge_into_template(template_content, created_content):
    created_df = pd.read_excel(created_content)
    wb = load_workbook(template_content)
    ws = wb.active
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None
    
    column_mapping = {
        "RFx Number": "A",
        "RFx Item No": "B",
        "PR Item No": "C",
        "Material No": "D",
        "Description": "E",
        "PO Text": "F",
        "QTY": "G",
        "UOM": "H"
    }
    
    for col_name, col_letter in column_mapping.items():
        for row_index, value in enumerate(created_df[col_name], start=2):
            ws[f"{col_letter}{row_index}"] = value
    
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def process_pdf_to_final_excel(pdf_file):
    clean_text = extract_text_from_pdf(pdf_file)
    rfq_text = extract_rfq_from_pdf(pdf_file)
    data = parse_text(clean_text, rfq_text)
    
    # Create raw template
    raw_template = insert_data_to_new_excel(data)
    
    # Process with HTS template
    hts_template_path = "upload file - HTS.xlsx"
    with open(hts_template_path, "rb") as f:
        hts_template_content = io.BytesIO(f.read())
    
    final_output = merge_into_template(hts_template_content, raw_template)
    return final_output

def process_final_sheet_from_pdf(pdf_file, upload_file_content):
    created_df = pd.read_excel(upload_file_content)
    
    # Create a new workbook for final sheet
    wb = load_workbook("FINAL SHEET.xlsx")
    final_sheet = wb.active
    
    # Clear existing data
    for row in final_sheet.iter_rows(min_row=2, max_row=final_sheet.max_row):
        for cell in row:
            cell.value = None
    
    # Populate data
    paste_row = 2
    for index, row in created_df.iterrows():
        final_sheet[f'A{paste_row}'] = row['RFx Item No']
        final_sheet[f'B{paste_row}'] = row['Description']
        final_sheet[f'C{paste_row}'] = row['QTY']
        final_sheet[f'D{paste_row}'] = row['UOM']
        po_text = row['PO Text']
        formatted_po_text = format_text(po_text)
        final_sheet[f'E{paste_row}'] = formatted_po_text
        manufacturer_name = manufacture_name(po_text)
        final_sheet[f'G{paste_row}'] = manufacturer_name
        paste_row += 1
    
    for row in final_sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def process_final_sheet_for_manufacturer(input_excel):
    df = pd.read_excel(input_excel)
    output_dict = defaultdict(lambda: {"items": [], "emails": []})

    for index, row in df.iterrows():
        manufacturers = row['Manufacturer']
        if pd.notna(manufacturers):
            manufacturer_list = [m.strip() for m in manufacturers.split('-')]
            item_number = row['Line item number']
            emails = [row[col] for col in df.columns if "mail" in col or "Unnamed" in col]
            filtered_emails = [email for email in emails if pd.notna(email)]
            email_str = "\n".join(filtered_emails) if filtered_emails else None

            for manufacturer in manufacturer_list:
                output_dict[manufacturer]["items"].append(item_number)
                if email_str:
                    output_dict[manufacturer]["emails"].append(email_str)

    formatted_output = []
    for manufacturer, details in output_dict.items():
        items_str = ", ".join(map(str, sorted(set(details["items"]))))
        emails_str = "\n".join(details["emails"])
        formatted_output.append(f"Item {items_str}: {manufacturer}\n{emails_str}\n")

    return "\n".join(formatted_output)

# Setting up the layout for the page
st.set_page_config(page_title="Data Processor", layout="wide", initial_sidebar_state="collapsed")

# Custom styling
st.markdown("""
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
    """, unsafe_allow_html=True)

# App Layout
col1, col2, col3 = st.columns([2, 2, 1])

# ---- Column 1: Excel Data Processor ----
with col1:
    st.subheader("🗃️ Excel Data Processor")
    techno_commercial_file = st.file_uploader("Upload Techno Commercial Envelope File (.xls)", type=['xls'])
    with st.expander("Upload Excel Files", expanded=True):
        upload_file = st.file_uploader("Upload File (.xlsx)", type=['xlsx'])
        final_sheet_file = st.file_uploader("Final Sheet File (.xlsx)", type=['xlsx'])

    if techno_commercial_file:
        custom_name_excel = st.text_input("Custom Name for Output Files")
        
        if st.button("🚀 Process Excel Files"):
            if custom_name_excel:
                try:
                    rfx_number = re.search(r'\d+', techno_commercial_file.name).group()
                    xls = pd.ExcelFile(techno_commercial_file)
                    required_columns = ['Description', 'InternalNote', 'Quantity', 'Unit of Measure']
                    correct_sheet_name = None
                    
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        if all(column in df.columns for column in required_columns):
                            correct_sheet_name = sheet_name
                            break
                    
                    if correct_sheet_name is None:
                        st.error("Could not find a sheet with the required columns.")
                        raise ValueError("Could not find a sheet with the required columns.")
                    
                    techno_df = pd.read_excel(techno_commercial_file, sheet_name=correct_sheet_name)
                    
                    # Create upload file
                    upload_wb = load_workbook("upload file - HTS.xlsx")
                    upload_sheet = upload_wb.active
                    
                    for row in upload_sheet.iter_rows(min_row=2, max_row=upload_sheet.max_row):
                        for cell in row:
                            cell.value = None
                    
                    paste_row = 2
                    rfx_item_no = 10
                    for i in range(len(techno_df)):
                        if pd.notna(techno_df['Description'].iloc[i]) and i != 1:
                            upload_sheet[f'A{paste_row}'] = rfx_number
                            upload_sheet[f'B{paste_row}'] = rfx_item_no
                            upload_sheet[f'E{paste_row}'] = techno_df['Description'].iloc[i]
                            upload_sheet[f'H{paste_row}'] = techno_df['Unit of Measure'].iloc[i]
                            upload_sheet[f'G{paste_row}'] = techno_df['Quantity'].iloc[i]
                            upload_sheet[f'F{paste_row}'] = techno_df['InternalNote'].iloc[i]
                            upload_sheet[f'I{paste_row}'] = techno_df['Number'].iloc[i]
                            paste_row += 1
                            rfx_item_no += 10
                    
                    for row in upload_sheet.iter_rows():
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                    
                    upload_output = io.BytesIO()
                    upload_wb.save(upload_output)
                    upload_output.seek(0)
                    
                    # Create final sheet
                    final_wb = load_workbook("FINAL SHEET.xlsx")
                    final_sheet = final_wb.active
                    
                    for row in final_sheet.iter_rows(min_row=2, max_row=final_sheet.max_row):
                        for cell in row:
                            cell.value = None
                    
                    paste_row1 = 2
                    rfx_item_no1 = 10
                    for i in range(len(techno_df)):
                        if pd.notna(techno_df['Description'].iloc[i]) and i != 1:
                            final_sheet[f'A{paste_row1}'] = rfx_item_no1
                            final_sheet[f'B{paste_row1}'] = techno_df['Description'].iloc[i]
                            final_sheet[f'C{paste_row1}'] = techno_df['Quantity'].iloc[i]
                            final_sheet[f'D{paste_row1}'] = techno_df['Unit of Measure'].iloc[i]
                            po_text = techno_df['InternalNote'].iloc[i]
                            formatted_po_text = format_text(po_text)
                            final_sheet[f'E{paste_row1}'] = formatted_po_text
                            time.sleep(1)
                            manufacturer_name = manufacture_name(po_text)
                            final_sheet[f'G{paste_row1}'] = manufacturer_name
                            time.sleep(1)
                            paste_row1 += 1
                            rfx_item_no1 += 10
                    
                    for row in final_sheet.iter_rows():
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                    
                    final_output = io.BytesIO()
                    final_wb.save(final_output)
                    final_output.seek(0)
                    
                    # Offer files for download
                    st.download_button(
                        label="📥 Download Upload File",
                        data=upload_output,
                        file_name=f"upload file - {custom_name_excel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.download_button(
                        label="📥 Download Final Sheet",
                        data=final_output,
                        file_name=f"FINAL SHEET - {custom_name_excel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("Files processed successfully! Click the buttons above to download.")
                except Exception as e:
                    st.error(f"An error occurred: {e}")
            else:
                st.warning("Please provide a custom name for the output files.")

# ---- Column 2: PDF Data Processor ----
with col2:
    st.subheader("📑 PDF Data Processor")
    pdf_file = st.file_uploader("Upload PDF File", type=['pdf'])
    with st.expander("Upload Excel Files", expanded=True):
        created_excel_template = st.file_uploader("Raw Template File (.xlsx)", type=['xlsx'])
        template_excel_path = st.file_uploader("HTS Template File (.xlsx)", type=['xlsx'])
        pdf_final_sheet = st.file_uploader("Final Sheet Template (.xlsx)", type=['xlsx'])

    if pdf_file:
        htsnum = st.text_input("HTS Number for Output Files")
        
        if st.button("🚀 Process PDF Files"):
            if htsnum:
                try:
                    # Process PDF to create upload file
                    upload_file_content = process_pdf_to_final_excel(pdf_file)
                    
                    # Process to create final sheet
                    final_sheet_content = process_final_sheet_from_pdf(pdf_file, upload_file_content)
                    
                    # Offer files for download
                    st.download_button(
                        label="📥 Download Upload File",
                        data=upload_file_content,
                        file_name=f"upload file - {htsnum}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.download_button(
                        label="📥 Download Final Sheet",
                        data=final_sheet_content,
                        file_name=f"FINAL SHEET - {htsnum}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("PDF processed successfully! Click the buttons above to download.")
                except Exception as e:
                    st.error(f"An error occurred while processing the PDF: {e}")
            else:
                st.warning("Please provide an HTS number for the output files.")

# ---- Column 3: List Maker ----
with col3:
    st.subheader("📝 List Maker")
    final_sheet_for_manufacturer = st.file_uploader("Final Sheet File for Manufacturer", type=['xlsx'])

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
            with open("FINAL SHEET.xlsx", "rb") as f:
                final_output = process_final_sheet_for_manufacturer(io.BytesIO(f.read()))
            st.text_area("Formatted Output", final_output, height=300)
            st_copy_to_clipboard(final_output)
        except Exception as e:
            st.error(f"An error occurred: {e}")
