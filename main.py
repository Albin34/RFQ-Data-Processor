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
        chars_to_exclude = "```"
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

def create_excel_in_memory(data):
    columns_order = ["RFx Number", "RFx Item No", "PR Item No", "Material No", "Description", "PO Text", "QTY", "UOM"]
    df = pd.DataFrame(data, columns=columns_order)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        workbook = writer.book
        sheet = workbook.active
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
    
    output.seek(0)
    return output

def process_final_sheet_in_memory(input_df):
    output = io.BytesIO()
    workbook = load_workbook(io.BytesIO(input_df))
    sheet = workbook.active
    
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        for cell in row:
            cell.value = None
    
    paste_row = 2
    for index, row in input_df.iterrows():
        sheet[f'A{paste_row}'] = row['RFx Item No']
        sheet[f'B{paste_row}'] = row['Description']
        sheet[f'C{paste_row}'] = row['QTY']
        sheet[f'D{paste_row}'] = row['UOM']
        po_text = row['PO Text']
        formatted_po_text = format_text(po_text)
        sheet[f'E{paste_row}'] = formatted_po_text
        manufacturer_name = manufacture_name(po_text)
        sheet[f'G{paste_row}'] = manufacturer_name
        paste_row += 1
    
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
    
    workbook.save(output)
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

# Streamlit UI
st.set_page_config(page_title="Data Processor", layout="wide", initial_sidebar_state="collapsed")

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

col1, col2, col3 = st.columns([2, 2, 1])

# Column 1: Excel Data Processor
with col1:
    st.subheader("🗃️ Excel Data Processor")
    techno_commercial_file = st.file_uploader("Upload Techno Commercial Envelope File (.xls)", type=['xls'])
    with st.expander("Upload Excel Files", expanded=True):
        upload_template = st.file_uploader("Upload Template File (.xlsx)", type=['xlsx'])
        final_sheet_template = st.file_uploader("Final Sheet Template (.xlsx)", type=['xlsx'])

    if techno_commercial_file and upload_template and final_sheet_template:
        custom_name_excel = st.text_input("Custom Name for Output Files")

        if st.button("🚀 Process Excel Files"):
            if custom_name_excel:
                try:
                    # Process the techno commercial file
                    rfx_number = re.search(r'\d+', techno_commercial_file.name).group()
                    xls = pd.ExcelFile(techno_commercial_file)
                    
                    # Find the correct sheet
                    required_columns = ['Description', 'InternalNote', 'Quantity', 'Unit of Measure']
                    correct_sheet_name = None
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        if all(column in df.columns for column in required_columns):
                            correct_sheet_name = sheet_name
                            break
                    
                    if correct_sheet_name is None:
                        st.error("Could not find a sheet with the required columns.")
                        raise ValueError("Required columns not found.")
                    
                    techno_df = pd.read_excel(techno_commercial_file, sheet_name=correct_sheet_name)
                    
                    # Create upload file in memory
                    upload_wb = load_workbook(upload_template)
                    upload_ws = upload_wb.active
                    
                    for row in upload_ws.iter_rows(min_row=2, max_row=upload_ws.max_row):
                        for cell in row:
                            cell.value = None
                    
                    paste_row = 2
                    rfx_item_no = 10
                    for i in range(len(techno_df)):
                        if pd.notna(techno_df['Description'].iloc[i]) and i != 1:
                            upload_ws[f'A{paste_row}'] = rfx_number
                            upload_ws[f'B{paste_row}'] = rfx_item_no
                            upload_ws[f'E{paste_row}'] = techno_df['Description'].iloc[i]
                            upload_ws[f'H{paste_row}'] = techno_df['Unit of Measure'].iloc[i]
                            upload_ws[f'G{paste_row}'] = techno_df['Quantity'].iloc[i]
                            upload_ws[f'F{paste_row}'] = techno_df['InternalNote'].iloc[i]
                            upload_ws[f'I{paste_row}'] = techno_df['Number'].iloc[i]
                            paste_row += 1
                            rfx_item_no += 10
                    
                    for row in upload_ws.iter_rows():
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                    
                    upload_buffer = io.BytesIO()
                    upload_wb.save(upload_buffer)
                    upload_buffer.seek(0)
                    
                    # Create final sheet in memory
                    final_wb = load_workbook(final_sheet_template)
                    final_ws = final_wb.active
                    
                    for row in final_ws.iter_rows(min_row=2, max_row=final_ws.max_row):
                        for cell in row:
                            cell.value = None
                    
                    paste_row1 = 2
                    rfx_item_no1 = 10
                    for i in range(len(techno_df)):
                        if pd.notna(techno_df['Description'].iloc[i]) and i != 1:
                            final_ws[f'A{paste_row1}'] = rfx_item_no1
                            final_ws[f'B{paste_row1}'] = techno_df['Description'].iloc[i]
                            final_ws[f'C{paste_row1}'] = techno_df['Quantity'].iloc[i]
                            final_ws[f'D{paste_row1}'] = techno_df['Unit of Measure'].iloc[i]
                            po_text = techno_df['InternalNote'].iloc[i]
                            formatted_po_text = format_text(po_text)
                            final_ws[f'E{paste_row1}'] = formatted_po_text
                            manufacturer_name = manufacture_name(po_text)
                            final_ws[f'G{paste_row1}'] = manufacturer_name
                            paste_row1 += 1
                            rfx_item_no1 += 10
                    
                    for row in final_ws.iter_rows():
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                    
                    final_buffer = io.BytesIO()
                    final_wb.save(final_buffer)
                    final_buffer.seek(0)
                    
                    # Create download buttons
                    st.download_button(
                        label="📥 Download Upload File",
                        data=upload_buffer,
                        file_name=f"upload file - {custom_name_excel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.download_button(
                        label="📥 Download Final Sheet",
                        data=final_buffer,
                        file_name=f"FINAL SHEET - {custom_name_excel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("Files processed successfully! Use the download buttons above.")
                    
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
            else:
                st.warning("Please provide a custom name for the output files.")

# Column 2: PDF Data Processor
with col2:
    st.subheader("📑 PDF Data Processor")
    pdf_file = st.file_uploader("Upload PDF File", type=['pdf'])
    with st.expander("Upload Excel Files", expanded=True):
        raw_template = st.file_uploader("Raw Template File (.xlsx)", type=['xlsx'])
        hts_template = st.file_uploader("HTS Template File (.xlsx)", type=['xlsx'])
        final_sheet_template_pdf = st.file_uploader("Final Sheet Template (.xlsx)", type=['xlsx'])

    if pdf_file and raw_template and hts_template and final_sheet_template_pdf:
        htsnum = st.text_input("HTS Number for Output Files")

        if st.button("🚀 Process PDF Files"):
            if htsnum:
                try:
                    # Extract and parse data from PDF
                    clean_text = extract_text_from_pdf(pdf_file)
                    rfq_text = extract_rfq_from_pdf(pdf_file)
                    data = parse_text(clean_text, rfq_text)
                    
                    # Create upload file in memory
                    upload_buffer = io.BytesIO()
                    with pd.ExcelWriter(upload_buffer, engine='openpyxl') as writer:
                        pd.DataFrame(data).to_excel(writer, index=False)
                        workbook = writer.book
                        sheet = workbook.active
                        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
                            for cell in row:
                                cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                    upload_buffer.seek(0)
                    
                    # Create final sheet in memory
                    final_wb = load_workbook(final_sheet_template_pdf)
                    final_ws = final_wb.active
                    
                    for row in final_ws.iter_rows(min_row=2, max_row=final_ws.max_row):
                        for cell in row:
                            cell.value = None
                    
                    paste_row = 2
                    for item in data:
                        final_ws[f'A{paste_row}'] = item['RFx Item No']
                        final_ws[f'B{paste_row}'] = item['Description']
                        final_ws[f'C{paste_row}'] = item['QTY']
                        final_ws[f'D{paste_row}'] = item['UOM']
                        po_text = item['PO Text']
                        formatted_po_text = format_text(po_text)
                        final_ws[f'E{paste_row}'] = formatted_po_text
                        manufacturer_name = manufacture_name(po_text)
                        final_ws[f'G{paste_row}'] = manufacturer_name
                        paste_row += 1
                    
                    for row in final_ws.iter_rows():
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                    
                    final_buffer = io.BytesIO()
                    final_wb.save(final_buffer)
                    final_buffer.seek(0)
                    
                    # Create download buttons
                    st.download_button(
                        label="📥 Download Upload File",
                        data=upload_buffer,
                        file_name=f"upload file - {htsnum}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.download_button(
                        label="📥 Download Final Sheet",
                        data=final_buffer,
                        file_name=f"FINAL SHEET - {htsnum}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("Files processed successfully! Use the download buttons above.")
                    
                except Exception as e:
                    st.error(f"An error occurred while processing the PDF: {str(e)}")
            else:
                st.warning("Please provide an HTS number for the output files.")

# Column 3: List Maker
with col3:
    st.subheader("📝 List Maker")
    manufacturer_file = st.file_uploader("Final Sheet File for Manufacturer", type=['xlsx'])

    if manufacturer_file:
        if st.button("🚀 Process Manufacturer File"):
            try:
                final_output = process_final_sheet_for_manufacturer(manufacturer_file)
                st.text_area("Manufacturer List", final_output, height=300)
                st_copy_to_clipboard(final_output)
                st.success("Manufacturer list generated! Copied to clipboard.")
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
