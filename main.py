import streamlit as st
from st_copy_to_clipboard import st_copy_to_clipboard
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import re
from collections import defaultdict
import openai
import time
import io  # <<--- NEW: For in-memory files
from mistralai import Mistral

# ... (Your other code, models, and helper functions remain unchanged) ...

# -- All your helper functions go here as before (not shown to save space) --

st.set_page_config(page_title="Data Processor", layout="wide", initial_sidebar_state="collapsed")

# ... (Custom CSS styling unchanged) ...

col1, col2, col3 = st.columns([2, 2, 1])

# ---- Column 1: Excel Data Processor ----
with col1:
    st.subheader("🗃️ Excel Data Processor")
    techno_commercial_file = st.file_uploader("Upload Techno Commercial Envelope File (.xls)", type=['xls'])
    with st.expander("Upload Excel Files", expanded=True):
        upload_file = st.file_uploader("Upload File (.xlsx)", type=['xlsx'])
        final_sheet_file = st.file_uploader("Final Sheet File (.xlsx)", type=['xlsx'])

    # Ensure default paths if files are not uploaded
    if not upload_file:
        upload_file = "upload file - HTS.xlsx"
    if not final_sheet_file:
        final_sheet_file = "FINAL SHEET.xlsx"

    if techno_commercial_file:
        custom_name_excel = st.text_input("Custom Name for 'Upload HTS'")
        # Removed Save Path input field

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

                    # Process upload file
                    upload_bytes = io.BytesIO()
                    workbook = load_workbook(upload_file)
                    sheet = workbook.active
                    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                        for cell in row:
                            cell.value = None
                    paste_row = 2
                    rfx_item_no = 10
                    for i in range(len(techno_df)):
                        if pd.notna(techno_df['Description'].iloc[i]) and i != 1:
                            sheet[f'A{paste_row}'] = rfx_number
                            sheet[f'B{paste_row}'] = rfx_item_no
                            sheet[f'E{paste_row}'] = techno_df['Description'].iloc[i]
                            sheet[f'H{paste_row}'] = techno_df['Unit of Measure'].iloc[i]
                            sheet[f'G{paste_row}'] = techno_df['Quantity'].iloc[i]
                            sheet[f'F{paste_row}'] = techno_df['InternalNote'].iloc[i]
                            sheet[f'I{paste_row}'] = techno_df['Number'].iloc[i]
                            paste_row += 1
                            rfx_item_no += 10
                    for row in sheet.iter_rows():
                        for cell in row:
                            cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                    workbook.save(upload_bytes)
                    upload_bytes.seek(0)

                    # Process final sheet
                    final_sheet_bytes = io.BytesIO()
                    final_workbook = load_workbook(final_sheet_file)
                    final_sheet = final_workbook.active
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
                    final_workbook.save(final_sheet_bytes)
                    final_sheet_bytes.seek(0)

                    st.success("Data has been successfully processed.")
                    st.download_button(
                        label="⬇️ Download Upload File",
                        data=upload_bytes,
                        file_name=f"upload file - {custom_name_excel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.download_button(
                        label="⬇️ Download Final Sheet",
                        data=final_sheet_bytes,
                        file_name=f"FINAL SHEET - {custom_name_excel}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"An error occurred: {e}")
            else:
                st.warning("Please provide a custom name.")

# ---- Column 2: PDF Data Processor ----
with col2:
    st.subheader("📑 PDF Data Processor")
    pdf_file = st.file_uploader("Upload PDF File", type=['pdf'])
    with st.expander("Upload Excel Files", expanded=True):
        created_excel_template = st.file_uploader("Raw Template File (.xlsx)", type=['xlsx'])
        template_excel_path = st.file_uploader("HTS Template File (.xlsx)", type=['xlsx'])
        pdf_final_sheet = st.file_uploader("Final Sheet Template (.xlsx)", type=['xlsx'])

    # Default paths for missing uploads
    if not created_excel_template:
        created_excel_template = "raw_template.xlsx"
    if not template_excel_path:
        template_excel_path = "upload file - HTS.xlsx"
    if not pdf_final_sheet:
        pdf_final_sheet = "FINAL SHEET.xlsx"

    if pdf_file:
        htsnum = st.text_input("HTS Number")
        # Removed Save Path input field

        if st.button("🚀 Process PDF Files"):
            if htsnum:
                try:
                    # Save all processed files in memory buffers
                    final_output_buffer = io.BytesIO()
                    final_sheet_output_buffer = io.BytesIO()

                    final_output_path = f"upload file - {htsnum}.xlsx"
                    final_sheet_output_path = f"FINAL SHEET - {htsnum}.xlsx"

                    # Process PDF to intermediate Excel file
                    process_pdf_to_final_excel(pdf_file, created_excel_template, template_excel_path, final_output_buffer)
                    final_output_buffer.seek(0)

                    # Process the final sheet using the content from final_output_buffer
                    process_final_sheet_from_pdf(pdf_final_sheet, final_output_buffer, final_sheet_output_buffer)
                    final_sheet_output_buffer.seek(0)

                    st.success(f"PDF data processed.")
                    st.download_button(
                        label="⬇️ Download Upload File",
                        data=final_output_buffer,
                        file_name=final_output_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.download_button(
                        label="⬇️ Download Final Sheet",
                        data=final_sheet_output_buffer,
                        file_name=final_sheet_output_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"An error occurred while processing the PDF: {e}")
            else:
                st.warning("Please provide an HTS number.")

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
            finalsheet = "FINAL SHEET.xlsx"
            final_output = process_final_sheet_for_manufacturer(finalsheet)
            st.text_area("Formatted Output", final_output, height=300)
            st_copy_to_clipboard(final_output)
        except Exception as e:
            st.error(f"An error occurred: {e}")
