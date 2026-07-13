# main.py
import functools
import math
import os
import re
import traceback
from collections import defaultdict
from copy import copy
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from PyPDF2 import PdfReader
from mistralai.client import Mistral
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)


# ============================================================
# Streamlit page setup
# ============================================================
st.set_page_config(
    page_title="RFQ Data Processor",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stButton button {
        background: #ff914d;
        color: #fff;
        border-radius: 8px;
        padding: 10px 16px;
        margin-top: 10px;
    }
    .stExpander {
        background: #333;
        border-radius: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# API configuration
# ============================================================
# Since you said you can only change GitHub code, paste the new key here.
# IMPORTANT: A key committed to a public repository is visible to everyone.
API_KEY = "PASTE_YOUR_NEW_MISTRAL_API_KEY_HERE"

MODEL = "mistral-medium-3-5"

# Formatting instructions are stored in prompt.txt in the same GitHub folder.
PROMPT_PATH = Path(__file__).with_name("prompt.txt")

try:
    FORMAT_PROMPT = PROMPT_PATH.read_text(encoding="utf-8").strip()
except OSError as exc:
    st.error(f"Unable to load prompt.txt: {exc}")
    st.stop()

if not FORMAT_PROMPT:
    st.error("prompt.txt is empty. Add the RFQ formatting instructions.")
    st.stop()


def api_key_is_configured() -> bool:
    return bool(
        API_KEY
        and API_KEY.strip()
        and API_KEY != "PASTE_YOUR_NEW_MISTRAL_API_KEY_HERE"
    )


if not api_key_is_configured():
    st.error(
        "Mistral API key is not configured. Open main.py in GitHub and replace "
        'API_KEY = "PASTE_YOUR_NEW_MISTRAL_API_KEY_HERE" with your valid key.'
    )
    st.stop()


client = Mistral(api_key=API_KEY.strip())


# ============================================================
# General helpers
# ============================================================
def _clean(value) -> str:
    if value is None:
        return ""

    if isinstance(value, float) and math.isnan(value):
        return ""

    return str(value).strip()


def workbook_bytes(workbook) -> bytes:
    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def wrap_all(worksheet) -> None:
    """
    Enable wrapping while preserving existing template alignment settings
    as much as possible.
    """
    for row in worksheet.iter_rows():
        for cell in row:
            alignment = copy(cell.alignment)
            alignment.wrap_text = True

            if not alignment.vertical:
                alignment.vertical = "top"

            cell.alignment = alignment


def clear_template_rows(worksheet, start_row: int = 2) -> None:
    """
    Clear values only. Existing styles and template formatting remain.
    """
    for row in worksheet.iter_rows(
        min_row=start_row,
        max_row=worksheet.max_row,
    ):
        for cell in row:
            cell.value = None


def extract_exception_details(exc: Exception) -> str:
    """
    Extract as much useful information as possible from Mistral SDK,
    Tenacity, HTTP, and nested exceptions.
    """
    lines = [
        f"Exception type: {type(exc).__name__}",
        f"Exception message: {exc!s}",
        f"Exception repr: {exc!r}",
    ]

    useful_attributes = (
        "status_code",
        "status",
        "code",
        "message",
        "body",
        "response",
        "request_id",
        "headers",
    )

    for attribute in useful_attributes:
        try:
            value = getattr(exc, attribute, None)
        except Exception:
            value = None

        if value not in (None, "", {}, []):
            lines.append(f"{attribute}: {value!r}")

    cause = getattr(exc, "__cause__", None)
    context = getattr(exc, "__context__", None)

    if cause is not None and cause is not exc:
        lines.extend(
            [
                "",
                f"Caused by: {type(cause).__name__}",
                f"Cause message: {cause!s}",
                f"Cause repr: {cause!r}",
            ]
        )

        for attribute in useful_attributes:
            try:
                value = getattr(cause, attribute, None)
            except Exception:
                value = None

            if value not in (None, "", {}, []):
                lines.append(f"Cause {attribute}: {value!r}")

    if context is not None and context is not exc and context is not cause:
        lines.extend(
            [
                "",
                f"Context: {type(context).__name__}",
                f"Context message: {context!s}",
                f"Context repr: {context!r}",
            ]
        )

    lines.extend(
        [
            "",
            "Traceback:",
            "".join(
                traceback.format_exception(
                    type(exc),
                    exc,
                    exc.__traceback__,
                )
            ),
        ]
    )

    return "\n".join(lines)


def show_api_error(title: str, exc: Exception) -> None:
    st.error(f"{title}: {type(exc).__name__}: {exc!s}")

    with st.expander(f"Detailed diagnostic: {title}", expanded=True):
        st.code(extract_exception_details(exc), language="text")


def get_message_content(response) -> str:
    try:
        content = response.choices[0].message.content
    except Exception as exc:
        raise ValueError(
            f"Unexpected Mistral response structure: {response!r}"
        ) from exc

    if content is None:
        raise ValueError("Mistral returned an empty message.")

    if isinstance(content, str):
        return content.strip()

    return str(content).strip()


# ============================================================
# Mistral diagnostics
# ============================================================
with st.sidebar:
    st.header("API Diagnostics")
    st.caption(
        "Run these tests before processing a file. "
        "The app now uses a supported chat model instead of the retired Agent model."
    )

    masked_key = (
        f"{API_KEY[:4]}...{API_KEY[-4:]}"
        if len(API_KEY) >= 10
        else "Configured"
    )
    st.write(f"API key: `{masked_key}`")
    st.write(f"Formatting model: `{MODEL}`")

    if st.button("1. Test standard chat API", key="test_chat_api"):
        try:
            response = client.chat.complete(
                model=MODEL,
                messages=[
                    {
                        "role": "user",
                        "content": "Reply with exactly: CHAT API OK",
                    }
                ],
                temperature=0,
            )
            st.success(get_message_content(response))
        except Exception as exc:
            show_api_error("Standard chat API test failed", exc)

    if st.button("2. Test RFQ formatting", key="test_format_api"):
        try:
            response = client.chat.complete(
                model=MODEL,
                messages=[
                    {
                        "role": "system",
                        "content": FORMAT_PROMPT,
                    },
                    {
                        "role": "user",
                        "content": (
                            "Basic Data Text:\n"
                            "VALVE, SOLENOID\n"
                            "PRESSURE RATING: 320 BAR\n"
                            "MANUFACTURER: DANA CORPORATION"
                        ),
                    },
                ],
                temperature=0,
            )

            st.success("RFQ formatting test succeeded.")
            st.code(get_message_content(response), language="text")
        except Exception as exc:
            show_api_error("RFQ formatting test failed", exc)

    st.info(
        "Interpretation:\n\n"
        "- Both tests succeed: the API and formatting model are working.\n"
        "- Both tests fail: check the API key, billing, quota, or model access.\n"
        "- Chat succeeds but formatting fails: review prompt.txt or the returned error."
    )


# ============================================================
# Mistral processing wrappers
# ============================================================
@functools.lru_cache(maxsize=1024)
@retry(
    wait=wait_exponential(multiplier=1, min=2, max=20),
    stop=stop_after_attempt(3),
    retry=retry_if_exception_type(Exception),
    reraise=True,
)
def _fmt_uncached(text: str) -> str:
    """
    Format RFQ text using a supported Mistral chat model and prompt.txt.
    """
    cleaned = _clean(text)

    if not cleaned:
        return ""

    response = client.chat.complete(
        model=MODEL,
        messages=[
            {
                "role": "system",
                "content": FORMAT_PROMPT,
            },
            {
                "role": "user",
                "content": cleaned,
            },
        ],
        temperature=0,
    )

    result = get_message_content(response)

    if not result:
        raise ValueError("Mistral returned an empty formatting response.")

    return re.sub(r"`+", "", result).strip()


def format_text(text) -> str:
    cleaned = _clean(text)

    if not cleaned:
        return ""

    try:
        return _fmt_uncached(cleaned)
    except Exception as exc:
        show_api_error("Format-text request failed", exc)
        return cleaned


@functools.lru_cache(maxsize=1024)
@retry(
    wait=wait_exponential(multiplier=1, min=2, max=20),
    stop=stop_after_attempt(3),
    retry=retry_if_exception_type(Exception),
    reraise=True,
)
def _manu_uncached(text: str) -> str:
    cleaned = _clean(text)

    if not cleaned:
        return ""

    response = client.chat.complete(
        model=MODEL,
        messages=[
            {
                "role": "user",
                "content": (
                    "Extract only the manufacturer or maker names from the "
                    "following RFQ text. Return a plain list separated by "
                    "hyphens. Do not add explanations.\n\n"
                    f"RFQ text:\n{cleaned}"
                ),
            }
        ],
        temperature=0,
    )

    return get_message_content(response)


def manufacture_name(text) -> str:
    cleaned = _clean(text)

    if not cleaned:
        return ""

    try:
        return _manu_uncached(cleaned)
    except Exception as exc:
        show_api_error("Manufacturer-name request failed", exc)
        return ""


# ============================================================
# PDF helpers
# ============================================================
def pdf_text(pdf_file) -> str:
    reader = PdfReader(pdf_file)
    extracted_pages = []

    for page_number, page in enumerate(reader.pages, start=1):
        try:
            extracted_pages.append(page.extract_text() or "")
        except Exception as exc:
            raise ValueError(
                f"Could not extract text from PDF page {page_number}: {exc}"
            ) from exc

    return "".join(extracted_pages)


def pdf_clean_body(pdf_file) -> str:
    return re.sub(
        r"(REQUEST FOR QUOTATION[\s\S]*?RFQ Number \d+)",
        "",
        pdf_text(pdf_file),
    )


def parse_pdf(body: str, full_rfq_text: str) -> list[dict]:
    rfx_match = re.search(r"RFQ Number (\d+)", full_rfq_text)
    rfx_number = rfx_match.group(1) if rfx_match else "Unknown"

    item_pattern = re.compile(
        r"(\d{5}) (\w?12\d{10}) "
        r"(\d+(?:\.\d+)?)\s*(\w+) "
        r".*?(\d{2}\.\d{2}\.\d{4})",
        re.DOTALL,
    )

    short_text_pattern = re.compile(
        r"Short Text :(.*?)\n",
        re.DOTALL,
    )

    po_text_pattern = re.compile(
        r"PO Material Text :(.*?)Agreement / LineNo.",
        re.DOTALL,
    )

    items = item_pattern.findall(body)
    short_texts = short_text_pattern.findall(body)
    po_texts = po_text_pattern.findall(body)

    output = []

    for index, item in enumerate(items):
        material_number = (
            item[1]
            if item[1].startswith(("B12", "12", "B16", "15"))
            else ""
        )

        output.append(
            {
                "RFx Number": rfx_number,
                "RFx Item No": item[0],
                "PR Item No": "",
                "Material No": material_number,
                "Description": (
                    short_texts[index]
                    if index < len(short_texts)
                    else ""
                ),
                "PO Text": (
                    po_texts[index]
                    if index < len(po_texts)
                    else ""
                ),
                "QTY": item[2],
                "UOM": item[3],
            }
        )

    return output


# ============================================================
# Main UI
# ============================================================
col1, col2, col3 = st.columns([2, 2, 1])


# ============================================================
# 1. Excel Processor
# ============================================================
with col1:
    st.subheader("Excel Data Processor")

    techno = st.file_uploader(
        "Techno-Commercial Envelope (.xls)",
        type=["xls"],
        key="techno",
    )

    with st.expander("Excel templates", expanded=True):
        uploaded_upload_template = st.file_uploader(
            "Upload template (.xlsx)",
            type=["xlsx"],
            key="tpl_upl",
        )

        uploaded_final_template = st.file_uploader(
            "Final Sheet template (.xlsx)",
            type=["xlsx"],
            key="tpl_fin",
        )

    upload_template = (
        uploaded_upload_template
        if uploaded_upload_template is not None
        else "upload file - HTS.xlsx"
    )

    final_template = (
        uploaded_final_template
        if uploaded_final_template is not None
        else "FINAL SHEET.xlsx"
    )

    suffix = st.text_input(
        "Output name suffix",
        key="suffix_excel",
    )

    if st.button("Process Excel", key="btn_excel"):
        if techno is None:
            st.warning("Upload the Techno-Commercial Envelope.")
        elif not suffix.strip():
            st.warning("Enter an output name suffix.")
        else:
            try:
                filename_number = re.search(r"\d+", techno.name)
                rfx_number = (
                    filename_number.group()
                    if filename_number
                    else "Unknown"
                )

                excel_file = pd.ExcelFile(techno)

                required_columns = {
                    "Description",
                    "InternalNote",
                    "Quantity",
                    "Unit of Measure",
                }

                matching_sheet = None

                for sheet_name in excel_file.sheet_names:
                    header = pd.read_excel(
                        excel_file,
                        sheet_name=sheet_name,
                        nrows=1,
                    )

                    if required_columns.issubset(set(header.columns)):
                        matching_sheet = sheet_name
                        break

                if matching_sheet is None:
                    raise ValueError(
                        "Required columns are missing. Expected: "
                        + ", ".join(sorted(required_columns))
                    )

                dataframe = pd.read_excel(
                    excel_file,
                    sheet_name=matching_sheet,
                    keep_default_na=False,
                )

                descriptions = (
                    dataframe["Description"]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                )

                quantities = (
                    dataframe["Quantity"]
                    .astype(str)
                    .str.strip()
                )

                units = (
                    dataframe["Unit of Measure"]
                    .astype(str)
                    .str.strip()
                )

                valid = dataframe[
                    descriptions.ne("item or lot description")
                    & quantities.ne("")
                    & units.ne("")
                    & units.str.lower().ne("unit of measure")
                ]

                if valid.empty:
                    raise ValueError(
                        "No valid RFQ line items were found in the selected sheet."
                    )

                progress = st.progress(0)
                status = st.empty()
                total_items = len(valid)

                # Upload workbook
                upload_workbook = load_workbook(upload_template)
                upload_sheet = upload_workbook.active
                clear_template_rows(upload_sheet)

                target_row = 2
                line_item = 10

                for _, record in valid.iterrows():
                    upload_sheet[f"A{target_row}"] = rfx_number
                    upload_sheet[f"B{target_row}"] = line_item
                    upload_sheet[f"E{target_row}"] = record["Description"]
                    upload_sheet[f"H{target_row}"] = record["Unit of Measure"]
                    upload_sheet[f"G{target_row}"] = record["Quantity"]
                    upload_sheet[f"F{target_row}"] = record["InternalNote"]
                    upload_sheet[f"I{target_row}"] = record.get("Number", "")

                    line_item += 10
                    target_row += 1

                wrap_all(upload_sheet)

                # Final Sheet workbook
                final_workbook = load_workbook(final_template)
                final_sheet = final_workbook.active
                clear_template_rows(final_sheet)

                target_row = 2
                line_item = 10

                for processed_count, (_, record) in enumerate(
                    valid.iterrows(),
                    start=1,
                ):
                    status.write(
                        f"Processing item {processed_count} of {total_items}"
                    )

                    po_text = record["InternalNote"]

                    final_sheet[f"A{target_row}"] = line_item
                    final_sheet[f"B{target_row}"] = record["Description"]
                    final_sheet[f"C{target_row}"] = record["Quantity"]
                    final_sheet[f"D{target_row}"] = record["Unit of Measure"]
                    final_sheet[f"E{target_row}"] = format_text(po_text)
                    final_sheet[f"G{target_row}"] = manufacture_name(po_text)

                    line_item += 10
                    target_row += 1
                    progress.progress(processed_count / total_items)

                wrap_all(final_sheet)

                st.session_state["excel_upload_bytes"] = workbook_bytes(
                    upload_workbook
                )
                st.session_state["excel_final_bytes"] = workbook_bytes(
                    final_workbook
                )
                st.session_state["excel_suffix"] = suffix.strip()

                status.empty()
                progress.empty()
                st.success("Excel processed.")

            except Exception as exc:
                show_api_error("Excel processing failed", exc)

    if "excel_upload_bytes" in st.session_state:
        saved_suffix = st.session_state["excel_suffix"]

        st.download_button(
            "Download Upload file",
            st.session_state["excel_upload_bytes"],
            file_name=f"upload file - {saved_suffix}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="dl_excel_up",
        )

        st.download_button(
            "Download FINAL SHEET",
            st.session_state["excel_final_bytes"],
            file_name=f"FINAL SHEET - {saved_suffix}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="dl_excel_fin",
        )


# ============================================================
# 2. PDF Processor
# ============================================================
with col2:
    st.subheader("PDF Data Processor")

    pdf = st.file_uploader(
        "RFQ PDF",
        type=["pdf"],
        key="pdf",
    )

    with st.expander("Excel templates", expanded=True):
        uploaded_raw_template = st.file_uploader(
            "Raw template",
            type=["xlsx"],
            key="tpl_raw",
        )

        uploaded_hts_template = st.file_uploader(
            "HTS template",
            type=["xlsx"],
            key="tpl_hts",
        )

        uploaded_pdf_final_template = st.file_uploader(
            "Final Sheet template",
            type=["xlsx"],
            key="tpl_final_pdf",
        )

    raw_template = (
        uploaded_raw_template
        if uploaded_raw_template is not None
        else "raw_template.xlsx"
    )

    hts_template = (
        uploaded_hts_template
        if uploaded_hts_template is not None
        else "upload file - HTS.xlsx"
    )

    pdf_final_template = (
        uploaded_pdf_final_template
        if uploaded_pdf_final_template is not None
        else "FINAL SHEET.xlsx"
    )

    # Kept for compatibility with the original interface.
    _ = raw_template

    hts_number = st.text_input(
        "HTS number",
        key="hts_no",
    )

    if st.button("Process PDF", key="btn_pdf"):
        if pdf is None:
            st.warning("Upload an RFQ PDF.")
        elif not hts_number.strip():
            st.warning("Enter the HTS number.")
        else:
            try:
                full_pdf_text = pdf_text(pdf)
                clean_body = re.sub(
                    r"(REQUEST FOR QUOTATION[\s\S]*?RFQ Number \d+)",
                    "",
                    full_pdf_text,
                )
                data = parse_pdf(clean_body, full_pdf_text)

                if not data:
                    raise ValueError(
                        "No RFQ items were parsed from the PDF. "
                        "The PDF format may not match the current parser."
                    )

                progress = st.progress(0)
                status = st.empty()
                total_items = len(data)

                # Upload workbook
                upload_workbook = load_workbook(hts_template)
                upload_sheet = upload_workbook.active
                clear_template_rows(upload_sheet)

                target_row = 2

                for record in data:
                    mapping = [
                        ("RFx Number", "A"),
                        ("RFx Item No", "B"),
                        ("PR Item No", "C"),
                        ("Material No", "D"),
                        ("Description", "E"),
                        ("PO Text", "F"),
                        ("QTY", "G"),
                        ("UOM", "H"),
                    ]

                    for field, column in mapping:
                        upload_sheet[f"{column}{target_row}"] = record[field]

                    target_row += 1

                wrap_all(upload_sheet)

                # Final Sheet workbook
                final_workbook = load_workbook(pdf_final_template)
                final_sheet = final_workbook.active
                clear_template_rows(final_sheet)

                target_row = 2

                for processed_count, record in enumerate(data, start=1):
                    status.write(
                        f"Processing item {processed_count} of {total_items}"
                    )

                    final_sheet[f"A{target_row}"] = record["RFx Item No"]
                    final_sheet[f"B{target_row}"] = record["Description"]
                    final_sheet[f"C{target_row}"] = record["QTY"]
                    final_sheet[f"D{target_row}"] = record["UOM"]
                    final_sheet[f"E{target_row}"] = format_text(
                        record["PO Text"]
                    )
                    final_sheet[f"G{target_row}"] = manufacture_name(
                        record["PO Text"]
                    )

                    target_row += 1
                    progress.progress(processed_count / total_items)

                wrap_all(final_sheet)

                st.session_state["pdf_upload_bytes"] = workbook_bytes(
                    upload_workbook
                )
                st.session_state["pdf_final_bytes"] = workbook_bytes(
                    final_workbook
                )
                st.session_state["pdf_hts_no"] = hts_number.strip()

                status.empty()
                progress.empty()
                st.success("PDF processed.")

            except Exception as exc:
                show_api_error("PDF processing failed", exc)

    if "pdf_upload_bytes" in st.session_state:
        saved_hts_number = st.session_state["pdf_hts_no"]

        st.download_button(
            "Download Upload file",
            st.session_state["pdf_upload_bytes"],
            file_name=f"upload file - {saved_hts_number}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="dl_pdf_up",
        )

        st.download_button(
            "Download FINAL SHEET",
            st.session_state["pdf_final_bytes"],
            file_name=f"FINAL SHEET - {saved_hts_number}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="dl_pdf_fin",
        )


# ============================================================
# 3. HTS Cleaner and List Maker
# ============================================================
with col3:
    st.subheader("HTS Cleaner")

    hts_upload = st.file_uploader(
        "upload file - HTS.xlsx",
        type=["xlsx"],
        key="hts_clean",
    )

    uploaded_clean_final_template = st.file_uploader(
        "Final Sheet template (optional)",
        type=["xlsx"],
        key="tpl_clean_fin",
    )

    clean_final_template = (
        uploaded_clean_final_template
        if uploaded_clean_final_template is not None
        else "FINAL SHEET.xlsx"
    )

    if st.button("Clean HTS", key="btn_clean_hts"):
        if hts_upload is None:
            st.warning("Upload the HTS workbook.")
        else:
            try:
                source_workbook = load_workbook(hts_upload)
                source_sheet = source_workbook.active

                final_workbook = load_workbook(clean_final_template)
                final_sheet = final_workbook.active
                clear_template_rows(final_sheet)

                source_rows = [
                    row
                    for row in source_sheet.iter_rows(
                        min_row=2,
                        max_row=source_sheet.max_row,
                    )
                    if any(cell.value for cell in row)
                ]

                if not source_rows:
                    raise ValueError(
                        "The uploaded HTS workbook contains no data rows."
                    )

                progress = st.progress(0)
                status = st.empty()
                total_rows = len(source_rows)
                target_row = 2

                for processed_count, source_row in enumerate(
                    source_rows,
                    start=1,
                ):
                    status.write(
                        f"Processing row {processed_count} of {total_rows}"
                    )

                    final_sheet[f"A{target_row}"] = source_row[1].value
                    final_sheet[f"B{target_row}"] = source_row[4].value
                    final_sheet[f"C{target_row}"] = source_row[6].value
                    final_sheet[f"D{target_row}"] = source_row[7].value

                    po_text = source_row[5].value or ""

                    final_sheet[f"E{target_row}"] = format_text(po_text)
                    final_sheet[f"G{target_row}"] = manufacture_name(po_text)

                    target_row += 1
                    progress.progress(processed_count / total_rows)

                wrap_all(final_sheet)

                st.session_state["clean_bytes"] = workbook_bytes(
                    final_workbook
                )

                status.empty()
                progress.empty()
                st.success("HTS cleaned.")

            except Exception as exc:
                show_api_error("HTS cleaning failed", exc)

    if "clean_bytes" in st.session_state:
        st.download_button(
            "Download cleaned FINAL SHEET",
            st.session_state["clean_bytes"],
            file_name="FINAL SHEET - cleaned.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="dl_clean",
        )

    st.subheader("List Maker")

    final_xlsx = st.file_uploader(
        "FINAL SHEET for manufacturers",
        type=["xlsx"],
        key="manuf",
    )

    if st.button("Build list", key="btn_list"):
        if final_xlsx is None:
            st.warning("Upload the FINAL SHEET workbook.")
        else:
            try:
                dataframe = pd.read_excel(final_xlsx)
                output = defaultdict(
                    lambda: {
                        "items": [],
                        "emails": [],
                    }
                )

                for _, row in dataframe.iterrows():
                    manufacturers = _clean(row.get("Manufacturer", ""))

                    if not manufacturers:
                        continue

                    item_number = row.get("Line item number", "")

                    email_values = [
                        row[column]
                        for column in dataframe.columns
                        if (
                            "mail" in str(column).lower()
                            or "unnamed" in str(column).lower()
                        )
                        and pd.notna(row[column])
                    ]

                    for manufacturer in [
                        item.strip()
                        for item in manufacturers.split("-")
                        if item.strip()
                    ]:
                        output[manufacturer]["items"].append(item_number)
                        output[manufacturer]["emails"].extend(email_values)

                result_lines = []

                for manufacturer, values in output.items():
                    unique_items = sorted(
                        {
                            str(item)
                            for item in values["items"]
                            if _clean(item)
                        }
                    )

                    result_lines.append(
                        f"Item {', '.join(unique_items)}: {manufacturer}"
                    )

                    unique_emails = list(
                        dict.fromkeys(
                            _clean(email)
                            for email in values["emails"]
                            if _clean(email)
                        )
                    )

                    if unique_emails:
                        result_lines.extend(unique_emails)

                    result_lines.append("")

                final_output = "\n".join(result_lines)

                st.text_area(
                    "Output",
                    final_output,
                    height=300,
                )

                try:
                    from st_copy_to_clipboard import st_copy_to_clipboard

                    st_copy_to_clipboard(final_output)
                except ImportError:
                    st.warning(
                        "Copy-to-clipboard component is not installed, "
                        "but the generated list is shown above."
                    )

            except Exception as exc:
                show_api_error("List creation failed", exc)
