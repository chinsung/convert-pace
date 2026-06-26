"""
Streamlit web app: Extract training pace zones (A1, L1a, L1b, L2, L3, L4, L5, L6, L7)
from Ergonizer "Running field test" PDF reports and export to Excel.

Run with:
    streamlit run app.py
"""

import io
import os
import re
import json

import pdfplumber
import streamlit as st
import pandas as pd
from openpyxl import Workbook

COLUMNS = ["File", "Name", "Test Date", "A1", "L1a", "L1b", "L2", "L3", "L4", "L5", "L6", "L7"]

ENABLE_GOOGLE_SHEET_IMPORT = False  # ตั้งเป็น True เมื่อพร้อมใช้ฟีเจอร์ Import to Google Sheet


def extract_name_from_filename(filename: str):
    """
    Filename pattern examples (supports space OR underscore separators):
        'Chin Thirawee 25062026 Running field test.pdf'   -> 'Chin_Thirawee'
        'Ball_Warong_04062026_Running_field_test.pdf'      -> 'Ball_Warong'

    Takes the words that appear before the 8-digit date, joins with underscore.
    """
    base = os.path.splitext(filename)[0]
    match = re.match(r"^(.*?)[\s_]+\d{8}[\s_]+Running[\s_]field[\s_]test", base, re.IGNORECASE)
    if not match:
        return None
    name_part = match.group(1).strip()
    # Normalize: split on either space or underscore, rejoin with underscore
    parts = re.split(r"[\s_]+", name_part)
    return "_".join(p for p in parts if p)


def mmss_to_seconds(mmss: str) -> int:
    m, s = mmss.split(":")
    return int(m) * 60 + int(s)


def seconds_to_mmss(total_seconds: int) -> str:
    m, s = divmod(total_seconds, 60)
    return f"{m}:{s:02d}"


def extract_name_and_date(text: str):
    name_match = re.search(r"Performance diagnostics for ([^,]+),\s*([^,]+),\s*b\.", text)
    date_match = re.search(r"On (\d{1,2}/\d{1,2}/\d{4}), a multi-stage test", text)
    name = None
    if name_match:
        last, first = name_match.group(1).strip(), name_match.group(2).strip()
        name = f"{first} {last}"
    date = date_match.group(1) if date_match else None
    return name, date


def extract_base_paces(text: str):
    l1b_match = re.search(r"slower than (\d{1,2}:\d{2}) min", text)
    per1000_matches = re.findall(r"(\d{1,2}:\d{2}) min\s*-\s*(\d{1,2}:\d{2}) min", text)

    if len(per1000_matches) < 3:
        raise ValueError("ไม่พบช่วง pace (MER/SER/EIT) ครบ 3 ช่วงในไฟล์นี้ — รูปแบบไฟล์อาจไม่ตรง template")
    if not l1b_match:
        raise ValueError("ไม่พบค่า L1b ('slower than X min') ในไฟล์นี้")

    mer_pair, ser_pair, eit_pair = per1000_matches[0], per1000_matches[1], per1000_matches[2]

    return {
        "L1b": mmss_to_seconds(l1b_match.group(1)),
        "L2":  mmss_to_seconds(mer_pair[1]),
        "L3":  mmss_to_seconds(eit_pair[0]),
        "L4":  mmss_to_seconds(ser_pair[1]),
        "L5":  mmss_to_seconds(eit_pair[1]),
    }


def compute_zones(base_paces: dict):
    l1b = base_paces["L1b"]
    l2 = base_paces["L2"]
    l3 = base_paces["L3"]
    l4 = base_paces["L4"]
    l5 = base_paces["L5"]

    l1a = l1b + 15
    a1 = l1a + 25
    l6 = l5 - 15
    l7 = l6 - 15

    return {
        "A1": a1, "L1a": l1a, "L1b": l1b,
        "L2": l2, "L3": l3, "L4": l4, "L5": l5,
        "L6": l6, "L7": l7,
    }


def process_pdf(file_name: str, file_bytes):
    with pdfplumber.open(file_bytes) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    pdf_name, date = extract_name_and_date(text)
    filename_name = extract_name_from_filename(file_name)
    name = filename_name or pdf_name  # prefer filename pattern, fall back to PDF text

    base_paces = extract_base_paces(text)
    zones = compute_zones(base_paces)

    row = {"File": file_name, "Name": name, "Test Date": date}
    for key in ["A1", "L1a", "L1b", "L2", "L3", "L4", "L5", "L6", "L7"]:
        row[key] = seconds_to_mmss(zones[key])
    return row


def to_excel_bytes(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Training Zones"
    ws.append(COLUMNS)
    for row in rows:
        ws.append([row.get(c, "") for c in COLUMNS])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def import_to_google_sheet(rows, sheet_id: str, credentials_json: dict):
    """
    Append rows to the bottom of the first worksheet in the given Google Sheet.
    Returns the sheet URL on success.
    """
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(credentials_json, scopes=scopes)
    client = gspread.authorize(creds)

    sheet = client.open_by_key(sheet_id).sheet1

    # If the sheet is empty, write the header first
    if not sheet.get_all_values():
        sheet.append_row(COLUMNS)

    for row in rows:
        sheet.append_row([row.get(c, "") for c in COLUMNS])

    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"


# ---------------- UI ----------------

st.set_page_config(page_title="Training Zone Extractor", page_icon="🏃", layout="centered")
st.title("🏃 Training Zone Extractor")
st.caption("อัปโหลด PDF รายงาน Ergonizer (Running field test) เพื่อดึงค่า pace แต่ละโซนอัตโนมัติ")

uploaded_files = st.file_uploader(
    "เลือกไฟล์ PDF (เลือกได้หลายไฟล์)",
    type="pdf",
    accept_multiple_files=True,
)

if uploaded_files:
    rows = []
    errors = []

    for f in uploaded_files:
        try:
            row = process_pdf(f.name, f)
            rows.append(row)
        except Exception as e:
            errors.append((f.name, str(e)))

    if rows:
        st.subheader("ผลลัพธ์")
        df = pd.DataFrame(rows, columns=COLUMNS)
        st.dataframe(df, use_container_width=True)

        excel_bytes = to_excel_bytes(rows)
        st.download_button(
            "⬇️ ดาวน์โหลด Excel",
            data=excel_bytes,
            file_name="training_zones_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if ENABLE_GOOGLE_SHEET_IMPORT:
            st.divider()
            st.subheader("📤 Import to Google Sheet")

            with st.expander("ตั้งค่า Google Sheet (ทำครั้งแรกครั้งเดียว ใส่ใหม่ทุกครั้งที่เปิดแอป)"):
                sheet_id_input = st.text_input(
                    "Google Sheet ID",
                    help="คัดลอกจาก URL ของชีต ส่วนระหว่าง /d/ กับ /edit",
                )
                credentials_file = st.file_uploader(
                    "Service Account JSON file",
                    type="json",
                    key="gcp_creds",
                )

            if st.button("📤 Import to Google Sheet"):
                if not sheet_id_input or not credentials_file:
                    st.warning("กรุณากรอก Sheet ID และอัปโหลดไฟล์ credentials .json ก่อน")
                else:
                    try:
                        creds_dict = json.load(credentials_file)
                        sheet_url = import_to_google_sheet(rows, sheet_id_input.strip(), creds_dict)
                        st.success("Import สำเร็จ!")
                        st.markdown(f"🔗 [เปิด Google Sheet]({sheet_url})")
                    except Exception as e:
                        st.error(f"Import ไม่สำเร็จ: {e}")

    if errors:
        st.subheader("⚠️ ไฟล์ที่อ่านไม่สำเร็จ")
        for fname, err in errors:
            st.error(f"{fname}: {err}")
else:
    st.info("ลากไฟล์ PDF มาวาง หรือกดเลือกไฟล์ด้านบนเพื่อเริ่ม")