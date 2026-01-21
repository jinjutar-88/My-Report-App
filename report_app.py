import streamlit as st
import openpyxl
import io
import uuid
import smtplib
import gspread
from google.oauth2.service_account import Credentials
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# --- 🛠 แก้ไขข้อมูลของคุณที่นี่ 🛠 ---
SENDER_EMAIL = "your-email@gmail.com"      # เมลของคุณ (ผู้ส่ง)
SENDER_PASSWORD = "abcd efgh ijkl mnop"   # รหัส 16 หลัก (ดอกที่ 1)
RECEIVER_EMAIL = "target@gmail.com"        # เมลที่จะให้รับรายงาน
GOOGLE_SHEET_NAME = "Smart Dev Report Log" # ชื่อไฟล์ Google Sheet

# --- เริ่มต้นหน้าเว็บ ---
st.set_page_config(page_title="Smart Dev Solution", layout="wide")
st.title("🛠 Smart Dev Solution - Report")

# ระบบจัดการรูปภาพ (Session State)
if 'photo_ids' not in st.session_state:
    st.session_state.photo_ids = [str(uuid.uuid4())]

def add_photo(): st.session_state.photo_ids.append(str(uuid.uuid4()))
def remove_photo(pid): 
    if len(st.session_state.photo_ids) > 1: st.session_state.photo_ids.remove(pid)

# ส่วนกรอกข้อมูลทั่วไป
st.subheader("📋 General Information")
col1, col2 = st.columns(2)
with col1:
    date_issue = st.date_input("Date")
    project_name = st.text_input("Project Name")
    location = st.text_input("Site / Location")
with col2:
    client_name = st.text_input("Client Name")
    service_type = st.selectbox("Service Type", ["Project", "Repairing", "Services", "Training", "Check", "Others"])
    eng_name = st.text_input("Engineer Name")

job_performed = st.text_area("Job Performed", height=150)

# ส่วนจัดการรูปภาพ
st.markdown("---")
st.subheader("📸 Photo Report")
photos_data = []
for i, pid in enumerate(st.session_state.photo_ids):
    with st.container():
        c1, c2 = st.columns([1, 1])
        with c1:
            file = st.file_uploader(f"Upload Photo {i+1}", key=f"f{pid}")
            if file: st.image(file, width=250)
        with c2:
            desc = st.text_area(f"Description {i+1}", key=f"d{pid}")
            if len(st.session_state.photo_ids) > 1:
                st.button("🗑️ Remove", key=f"btn{pid}", on_click=remove_photo, args=(pid,))
        photos_data.append({"file": file, "desc": desc})
        st.write("---")

st.button("➕ Add More Photo", on_click=add_photo)

# ปุ่มดำเนินการหลัก
if st.button("🚀 SUBMIT & SEND REPORT", use_container_width=True):
    with st.spinner('กำลังประมวลผล...'):
        try:
            # 1. สร้างไฟล์ Excel จาก Template
            wb = openpyxl.load_workbook("template.xlsx")
            ws = wb.active
            ws["J5"], ws["H7"], ws["C9"], ws["B16"], ws["D17"] = date_issue.strftime('%d/%m/%Y'), location, client_name, project_name, job_performed
            
            output = io.BytesIO()
            wb.save(output)
            excel_bytes = output.getvalue()

            # 2. บันทึกข้อมูลลง Google Sheet
            try:
                scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
                # ดึงกุญแจจาก Secrets ที่เราตั้งไว้
                creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
                client = gspread.authorize(creds)
                gs = client.open(GOOGLE_SHEET_NAME).sheet1
                gs.append_row([date_issue.strftime('%d/%m/%Y'), project_name, location, client_name, service_type, eng_name])
                st.success("✅ บันทึกลง Google Sheet เรียบร้อย!")
            except Exception as gs_err:
                st.error(f"Google Sheet Error: {gs_err}")

            # 3. ส่งอีเมลพร้อมแนบไฟล์
            try:
                msg = MIMEMultipart()
                msg['From'], msg['To'], msg['Subject'] = SENDER_EMAIL, RECEIVER_EMAIL, f"Report: {project_name}"
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(excel_bytes)
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename=Report_{project_name}.xlsx")
                msg.attach(part)
                
                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                    server.login(SENDER_EMAIL, SENDER_PASSWORD)
                    server.send_message(msg)
                st.success("📧 ส่งอีเมลเรียบร้อย!")
            except Exception as em_err:
                st.error(f"Email Error: {em_err}")

            st.download_button("📥 Download Excel Copy", excel_bytes, f"Report_{project_name}.xlsx")

        except Exception as e:
            st.error(f"System Error: {e}")
