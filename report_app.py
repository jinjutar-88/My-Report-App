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

# --- CONFIGURATION (ตรวจสอบข้อมูลให้ถูกต้อง) ---
SENDER_EMAIL = "your-email@gmail.com"      
SENDER_PASSWORD = "your-16-digit-app-password"   
RECEIVER_EMAIL = "target@gmail.com"        
GOOGLE_SHEET_NAME = "Smart Dev Report Log" 

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Smart Dev Solution - Report", layout="wide")
st.title("🛠 Smart Dev Solution - Report")

# --- Session State สำหรับจัดการรูปภาพ ---
if 'photo_ids' not in st.session_state:
    st.session_state.photo_ids = [str(uuid.uuid4())]

def add_photo(): st.session_state.photo_ids.append(str(uuid.uuid4()))
def remove_photo(pid): 
    if len(st.session_state.photo_ids) > 1: st.session_state.photo_ids.remove(pid)

# --- PART 1: ข้อมูลทั่วไป (General Information) ---
st.subheader("📋 Part 1: General Information")
col1, col2 = st.columns(2)

with col1:
    date_issue = st.date_input("Date of Issue")
    ref_po_no = st.text_input("Ref. PO No.") # เพิ่ม PO
    project_name = st.text_input("Project Name")
    location = st.text_input("Site / Location")

with col2:
    doc_no = st.text_input("Doc. No.") # เพิ่ม Doc No.
    client_name = st.text_input("Contact Person (Client)")
    contact_co_ltd = st.text_input("Contact (Co., Ltd.)") # เพิ่ม Co Ltd.
    service_type = st.selectbox("Service Type", ["Project", "Repairing", "Services", "Training", "Check", "Others"])

eng_name = st.text_input("Engineer Name (Prepared By)")

# --- PART 2: รายละเอียดงาน (Service Details) ---
st.markdown("---")
st.subheader("🔧 Part 2: Service Details")
job_performed = st.text_area("Job Performed (รายละเอียดงานที่ปฏิบัติ)", height=150)
note = st.text_area("Note (หมายเหตุเพิ่มเติม)") # เพิ่ม Note

# --- PART 3: รูปภาพและคำบรรยาย (Photo Report) ---
st.markdown("---")
st.subheader("📸 Part 3: Photo Report")

photos_data = []
for i, photo_id in enumerate(st.session_state.photo_ids):
    with st.container():
        head_col, del_col = st.columns([10, 1])
        with head_col: st.write(f"**Photo {i+1}**")
        with del_col:
            if len(st.session_state.photo_ids) > 1:
                st.button("🗑️", key=f"del_{photo_id}", on_click=remove_photo, args=(photo_id,))
        
        col_img, col_txt = st.columns([1, 1])
        with col_img:
            up_file = st.file_uploader(f"Upload Photo {i+1}", type=['jpg','jpeg','png'], key=f"file_{photo_id}")
            if up_file: st.image(up_file, width=300)
        with col_txt:
            desc = st.text_area(f"Description for Photo {i+1}", key=f"desc_{photo_id}", height=150)
        
        photos_data.append({"file": up_file, "desc": desc})
        st.markdown("---")

st.button("➕ Add More Photo", on_click=add_photo)

# --- ส่วนประมวลผล ---
st.write(" ")
if st.button("🚀 SUBMIT, SEND EMAIL & SYNC TO SHEET", use_container_width=True):
    with st.spinner('กำลังดำเนินการ...'):
        try:
            # 1. จัดการ Excel (Template)
            wb = openpyxl.load_workbook("template.xlsx")
            ws = wb.active
            # เติมข้อมูลลงตำแหน่งเดิมของคุณ
            ws["J5"] = date_issue.strftime('%d/%m/%Y')
            ws["H7"] = location
            ws["C9"] = client_name
            ws["B16"] = project_name
            ws["D17"] = job_performed
            
            output = io.BytesIO()
            wb.save(output)
            excel_bytes = output.getvalue()

            # 2. บันทึกลง Google Sheet
            try:
                scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
                creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
                client = gspread.authorize(creds)
                gs = client.open(GOOGLE_SHEET_NAME).sheet1
                # บันทึกข้อมูลแบบแถวเดียว
                gs.append_row([
                    date_issue.strftime('%d/%m/%Y'), 
                    project_name, 
                    location, 
                    client_name, 
                    service_type, 
                    eng_name,
                    datetime.now().strftime('%H:%M:%S')
                ])
                st.success("✅ บันทึกลง Google Sheet เรียบร้อย!")
            except Exception as gs_err:
                st.error(f"Google Sheet Error: {gs_err}")

            # 3. ส่งอีเมล
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
