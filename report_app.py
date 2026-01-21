import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import io
import uuid
import smtplib
import gspread
from google.oauth2.service_account import Credentials
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from PIL import Image

# --- 🛠 ส่วนที่ 1: CONFIGURATION (ใส่ข้อมูลของคุณตรงนี้) ---
SENDER_EMAIL = "jinjutar.smartdev@gmail.com"      
SENDER_PASSWORD = "uzfs bdtc xclz rzsq" # รหัส 16 หลักจาก Google
RECEIVER_EMAIL = "jinjutar.smartdev@gmail.com"        
GOOGLE_SHEET_NAME = "Smart Dev Report Log" 

# --- ส่วนที่ 2: ตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="Smart Dev Solution - Report", layout="wide")
st.title("🛠 Smart Dev Solution - Report")

if 'photo_ids' not in st.session_state:
    st.session_state.photo_ids = [str(uuid.uuid4())]

def add_photo(): st.session_state.photo_ids.append(str(uuid.uuid4()))
def remove_photo(pid): 
    if len(st.session_state.photo_ids) > 1: st.session_state.photo_ids.remove(pid)

# --- แบบฟอร์มกรอกข้อมูล ---
st.subheader("📋 Part 1: General Information")
col1, col2 = st.columns(2)
with col1:
    date_issue = st.date_input("Date of Issue")
    ref_po_no = st.text_input("Ref. PO No.")
    project_name = st.text_input("Project Name")
    location = st.text_input("Site / Location")
with col2:
    doc_no = st.text_input("Doc. No.")
    client_name = st.text_input("Contact Person (Client)")
    contact_co_ltd = st.text_input("Contact (Co., Ltd.)")
    service_type = st.selectbox("Service Type", ["Project", "Repairing", "Services", "Training", "Check", "Others"])
eng_name = st.text_input("Engineer Name")

st.markdown("---")
st.subheader("🔧 Part 2: Service Details")
job_performed = st.text_area("Job Performed", height=150)
note = st.text_area("Note")

st.markdown("---")
st.subheader("📸 Part 3: Photo Report")
photos_data = []
for i, pid in enumerate(st.session_state.photo_ids):
    with st.container():
        c1, c2 = st.columns([1, 1])
        with c1:
            up_file = st.file_uploader(f"Upload Photo {i+1}", type=['jpg','jpeg','png'], key=f"f{pid}")
            if up_file: st.image(up_file, width=250)
        with c2:
            desc = st.text_area(f"Description for Photo {i+1}", key=f"d{pid}", height=120)
            if len(st.session_state.photo_ids) > 1:
                st.button("🗑️ Remove", key=f"r{pid}", on_click=remove_photo, args=(pid,))
        photos_data.append({"file": up_file, "desc": desc})
        st.write("---")
st.button("➕ Add More Photo", on_click=add_photo)

# --- ส่วนประมวลผล ---
if st.button("🚀 SUBMIT & SEND", use_container_width=True):
    if not project_name:
        st.error("กรุณากรอก Project Name")
    else:
        with st.spinner('กำลังประมวลผล...'):
            try:
                wb = openpyxl.load_workbook("template.xlsx")
                ws = wb.active
                
                # ฟังก์ชันป้องกัน MergedCell Error
                def safe_write(cell, val):
                    try: ws[cell] = val
                    except: pass

                safe_write("J5", date_issue.strftime('%d/%m/%Y'))
                safe_write("B5", doc_no)
                safe_write("F6", ref_po_no)
                safe_write("H7", location)
                safe_write("C9", client_name)
                safe_write("A7", contact_co_ltd)
                safe_write("B16", project_name)
                safe_write("D17", job_performed)
                safe_write("B36", note)

                # ใส่รูปและคำบรรยาย (เริ่มแถว 49)
                start_row = 49
                for i, data in enumerate(photos_data):
                    if data["file"]:
                        cur_row = start_row + (i * 20)
                        safe_write(f"H{cur_row}", data["desc"])
                        img_pil = Image.open(data["file"])
                        img_pil.thumbnail((400, 400))
                        img_io = io.BytesIO()
                        img_pil.save(img_io, format='PNG')
                        xl_img = XLImage(img_io)
                        ws.add_image(xl_img, f"A{cur_row}")

                excel_io = io.BytesIO()
                wb.save(excel_io)
                excel_bytes = excel_io.getvalue()
                # --- ส่วนบันทึก Google Sheets ---
                try:
                    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
                    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
                    client = gspread.authorize(creds)
                    
                    sheet = client.open(GOOGLE_SHEET_NAME).sheet1
                    
                    row = [
                        date_issue.strftime('%d/%m/%Y'), 
                        project_name, 
                        location, 
                        eng_name, 
                        datetime.now().strftime('%H:%M:%S')
                    ]
                    
                    sheet.append_row(row)
                    st.success("✅ บันทึกลง Google Sheet สำเร็จ")
                except Exception as e:
                    # ถ้า Error แล้วมีเลข 200 แสดงว่าจริงๆ แล้วมันบันทึกสำเร็จ
                    if "200" in str(e):
                        st.success("✅ บันทึกลง Google Sheet เรียบร้อยแล้ว")
                    else:
                        st.warning(f"⚠️ Sheet Connection: {e}")

                # --- ส่วนส่งอีเมล ---
                try:
                    msg = MIMEMultipart()
                    msg['From'] = SENDER_EMAIL
                    msg['To'] = RECEIVER_EMAIL
                    msg['Subject'] = f"Report: {project_name}"
                    
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(excel_bytes)
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename=Report_{project_name}.xlsx")
                    msg.attach(part)
                    
                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                        server.login(SENDER_EMAIL, SENDER_PASSWORD)
                        server.send_message(msg)
                    st.success("📧 ส่งอีเมลสำเร็จ!")
                except Exception as e:
                    st.error(f"❌ Email Error: {e}")

                # ปุ่มดาวน์โหลด
                st.download_button("📥 Download Excel", excel_bytes, f"Report_{project_name}.xlsx")

            except Exception as e:
                st.error(f"🚨 System Error: {e}")
