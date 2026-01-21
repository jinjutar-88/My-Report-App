import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
import io
import gspread
from google.oauth2.service_account import Credentials
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# --- 1. CONFIGURATION (ข้อมูลจากรูปภาพของคุณ) ---
SENDER_EMAIL = "jinjutar.smartdev@gmail.com"
SENDER_PASSWORD = "uzfs bdtc xclz rzsq" # App Password 16 หลัก
RECEIVER_EMAIL = "jinjutar.smartdev@gmail.com"
GOOGLE_SHEET_NAME = "Smart Dev Report Log"

# --- 2. ฟังก์ชันส่งอีเมล ---
def send_email_with_report(file_data, filename, doc_no):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = RECEIVER_EMAIL
        msg['Subject'] = f"New Service Report: {doc_no}"

        # แนบไฟล์ Excel
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={filename}')
        msg.attach(part)

        # เชื่อมต่อ Server Gmail
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Email Error: {e}")
        return False

# --- 3. ฟังก์ชันจัดการรูปภาพ ---
def add_image_to_excel(ws, img_file, cell_address):
    if img_file is None: return
    temp_path = f"temp_{cell_address}.png"
    with open(temp_path, "wb") as f:
        f.write(img_file.getbuffer())
    img = Image(temp_path)
    
    for m_range in ws.merged_cells.ranges:
        if cell_address in m_range:
            target_width = 0
            target_height = 0
            for col in range(m_range.min_col, m_range.max_col + 1):
                col_letter = ws.cell(row=1, column=col).column_letter
                target_width += (ws.column_dimensions[col_letter].width or 8.43) * 7.5
            for row in range(m_range.min_row, m_range.max_row + 1):
                target_height += (ws.row_dimensions[row].height or 15) * 1.33
            img.width, img.height = target_width - 10, target_height - 10
            ws.add_image(img, cell_address)
            return
    img.width, img.height = 300, 200
    ws.add_image(img, cell_address)

# --- 4. ระบบจัดการรูปภาพ (Session State) ---
if 'photos' not in st.session_state:
    st.session_state.photos = [0]

def add_photo():
    new_id = max(st.session_state.photos) + 1 if st.session_state.photos else 0
    st.session_state.photos.append(new_id)

def delete_photo(index):
    if len(st.session_state.photos) > 1:
        st.session_state.photos.remove(index)

# --- 5. หน้าเว็บ UI ---
st.set_page_config(page_title="Smart Dev Solution - Report", layout="wide")
st.title("🚀 Smart Dev Solution Report 🚀")

# Input Fields
st.header("📄 Part 1: Document Details")
doc_no = st.text_input("Doc. No.(Indent)")
ref_po_no = st.text_input("Ref. PO No.(ไม่ทราบเขียน - )")
date_issue = st.date_input("Date of Issue", datetime.now())

st.header("🏢 Part 2: Project & Client")
project_name = st.text_input("Project Name")
site_location = st.text_input("Site / Location")
contact_client = st.text_input("Contact Person (Client)")
contact_co_ltd = st.text_input("Contact (ex: Smart Dev Solution Co., Ltd.)")
engineer_name = st.text_input("Engineer Name (Prepared By)")

st.header("🛠 Part 3: Service Details")
service_type = st.selectbox("Service Type", ["Project", "Commissioning", "Repairing", "Services", "Training", "Check", "Other"])
job_performed = st.text_area("Job Performed")

st.header("📸 Part 4: Photo Report")
final_photo_data = []
for i in st.session_state.photos:
    with st.container():
        col_img, col_del = st.columns([8, 1])
        with col_img:
            img = st.file_uploader(f"Upload Image", type=['png', 'jpg', 'jpeg'], key=f"file_{i}")
            desc = st.text_input(f"Description", key=f"desc_{i}", placeholder="พิมพ์คำบรรยาย...")
        with col_del:
            st.write("") 
            st.write("") 
            if st.button("🗑️", key=f"del_{i}"):
                delete_photo(i)
                st.rerun()
        final_photo_data.append({"img": img, "desc": desc})
        st.markdown("---")

st.button("➕ Add More Photo", on_click=add_photo)

# --- 6. ประมวลผลและส่งอีเมล ---
if st.button("🚀 Generate, Save & Send Email", type="primary"):
    try:
        wb = load_workbook("template.xlsx")
        ws = wb.active

        def write_safe(ws, cell_addr, value):
            for m_range in ws.merged_cells.ranges:
                if cell_addr in m_range:
                    ws.cell(row=m_range.min_row, column=m_range.min_col).value = value
                    return
            ws[cell_addr] = value

        # เขียนข้อมูลลง Excel
        write_safe(ws, "B5", doc_no)
        write_safe(ws, "F6", ref_po_no)
        write_safe(ws, "J5", date_issue.strftime('%d/%m/%Y'))
        write_safe(ws, "B16", project_name)
        write_safe(ws, "H7", site_location)
        write_safe(ws, "C9", contact_client)
        write_safe(ws, "A7", contact_co_ltd)
        write_safe(ws, "B42", engineer_name)
        write_safe(ws, "D17", job_performed)

        # จัดการรูปภาพ
        loc_map = ["A49", "A65", "A81", "A97", "A113"] 
        desc_map = ["H49", "H65", "H81", "H97", "H113"]
        count = 0
        for item in final_photo_data:
            if item["img"] and count < len(loc_map):
                add_image_to_excel(ws, item["img"], loc_map[count])
                write_safe(ws, desc_map[count], item["desc"])
                count += 1

        excel_out = io.BytesIO()
        wb.save(excel_out)
        file_bytes = excel_out.getvalue()
        filename = f"Report_{doc_no}.xlsx"
        
        # --- ส่งอีเมล ---
        with st.spinner('กำลังส่งอีเมล...'):
            if send_email_with_report(file_bytes, filename, doc_no):
                st.success(f"📩 ส่งอีเมลไปยัง {RECEIVER_EMAIL} สำเร็จ!")

        # บันทึกลง Google Sheet
        try:
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            client = gspread.authorize(creds)
            gs = client.open(GOOGLE_SHEET_NAME).sheet1
            gs.append_row([date_issue.strftime('%d/%m/%Y'), doc_no, project_name, engineer_name, datetime.now().strftime('%H:%M:%S')])
        except: pass 

        st.success("🎉 บันทึกรายงานสำเร็จ!")
        st.download_button("📥 Download Excel Copy", file_bytes, filename)
        st.balloons()

    except Exception as e:
        st.error(f"🚨 Error: {e}")
