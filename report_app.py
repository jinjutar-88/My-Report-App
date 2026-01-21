import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
import io
import gspread
from google.oauth2.service_account import Credentials

# --- 1. การตั้งค่าพื้นฐาน ---
GOOGLE_SHEET_NAME = "Smart Dev Report Log"

# --- 2. ฟังก์ชันจัดการรูปภาพให้พอดีช่อง ---
def add_image_to_excel(ws, img_file, cell_address):
    if img_file is None: return
    temp_path = f"temp_{cell_address}.png"
    with open(temp_path, "wb") as f:
        f.write(img_file.getbuffer())
    img = Image(temp_path)
    
    # ตรวจสอบพื้นที่ (รองรับช่องที่ Merge)
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
    
    img.width, img.height = 300, 200 # Default หากไม่พบช่อง Merge
    ws.add_image(img, cell_address)

# --- 3. ระบบจัดการรูปภาพ (Session State) ---
if 'photos' not in st.session_state:
    st.session_state.photos = [0] # เริ่มต้นด้วย 1 รูป (ID 0)

def add_photo():
    new_id = max(st.session_state.photos) + 1 if st.session_state.photos else 0
    st.session_state.photos.append(new_id)

def delete_photo(index):
    if len(st.session_state.photos) > 1:
        st.session_state.photos.remove(index)

# --- 4. หน้าเว็บ UI ---
st.title("🚀 Smart Dev Report Generator")

# Part 1: ข้อมูลเอกสาร
st.header("📄 Part 1: Document Details")
doc_no = st.text_input("Doc. No.")
ref_po_no = st.text_input("Ref. PO No.")
date_issue = st.date_input("Date of Issue", datetime.now())

# Part 2: ข้อมูลโครงการและผู้ติดต่อ
st.header("🏢 Part 2: Project & Client")
project_name = st.text_input("Project Name")
site_location = st.text_input("Site / Location")
contact_client = st.text_input("Contact Person (Client)")
# ปรับแก้ข้อความตามที่ต้องการ: contact (ex: smart dev solution co,. Ltd.)
contact_co_ltd = st.text_input("Contact (ex: Smart Dev Solution Co., Ltd.)")
engineer_name = st.text_input("Engineer Name (Prepared By)")

# Part 3: รายละเอียดงาน
st.header("🛠 Part 3: Service Details")
service_type = st.selectbox("Service Type", ["New", "Commissioning", "Repairing", "Services", "Training", "Check", "Other"])
job_performed = st.text_area("Job Performed")

# Part 4: รายงานรูปภาพ (แบบ Dynamic พร้อมปุ่มถังขยะ)
st.header("📸 Part 4: Photo Report")

final_photo_data = []

for i in st.session_state.photos:
    with st.container():
        # แบ่งคอลัมน์: ช่องกรอกข้อมูล(8) และ ปุ่มถังขยะ(1)
        col_img, col_del = st.columns([8, 1])
        
        with col_img:
            img = st.file_uploader(f"Upload Image", type=['png', 'jpg', 'jpeg'], key=f"file_{i}")
            desc = st.text_input(f"Description", key=f"desc_{i}", placeholder="พิมพ์คำบรรยายรูปภาพที่นี่...")
        
        with col_del:
            st.write("") # เว้นระยะให้ปุ่มตรงกับช่องอัปโหลด
            st.write("") 
            if st.button("🗑️", key=f"del_{i}"):
                delete_photo(i)
                st.rerun()
        
        final_photo_data.append({"img": img, "desc": desc})
        st.markdown("---")

# ปุ่มเพิ่มรูป
st.button("➕ Add More Photo", on_click=add_photo)

# --- 5. ส่วนประมวลผลเมื่อกดปุ่ม Submit ---
st.markdown("###")
if st.button("🚀 Generate & Save Report", type="primary"):
    try:
        wb = load_workbook("template.xlsx")
        ws = wb.active

        # Mapping ข้อมูลลงใน Excel
        ws["B5"] = f"Doc.No. : {doc_no}"
        ws["F6"] = f"Ref.PO.No. : {ref_po_no}"
        ws["J5"] = date_issue.strftime('%d/%m/%Y')
        ws["B16"] = project_name
        ws["H7"] = site_location
        ws["B10"] = contact_client
        ws["A7"] = contact_co_ltd
        ws["B42"] = engineer_name
        ws["B21"] = job_performed

        # พิกัดสำหรับรูปภาพและคำบรรยาย (ปรับได้ตาม Template จริง)
        loc_map = ["A49"]
        desc_map = ["H49"]
        count = 0
        for item in final_photo_data:
            if item["img"] and count < len(loc_map):
                add_image_to_excel(ws, item["img"], loc_map[count])
                ws[desc_map[count]] = item["desc"]
                count += 1

        # เตรียมไฟล์ Excel สำหรับดาวน์โหลด
        excel_out = io.BytesIO()
        wb.save(excel_out)
        
        # บันทึกประวัติลง Google Sheet
        try:
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            client = gspread.authorize(creds)
            gs = client.open(GOOGLE_SHEET_NAME).sheet1
            gs.append_row([date_issue.strftime('%d/%m/%Y'), doc_no, project_name, engineer_name, datetime.now().strftime('%H:%M:%S')])
        except:
            pass # ถ้า Sheet ต่อไม่ได้ก็ยังให้โหลดไฟล์ได้

        st.success("🎉 รายงานถูกสร้างเรียบร้อยแล้ว!")
        st.download_button("📥 Download Excel Report", excel_out.getvalue(), f"Report_{doc_no}.xlsx")
        st.balloons()

    except Exception as e:
        st.error(f"🚨 เกิดข้อผิดพลาด: {e}")
