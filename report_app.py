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

# --- 3. ระบบจัดการรูปภาพ (Session State) ---
if 'photos' not in st.session_state:
    st.session_state.photos = [0]

def add_photo():
    new_id = max(st.session_state.photos) + 1 if st.session_state.photos else 0
    st.session_state.photos.append(new_id)

def delete_photo(index):
    if len(st.session_state.photos) > 1:
        st.session_state.photos.remove(index)

# --- 4. หน้าเว็บ UI ---
st.title("🚀 Smart Dev Report Generator")

st.header("📄 Part 1: Document Details")
doc_no = st.text_input("Doc. No.")
ref_po_no = st.text_input("Ref. PO No.")
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
            desc = st.text_input(f"Description", key=f"desc_{i}", placeholder="พิมพ์คำบรรยายรูปภาพที่นี่...")
        with col_del:
            st.write("") 
            st.write("") 
            if st.button("🗑️", key=f"del_{i}"):
                delete_photo(i)
                st.rerun()
        final_photo_data.append({"img": img, "desc": desc})
        st.markdown("---")

st.button("➕ Add More Photo", on_click=add_photo)

# --- 5. ส่วนประมวลผลเมื่อกดปุ่ม Submit ---
if st.button("🚀 Generate & Save Report", type="primary"):
    try:
        wb = load_workbook("template.xlsx")
        ws = wb.active

        # ฟังก์ชันแก้ไขปัญหา MergedCell Read-only
        def write_safe(ws, cell_addr, value):
            # ตรวจสอบว่าเซลล์นั้นอยู่ในช่วงที่ Merge หรือไม่
            for m_range in ws.merged_cells.ranges:
                if cell_addr in m_range:
                    # ถ้าเจอว่า Merge ให้เขียนลงที่เซลล์ซ้ายบนสุดของ Range นั้นๆ
                    ws.cell(row=m_range.min_row, column=m_range.min_col).value = value
                    return
            # ถ้าไม่ Merge เขียนปกติ
            ws[cell_addr] = value

        # Mapping ข้อมูล (ตามพิกัดที่คุณแจ้ง)
        write_safe(ws, "B5", Doc.No.)
        write_safe(ws, "F6", Ref.PO.No.)
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
        
        # ส่วน Google Sheet
        try:
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            client = gspread.authorize(creds)
            gs = client.open(GOOGLE_SHEET_NAME).sheet1
            gs.append_row([date_issue.strftime('%d/%m/%Y'), doc_no, project_name, engineer_name, datetime.now().strftime('%H:%M:%S')])
        except: pass 

        st.success("🎉 รายงานถูกสร้างเรียบร้อยแล้ว!")
        st.download_button("📥 Download Excel Report", excel_out.getvalue(), f"Report_{doc_no}.xlsx")
        st.balloons()

    except Exception as e:
        st.error(f"🚨 เกิดข้อผิดพลาด: {e}")
