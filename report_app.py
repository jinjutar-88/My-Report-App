import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
import io
import gspread
from google.oauth2.service_account import Credentials

# --- 1. ตั้งค่าพื้นฐาน ---
GOOGLE_SHEET_NAME = "Smart Dev Report Log"

# --- 2. ฟังก์ชันจัดการรูปภาพให้พอดีช่อง ---
def add_image_to_excel(ws, img_file, cell_address):
    if img_file is None: return
    temp_path = f"temp_{cell_address}.png"
    with open(temp_path, "wb") as f:
        f.write(img_file.getbuffer())
    img = Image(temp_path)
    
    # ตรวจสอบพื้นที่ (รองรับช่องที่ Merge)
    target_width, target_height = 0, 0
    for m_range in ws.merged_cells.ranges:
        if cell_address in m_range:
            for col in range(m_range.min_col, m_range.max_col + 1):
                col_letter = ws.cell(row=1, column=col).column_letter
                target_width += (ws.column_dimensions[col_letter].width or 8.43) * 7.5
            for row in range(m_range.min_row, m_range.max_row + 1):
                target_height += (ws.row_dimensions[row].height or 15) * 1.33
            img.width, img.height = target_width - 10, target_height - 10
            ws.add_image(img, cell_address)
            return
    # กรณีช่องปกติ
    img.width, img.height = 300, 200 # ขนาด Default หากไม่พบช่อง Merge
    ws.add_image(img, cell_address)

# --- 3. ระบบจัดการจำนวนรูปภาพ (เพิ่ม/ลบ) ---
if 'num_photos' not in st.session_state:
    st.session_state.num_photos = 1

def add_photo(): st.session_state.num_photos += 1
def remove_photo(): 
    if st.session_state.num_photos > 1: st.session_state.num_photos -= 1

# --- 4. หน้าเว็บ UI ---
st.title("🚀 Smart Dev Report Generator")

# PART 1: Document Details
st.header("📄 Part 1: Document Details")
doc_no = st.text_input("Doc. No.")
ref_po_no = st.text_input("Ref. PO No.")
date_issue = st.date_input("Date of Issue", datetime.now())

# PART 2: Project & Client
st.header("🏢 Part 2: Project & Client")
project_name = st.text_input("Project Name")
site_location = st.text_input("Site / Location")
contact_client = st.text_input("Contact Person (Client)")
contact_co_ltd = st.text_input("Contact (Smart Dev Solution Co., Ltd.)")
engineer_name = st.text_input("Engineer Name (Prepared By)")

# PART 3: Service Details
st.header("🛠 Part 3: Service Details")
service_type = st.selectbox("Service Type", ["New", "Commissioning", "Repairing", "Services", "Training", "Check", "Other"])
job_performed = st.text_area("Job Performed")

# PART 4: Photo Report (Dynamic)
st.header("📸 Part 4: Photo Report")
photo_data = []
for i in range(st.session_state.num_photos):
    st.subheader(f"Photo {i+1}")
    img = st.file_uploader(f"Upload Image {i+1}", type=['png', 'jpg', 'jpeg'], key=f"img_{i}")
    desc = st.text_input(f"Description for Photo {i+1}", key=f"desc_{i}")
    photo_data.append({"img": img, "desc": desc})

col_btn1, col_btn2 = st.columns(2)
with col_btn1: st.button("➕ Add Photo", on_click=add_photo)
with col_btn2: st.button("➖ Remove Photo", on_click=remove_photo)

st.markdown("---")

if st.button("Submit & Generate Report"):
    try:
        wb = load_workbook("template.xlsx")
        ws = wb.active

        #Mapping ข้อมูลลง Excel (ตัวอย่างตำแหน่ง - โปรดปรับตามไฟล์จริง)
        ws["A5"] = f"Doc.No. : {doc_no}"
        ws["E5"] = f"Ref.PO.No. : {ref_po_no}"
        ws["I5"] = date_issue.strftime('%d/%m/%Y')
        ws["B20"] = project_name
        ws["G8"] = site_location
        ws["B10"] = contact_client
        ws["B52"] = contact_co_ltd # Contact Co.Ltd
        ws["B60"] = engineer_name

        # ตำแหน่งสำหรับวางรูป (วนลูปตามจำนวนรูปที่อัปโหลด)
        # ตัวอย่างพิกัดรูป: B58, F58, B75, F75 ...
        loc_map = ["B58", "F58", "B75", "F75", "B92", "F92"] # เพิ่มตำแหน่งรองรับรูปเยอะๆ
        desc_map = ["B70", "F70", "B87", "F87", "B104", "F104"]

        for idx, item in enumerate(photo_data):
            if item["img"] and idx < len(loc_map):
                add_image_to_excel(ws, item["img"], loc_map[idx])
                ws[desc_map[idx]] = item["desc"]

        # บันทึกและส่งออก
        excel_out = io.BytesIO()
        wb.save(excel_out)
        
        # บันทึกลง Google Sheet
        try:
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            client = gspread.authorize(creds)
            gs = client.open(GOOGLE_SHEET_NAME).sheet1
            gs.append_row([date_issue.strftime('%d/%m/%Y'), doc_no, project_name, engineer_name])
            st.success("✅ บันทึกลง Google Sheet เรียบร้อย")
        except: pass

        st.download_button("📥 Download Excel Report", excel_out.getvalue(), f"Report_{doc_no}.xlsx")
        st.balloons()

    except Exception as e:
        st.error(f"🚨 Error: {e}")
