import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
import io
import gspread
from google.oauth2.service_account import Credentials

# --- 1. ตั้งค่าพื้นฐาน ---
GOOGLE_SHEET_NAME = "Smart Dev Report Log"

# --- 2. ฟังก์ชันจัดการรูปภาพให้ลงล็อคช่อง Excel ---
def add_image_to_excel(ws, img_file, cell_address):
    if img_file is None:
        return
    
    # สร้างไฟล์ชั่วคราวสำหรับรูป
    temp_path = f"temp_{cell_address}.png"
    with open(temp_path, "wb") as f:
        f.write(img_file.getbuffer())
        
    img = Image(temp_path)
    
    # คำนวณขนาดพื้นที่จากช่อง (รองรับช่องที่ถูก Merge ไว้)
    target_width = 0
    target_height = 0
    found_merge = False
    
    for m_range in ws.merged_cells.ranges:
        if cell_address in m_range:
            for col in range(m_range.min_col, m_range.max_col + 1):
                col_letter = ws.cell(row=1, column=col).column_letter
                target_width += (ws.column_dimensions[col_letter].width or 8.43) * 7.5
            for row in range(m_range.min_row, m_range.max_row + 1):
                target_height += (ws.row_dimensions[row].height or 15) * 1.33
            found_merge = True
            break
            
    if not found_merge:
        col_letter = cell_address[0]
        row_num = int(''.join(filter(str.isdigit, cell_address)))
        target_width = (ws.column_dimensions[col_letter].width or 8.43) * 7.5
        target_height = (ws.row_dimensions[row_num].height or 15) * 1.33

    # ปรับขนาดรูปให้เล็กกว่าช่องนิดหน่อย (Padding) เพื่อความสวยงาม
    img.width = target_width - 10
    img.height = target_height - 10
    ws.add_image(img, cell_address)

# --- 3. หน้าจอ UI แบบเก่า (เรียงลงมาตรงๆ) ---
st.title("🚀 Smart Dev Report Generator")

# ส่วนกรอกข้อมูลแบบเรียงลำดับ
project_name = st.text_input("Project Name")
location = st.text_input("Location")
eng_name = st.text_input("Engineer Name")
date_issue = st.date_input("Date of Issue", datetime.now())

st.markdown("---")
st.subheader("📸 Photo Report")

# ส่วนอัปโหลดรูปแบบเรียงลงมา (แบบเก่า)
img1 = st.file_uploader("Upload Photo 1", type=['png', 'jpg', 'jpeg'], key="1")
img2 = st.file_uploader("Upload Photo 2", type=['png', 'jpg', 'jpeg'], key="2")
img3 = st.file_uploader("Upload Photo 3", type=['png', 'jpg', 'jpeg'], key="3")
img4 = st.file_uploader("Upload Photo 4", type=['png', 'jpg', 'jpeg'], key="4")

# รวมตัวแปรไว้ใน List (ประกาศหลังจากมี fule_uploader เพื่อกัน NameError)
uploaded_imgs = [img1, img2, img3, img4]

st.markdown("---")

if st.button("Submit & Generate Report"):
    try:
        # 1. โหลด Template
        wb = load_workbook("template.xlsx")
        ws = wb.active

        # 2. ใส่ข้อมูล Text (ตำแหน่งสมมติ ปรับตามไฟล์จริงของคุณ)
        ws["B12"] = project_name
        ws["B13"] = location
        ws["I5"] = date_issue.strftime('%d/%m/%Y')

        # 3. ใส่รูปภาพตามตำแหน่งที่คำนวณไว้ในไฟล์ template
        # ตำแหน่งช่องใน Excel ที่คุณเตรียมไว้ (เช่น B58, F58, B75, F75)
        photo_locations = ["B58", "F58", "B75", "F75"]
        
        for loc, img_file in zip(photo_locations, uploaded_imgs):
            if img_file:
                add_image_to_excel(ws, img_file, loc)

        # 4. เตรียมไฟล์สำหรับดาวน์โหลด
        excel_out = io.BytesIO()
        wb.save(excel_out)
        excel_bytes = excel_out.getvalue()

        # 5. บันทึกลง Google Sheet
        try:
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            client = gspread.authorize(creds)
            gs = client.open(GOOGLE_SHEET_NAME).sheet1
            
            row = [date_issue.strftime('%d/%m/%Y'), project_name, location, eng_name, datetime.now().strftime('%H:%M:%S')]
            gs.append_row(row)
            st.success("✅ บันทึกลง Google Sheet สำเร็จ")
        except Exception as e:
            if "200" in str(e):
                st.success("✅ บันทึกลง Google Sheet สำเร็จ (200)")
            else:
                st.warning(f"⚠️ Google Sheet Error: {e}")

        # 6. ปุ่มดาวน์โหลด
        st.download_button("📥 Download Excel Report", excel_bytes, f"Report_{project_name}.xlsx")
        st.balloons()

    except Exception as e:
        st.error(f"🚨 เกิดข้อผิดพลาด: {e}")
