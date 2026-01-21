import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
import io
import gspread
from google.oauth2.service_account import Credentials

# --- 1. การตั้งค่าพื้นฐาน ---
GOOGLE_SHEET_NAME = "Smart Dev Report Log"

# --- 2. ฟังก์ชันจัดการรูปภาพให้พอดีช่อง Excel ---
def add_image_to_excel(ws, img_file, cell_address):
    if img_file is None:
        return
    temp_path = f"temp_{cell_address}.png"
    with open(temp_path, "wb") as f:
        f.write(img_file.getbuffer())
    img = Image(temp_path)
    
    # คำนวณขนาด (รองรับช่องที่ Merge)
    target_width = 0
    target_height = 0
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

    # ถ้าไม่ Merge ใช้ขนาดปกติ
    col_letter = cell_address[0]
    row_num = int(''.join(filter(str.isdigit, cell_address)))
    img.width = (ws.column_dimensions[col_letter].width or 8.43) * 7.5 - 10
    img.height = (ws.row_dimensions[row_num].height or 15) * 1.33 - 10
    ws.add_image(img, cell_address)

# --- 3. หน้าเว็บ (UI) แบ่งเป็น Part ตามแบบเดิม ---
st.title("🚀 Smart Dev Report Generator")

# --- PART 1: General Information ---
st.header("📋 Part 1: General Information")
date_issue = st.date_input("Date of Issue", datetime.now())
project_name = st.text_input("Project Name")
site_location = st.text_input("Site / Location")
engineer_name = st.text_input("Engineer Name")

# --- PART 2: Contact Details ---
st.header("👤 Part 2: Contact Details")
contact_client = st.text_input("Contact Person (Client)")
contact_co_ltd = st.text_input("Contact (Smart Dev Solution Co., Ltd.)")

# --- PART 3: Service Type & Job Performed ---
st.header("🛠 Part 3: Service Details")
service_type = st.selectbox("Service Type", [
    "New", "Commissioning", "Repairing", "Services", "Training", "Check", "Other"
])
job_performed = st.text_area("Job Performed", height=100)

# --- PART 4: Photo Report & Description ---
st.header("📸 Part 4: Photo & Description")

# รูปที่ 1
st.subheader("Photo 1")
img1 = st.file_uploader("Upload Image 1", type=['png', 'jpg', 'jpeg'], key="img1")
desc1 = st.text_input("Description 1", key="desc1")

# รูปที่ 2
st.subheader("Photo 2")
img2 = st.file_uploader("Upload Image 2", type=['png', 'jpg', 'jpeg'], key="img2")
desc2 = st.text_input("Description 2", key="desc2")

# รูปที่ 3
st.subheader("Photo 3")
img3 = st.file_uploader("Upload Image 3", type=['png', 'jpg', 'jpeg'], key="img3")
desc3 = st.text_input("Description 3", key="desc3")

# รูปที่ 4
st.subheader("Photo 4")
img4 = st.file_uploader("Upload Image 4", type=['png', 'jpg', 'jpeg'], key="img4")
desc4 = st.text_input("Description 4", key="desc4")

st.markdown("---")

# --- ส่วนประมวลผลเมื่อกด Submit ---
if st.button("Generate & Save Report"):
    try:
        # 1. โหลด Template
        wb = load_workbook("template.xlsx")
        ws = wb.active

        # 2. เขียนข้อมูลลง Excel (อิงตามตำแหน่งใน Template ของคุณ)
        ws["I5"] = date_issue.strftime('%d/%m/%Y')
        ws["B20"] = project_name
        ws["G8"] = site_location
        ws["B60"] = engineer_name  # สมมติจุดเซ็นชื่อ Prepared By
        ws["B10"] = contact_client
        ws["B52"] = contact_co_ltd
        ws["D14"] = service_type
        ws["B21"] = job_performed

        # 3. จัดการรูปและคำบรรยาย (ใส่ตามพิกัดที่ระบุไว้ใน Template)
        # ตัวอย่าง: รูปวางช่อง B58, คำบรรยายวางช่อง B70
        photo_configs = [
            (img1, "B58", desc1, "B70"),
            (img2, "F58", desc2, "F70"),
            (img3, "B75", desc3, "B87"),
            (img4, "F75", desc4, "F87")
        ]

        for img_file, img_loc, desc_text, desc_loc in photo_configs:
            if img_file:
                add_image_to_excel(ws, img_file, img_loc)
            if desc_text:
                ws[desc_loc] = desc_text

        # 4. เตรียมไฟล์ดาวน์โหลด
        excel_out = io.BytesIO()
        wb.save(excel_out)
        excel_bytes = excel_out.getvalue()

        # 5. บันทึกลง Google Sheet
        try:
            scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
            creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
            client = gspread.authorize(creds)
            gs = client.open(GOOGLE_SHEET_NAME).sheet1
            gs.append_row([
                date_issue.strftime('%d/%m/%Y'), project_name, site_location, 
                engineer_name, service_type, datetime.now().strftime('%H:%M:%S')
            ])
            st.success("✅ บันทึกลง Google Sheet สำเร็จ")
        except Exception as e:
            if "200" in str(e): st.success("✅ บันทึกลง Google Sheet สำเร็จ (200)")
            else: st.warning(f"⚠️ Sheet Error: {e}")

        # 6. ปุ่มดาวน์โหลด
        st.download_button("📥 Download Excel", excel_bytes, f"Report_{project_name}.xlsx")
        st.balloons()

    except Exception as e:
        st.error(f"🚨 ข้อผิดพลาด: {e}")
