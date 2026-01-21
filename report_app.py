import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import io
from datetime import datetime
from PIL import Image

st.set_page_config(page_title="Engineer Report Generator", layout="wide")
st.title("🛠 Smart Dev Solution - Service Report")

# --- PART 1 & 2: ข้อมูลทั่วไป (เหมือนเดิม) ---
st.subheader("📋 General Information & Service Details")
col1, col2 = st.columns(2)
with col1:
    date_issue = st.date_input("Date of Issue")
    ref_qt_no = st.text_input("Ref. QT No.")
    ref_po_no = st.text_input("Ref. PO No.")
    project_name = st.text_input("Project Name")
    location = st.text_input("Site / Location")
with col2:
    doc_no = st.text_input("Doc. No.")
    client_name = st.text_input("Contact Person (Client)")
    contact_co_ltd = st.text_input("Contact (Co., Ltd.)")
    service_type = st.selectbox("Service Type", ["New", "Repairing", "Services", "Training", "Check", "Others"])
    eng_name = st.text_input("Engineer Name (Prepared By)")

job_performed = st.text_area("Job Performed (รายละเอียดงาน)")
note = st.text_area("Note (หมายเหตุ)")

# --- PART 3: PHOTO & DESCRIPTION (แบ่งเป็น 4 ส่วนตามหน้าไฟล์) ---
st.markdown("---")
st.subheader("📸 Part 3: Photo Report & Description")

photo_data = [] # เก็บข้อมูลรูปและคำบรรยาย
for i in range(1, 5): # สร้างช่องสำหรับ 4 รูป
    st.write(f"**Photo {i}**")
    col_img, col_txt = st.columns([1, 2])
    with col_img:
        up_file = st.file_uploader(f"Upload Photo {i}", type=['jpg', 'jpeg', 'png'], key=f"img_{i}")
    with col_txt:
        desc = st.text_area(f"Description for Photo {i}", key=f"desc_{i}", height=100)
    photo_data.append({"file": up_file, "desc": desc})

# --- ส่วนการสร้างไฟล์ Excel ---
if st.button("🚀 Generate Excel Report"):
    try:
        wb = openpyxl.load_workbook("template.xlsx")
        sheet = wb.active 

        # 1. เติมข้อมูล Text
        sheet["J5"] = date_issue.strftime('%d/%m/%Y')
        sheet["H7"] = location
        sheet["C9"] = client_name
        sheet["B16"] = project_name
        sheet["D17"] = job_performed
        # เพิ่มเติมตามต้องการ เช่น sheet["C7"] = ref_qt_no

        # 2. จัดการรูปภาพและคำบรรยาย (ตัวอย่างการวางในตำแหน่งต่างๆ)
        # หมายเหตุ: ตำแหน่ง Cell สำหรับรูปภาพต้องเช็คจากไฟล์ Excel ของคุณอีกครั้ง
        # สมมติ Photo 1 อยู่หน้า 2 Cell A30, Photo 2 อยู่หน้า 3...
        # ในที่นี้ผมจะเติม Description ลงไปในช่องที่เหมาะสม (สมมติช่องใต้รูป)
        
        # ตัวอย่างการเติมคำบรรยายลงใน Excel
        # sheet["B35"] = photo_data[0]["desc"] 
        # sheet["B70"] = photo_data[1]["desc"]

        # 3. เตรียมไฟล์ดาวน์โหลด
        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)

        st.success("🎉 บันทึกข้อมูลและคำบรรยายเรียบร้อย!")
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_data,
            file_name=f"Service_Report_{project_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
