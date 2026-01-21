import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import io
from datetime import datetime

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Engineer Report Generator", layout="wide")
st.title("🛠 Smart Dev Solution - Service Report")

# --- PART 1: ข้อมูลทั่วไป (ตัด Ref. QT No. ออกแล้ว) ---
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
    service_type = st.selectbox("Service Type", ["New", "Repairing", "Services", "Training", "Check", "Others"])

eng_name = st.text_input("Engineer Name (Prepared By)")

# --- PART 2: รายละเอียดงาน ---
st.markdown("---")
st.subheader("🔧 Part 2: Service Details")
job_performed = st.text_area("Job Performed (รายละเอียดงานที่ปฏิบัติ)", height=150)
note = st.text_area("Note (หมายเหตุเพิ่มเติม)")

# --- PART 3: รูปภาพและคำบรรยาย ---
st.markdown("---")
st.subheader("📸 Part 3: Photo Report")
col_img, col_txt = st.columns([1, 1])

with col_img:
    uploaded_photo = st.file_uploader("Upload Photo (อัปโหลดรูปภาพ)", type=['jpg', 'jpeg', 'png'])
    if uploaded_photo:
        st.image(uploaded_photo, caption="รูปภาพที่เลือก", width=350)

with col_txt:
    photo_description = st.text_area("Photo Description (คำบรรยายรูปภาพ)", height=200, placeholder="พิมพ์คำบรรยายใต้รูปภาพที่นี่...")

# --- ปุ่มสร้างไฟล์ Excel ---
st.markdown("---")
if st.button("🚀 Generate Excel Report", use_container_width=True):
    try:
        # 1. โหลดเทมเพลต (ต้องชื่อ template.xlsx อยู่ใน GitHub)
        wb = openpyxl.load_workbook("template.xlsx")
        sheet = wb.active 

        # 2. เติมข้อมูลลงใน Cell (ตำแหน่งเดิมที่คุณเคยระบุไว้)
        sheet["J5"] = date_issue.strftime('%d/%m/%Y') # Date
        sheet["H7"] = location                        # Site/Location
        sheet["C9"] = client_name                     # Contact Person
        sheet["B16"] = project_name                    # Project
        sheet["D17"] = job_performed                   # Job Performed
        
        # ถ้าต้องการให้ข้อมูลอื่นลงช่องไหน เพิ่มได้ที่นี่ เช่น:
        # sheet["F25"] = eng_name

        # 3. เตรียมไฟล์สำหรับการดาวน์โหลด
        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)

        st.success("🎉 บันทึกข้อมูลเรียบร้อยแล้ว!")
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_data,
            file_name=f"Report_{project_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
