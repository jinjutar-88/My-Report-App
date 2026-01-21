import streamlit as st
import openpyxl
import io
from datetime import datetime

# ตั้งค่าหน้าเว็บให้กว้าง
st.set_page_config(page_title="Engineer Report Generator", layout="wide")
st.title("🛠 Smart Dev Solution - Service Report")

# --- PART 1: ข้อมูลทั่วไป (Header Information) ---
st.subheader("📋 Part 1: General Information")
col1, col2 = st.columns(2)

with col1:
    date_issue = st.date_input("Date of Issue")
    ref_qt_no = st.text_input("Ref. QT No.")
    ref_po_no = st.text_input("Ref. PO No.")
    project_name = st.text_input("Project Name")

with col2:
    doc_no = st.text_input("Doc. No.")
    location = st.text_input("Site / Location")
    client_name = st.text_input("Contact Person (Client)")
    contact_co_ltd = st.text_input("Contact (Co., Ltd.)")

# --- PART 2: ประเภทบริการและรายละเอียด (Service Details) ---
st.markdown("---")
st.subheader("🔧 Part 2: Service Details")
service_type = st.radio(
    "Service Type", 
    ["New", "Repairing", "Services", "Training", "Check", "Others"],
    horizontal=True
)

job_performed = st.text_area("Job Performed (รายละเอียดงานที่ปฏิบัติ)", height=150)
note = st.text_area("Note (หมายเหตุเพิ่มเติม)")
eng_name = st.text_input("Engineer Name (Prepared By)")

# --- PART 3: รูปภาพและคำบรรยาย (Photo & Description) ---
st.markdown("---")
st.subheader("📸 Part 3: Photo Report")
uploaded_images = st.file_uploader("Upload Photos", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])

photo_descriptions = []
if uploaded_images:
    cols = st.columns(2) # แบ่งแสดงรูปภาพเป็น 2 คอลัมน์ในหน้าเว็บ
    for i, img in enumerate(uploaded_images):
        with cols[i % 2]:
            st.image(img, width=300)
            desc = st.text_input(f"Description for Photo {i+1}", key=f"desc_{i}")
            photo_descriptions.append(desc)

# --- ส่วนการสร้างไฟล์ Excel ---
st.markdown("---")
if st.button("🚀 Generate Excel Report"):
    try:
        # 1. โหลดเทมเพลต Excel
        wb = openpyxl.load_workbook("template.xlsx")
        sheet = wb.active 

        # 2. เติมข้อมูลลงใน Cell ตามตำแหน่งที่คุณระบุ
        sheet["J5"] = date_issue.strftime('%d/%m/%Y')
        sheet["H7"] = location
        sheet["C9"] = client_name
        sheet["B16"] = project_name
        sheet["D17"] = job_performed
        
        # ตำแหน่งเพิ่มเติม (ปรับเลข Cell ตามไฟล์จริงของคุณ)
        # sheet["C7"] = ref_qt_no
        # sheet["F25"] = eng_name
        # sheet["D18"] = note

        # หมายเหตุ: การแทรกรูปภาพลง Excel โดยอัตโนมัติต้องใช้โค้ดเพิ่ม 
        # หากเน้นเก็บข้อมูลข้อความก่อน โค้ดนี้จะทำงานได้ทันทีครับ

        # 3. เตรียมไฟล์ดาวน์โหลด
        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)

        st.success("🎉 บันทึกข้อมูลลงฟอร์มเรียบร้อย!")
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_data,
            file_name=f"Service_Report_{project_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
