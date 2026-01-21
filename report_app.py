import streamlit as st
import openpyxl
import io
from datetime import datetime

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Engineer Report Generator", layout="wide")
st.title("🛠 Smart Dev Solution - Service Report")

# --- ส่วนที่ 1: รับข้อมูลจากหน้าเว็บ ---
col1, col2 = st.columns(2)
with col1:
    date_issue = st.date_input("Date of Issue")
    project_name = st.text_input("Project Name")
    location = st.text_input("Site/Location")
with col2:
    client_name = st.text_input("Contact Person (Client)")
    eng_name = st.text_input("Engineer Name")

job_performed = st.text_area("Job Performed")

# --- ส่วนที่ 2: ปุ่มสร้างไฟล์ Excel ---
if st.button("🚀 Generate Excel Report"):
    try:
        # 1. โหลดเทมเพลต Excel (ต้องชื่อ template.xlsx อยู่ใน GitHub)
        wb = openpyxl.load_workbook("template.xlsx")
        sheet = wb.active 

        # 2. เติมข้อมูลตามตำแหน่ง Cell ที่ระบุ
        sheet["J5"] = date_issue.strftime('%d/%m/%Y')
        sheet["H7"] = location
        sheet["C9"] = client_name
        sheet["B16"] = project_name
        sheet["D17"] = job_performed

        # 3. เตรียมไฟล์สำหรับการดาวน์โหลด
        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)

        st.success("🎉 บันทึกข้อมูลลงในฟอร์มสำเร็จ!")
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_data,
            file_name=f"Service_Report_{project_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
