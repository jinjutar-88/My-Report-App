import streamlit as st
import openpyxl
import io
from datetime import datetime

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Engineer Report Generator", layout="wide")
st.title("🛠 Smart Dev Solution - Service Report")

# --- ส่วนของ Session State สำหรับจัดการจำนวนรูปภาพ ---
if 'photo_count' not in st.session_state:
    st.session_state.photo_count = 1 # เริ่มต้นที่ 1 รูป

def add_photo():
    st.session_state.photo_count += 1

# --- PART 1: ข้อมูลทั่วไป ---
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

# --- PART 3: รูปภาพและคำบรรยาย (แบบเพิ่มได้) ---
st.markdown("---")
st.subheader("📸 Part 3: Photo Report")

photos = []
for i in range(st.session_state.photo_count):
    st.write(f"**Photo {i+1}**")
    col_img, col_txt = st.columns([1, 1])
    with col_img:
        up_file = st.file_uploader(f"Upload Photo {i+1}", type=['jpg', 'jpeg', 'png'], key=f"file_{i}")
        if up_file:
            st.image(up_file, width=250)
    with col_txt:
        desc = st.text_area(f"Description for Photo {i+1}", key=f"desc_{i}", height=100)
    photos.append({"file": up_file, "desc": desc})
    st.markdown("---")

# ปุ่มกดเพื่อเพิ่มช่องรูปภาพ
st.button("➕ Add More Photo", on_click=add_photo)

# --- ปุ่มสร้างไฟล์ Excel ---
st.write(" ")
if st.button("🚀 Generate Excel Report", use_container_width=True):
    try:
        wb = openpyxl.load_workbook("template.xlsx")
        sheet = wb.active 

        # เติมข้อมูลลงใน Cell
        sheet["J5"] = date_issue.strftime('%d/%m/%Y')
        sheet["H7"] = location
        sheet["C9"] = client_name
        sheet["B16"] = project_name
        sheet["D17"] = job_performed
        
        # ตัวอย่างการวนลูปเติมคำบรรยายรูปภาพลงใน Excel
        # for idx, p in enumerate(photos):
        #    if p["desc"]:
        #        sheet[f"A{35 + (idx*5)}"] = p["desc"] # สมมติว่าวางห่างกันทุก 5 บรรทัด

        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)

        st.success(f"🎉 บันทึกข้อมูลพร้อมรูปภาพ {st.session_state.photo_count} ชุด เรียบร้อยแล้ว!")
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_data,
            file_name=f"Report_{project_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
