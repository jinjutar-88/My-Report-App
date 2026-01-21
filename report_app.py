import streamlit as st
import openpyxl
import io
import uuid # เพิ่มสำหรับสร้าง ID ให้แต่ละช่อง
from datetime import datetime

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Engineer Report Generator", layout="wide")
st.title("🛠 Smart Dev Solution - Service Report")

# --- ส่วนของ Session State สำหรับจัดการลิสต์ของรูปภาพ ---
if 'photo_ids' not in st.session_state:
    # เริ่มต้นด้วย 1 ช่องรูปภาพ โดยให้ ID สุ่มมา 1 ตัว
    st.session_state.photo_ids = [str(uuid.uuid4())]

def add_photo_callback():
    st.session_state.photo_ids.append(str(uuid.uuid4()))

def remove_photo_callback(id_to_remove):
    # ลบ ID ที่ระบุออกจากลิสต์ (แต่ต้องเหลือไว้อย่างน้อย 1 ช่อง)
    if len(st.session_state.photo_ids) > 1:
        st.session_state.photo_ids.remove(id_to_remove)

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
job_performed = st.text_area("Job Performed", height=150)
note = st.text_area("Note")

# --- PART 3: รูปภาพและคำบรรยาย (แบบลบแยกตามช่อง) ---
st.markdown("---")
st.subheader("📸 Part 3: Photo Report")

photos_data = []

# วนลูปตาม ID ที่มีในลิสต์
for i, photo_id in enumerate(st.session_state.photo_ids):
    # สร้าง Container สำหรับแต่ละช่องรูปภาพ
    with st.container():
        # แถวหัวข้อและปุ่มลบ
        head_col1, head_col2 = st.columns([10, 1])
        with head_col1:
            st.write(f"**Photo {i+1}**")
        with head_col2:
            # ปุ่มลบเฉพาะช่องนี้ (แสดงเมื่อมีมากกว่า 1 ช่อง)
            if len(st.session_state.photo_ids) > 1:
                st.button("🗑️", key=f"del_{photo_id}", on_click=remove_photo_callback, args=(photo_id,))
        
        # ส่วนอัปโหลดและคำบรรยาย
        col_img, col_txt = st.columns([1, 1])
        with col_img:
            up_file = st.file_uploader(f"Upload Photo {i+1}", type=['jpg','jpeg','png'], key=f"file_{photo_id}")
            if up_file:
                st.image(up_file, width=250)
        with col_txt:
            desc = st.text_area(f"Description for Photo {i+1}", key=f"desc_{photo_id}", height=120)
        
        photos_data.append({"file": up_file, "desc": desc})
        st.markdown("---")

# ปุ่มเพิ่มรูปภาพ
st.button("➕ Add More Photo", on_click=add_photo_callback)

# --- ปุ่มสร้างไฟล์ Excel ---
if st.button("🚀 Generate Excel Report", use_container_width=True):
    try:
        wb = openpyxl.load_workbook("template.xlsx")
        sheet = wb.active 
        
        sheet["J5"] = date_issue.strftime('%d/%m/%Y')
        sheet["H7"] = location
        sheet["C9"] = client_name
        sheet["B16"] = project_name
        sheet["D17"] = job_performed

        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)

        st.success(f"🎉 บันทึกสำเร็จ! ทั้งหมด {len(st.session_state.photo_ids)} ช่อง")
        st.download_button("📥 Download Excel Report", excel_data, f"Report_{project_name}.xlsx")
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
