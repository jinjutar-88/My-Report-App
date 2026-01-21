import streamlit as st
from docx import Document
from docx.shared import Inches
import io

st.set_page_config(page_title="Engineer Report Generator", layout="wide")

st.title("🛠 Engineer Service Report Generator")

# --- ส่วนที่ 1: ข้อมูลรายงานและวิศวกร ---
st.subheader("📋 General Information & Engineer Details")
col1, col2 = st.columns(2)

with col1:
    date_issue = st.date_input("Date of Issue")
    project_name = st.text_input("Project Name")
    indent_no = st.text_input("Indent No.")
    service_type = st.text_input("Service Type")
    func_since = st.date_input("Functioning Since")

with col2:
    eng_title = st.selectbox("Name Title (Engineer)", ["Mr.", "Ms.", "Mrs.", "Dr."])
    eng_name = st.text_input("Full Name (Engineer)")
    location = st.text_input("Site/Location")

st.markdown("---")

# --- ส่วนที่ 2: ข้อมูลลูกค้า ---
st.subheader("🏢 Client Information")
c_col1, c_col2 = st.columns(2)

with c_col1:
    contact_co = st.text_input("Contact (Co., Ltd.)")
    client_title = st.selectbox("Name Title (Client)", ["Mr.", "Ms.", "Mrs."])

with c_col2:
    client_name = st.text_input("Full Name (Client)")

st.markdown("---")

# --- ส่วนที่ 3: รายละเอียดงาน ---
st.subheader("📝 Work Details")
job_performed = st.text_area("Job Performed", height=150)

# --- ส่วนที่ 4: รูปภาพประกอบ ---
st.subheader("📸 Photos & Descriptions")
if 'photo_count' not in st.session_state:
    st.session_state.photo_count = 1

photos = []
for i in range(st.session_state.photo_count):
    row_col1, row_col2 = st.columns([1, 2])
    with row_col1:
        img_file = st.file_uploader(f"Upload Image {i+1}", type=['jpg', 'png', 'jpeg'], key=f"img_{i}")
        if img_file:
            st.image(img_file, width=200)
    with row_col2:
        img_desc = st.text_input(f"Description for Image {i+1}", key=f"desc_{i}")
    photos.append({"file": img_file, "desc": img_desc})

if st.button("➕ Add More Photo"):
    st.session_state.photo_count += 1
    st.rerun()

st.markdown("---")

# --- ส่วนการสร้างไฟล์ Word ---
if st.button("🚀 Generate Report"):
    if not project_name:
        st.error("กรุณากรอกชื่อ Project Name ก่อนสร้างรายงานครับ")
    else:
        doc = Document()
        doc.add_heading('SERVICE REPORT', 0)

        # ข้อมูลทั่วไป
        p1 = doc.add_paragraph()
        p1.add_run(f"Date of Issue: ").bold = True
        p1.add_run(f"{date_issue}\n")
        p1.add_run(f"Project Name: ").bold = True
        p1.add_run(f"{project_name}\n")
        p1.add_run(f"Indent No.: ").bold = True
        p1.add_run(f"{indent_no}\n")
        p1.add_run(f"Service Type: ").bold = True
        p1.add_run(f"{service_type}\n")
        p1.add_run(f"Functioning Since: ").bold = True
        p1.add_run(f"{func_since}\n")

        # ข้อมูลวิศวกร
        p2 = doc.add_paragraph()
        p2.add_run(f"Engineer: ").bold = True
        p2.add_run(f"{eng_title} {eng_name}\n")
        p2.add_run(f"Location: ").bold = True
        p2.add_run(f"{location}")

        # ข้อมูลลูกค้า
        doc.add_heading('Client Details', level=1)
        p3 = doc.add_paragraph()
        p3.add_run(f"Company: ").bold = True
        p3.add_run(f"{contact_co}\n")
        p3.add_run(f"Contact Person: ").bold = True
        p3.add_run(f"{client_title} {client_name}")

        # รายละเอียดงาน
        doc.add_heading('Job Performed', level=1)
        doc.add_paragraph(job_performed)

        # รูปภาพ
        if any(p['file'] for p in photos):
            doc.add_heading('Photos', level=1)
            for p_item in photos:
                if p_item["file"]:
                    doc.add_picture(p_item["file"], width=Inches(4))
                    doc.add_paragraph(p_item["desc"])

        # ส่งไฟล์ให้ Download
        bio = io.BytesIO()
        doc.save(bio)
        st.download_button(
            label="📥 Download Word Report",
            data=bio.getvalue(),
            file_name=f"Service_Report_{project_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
