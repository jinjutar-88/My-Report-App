import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as XLImage
import io
from datetime import datetime

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Engineer Report Generator", layout="wide")
st.title("🛠 Smart Dev Solution - Service Report")

# ใช้ Tabs เพื่อแบ่งส่วนการกรอกข้อมูลให้ดูง่ายเหมือนเว็บแรก
tab1, tab2 = st.tabs(["📄 General Info & Job Detail", "📸 Photo Report"])

with tab1:
    st.subheader("Part 1: General Information")
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
    
    st.markdown("---")
    st.subheader("Part 2: Service Details")
    service_type = st.selectbox("Service Type", ["New", "Repairing", "Services", "Training", "Check", "Others"])
    job_performed = st.text_area("Job Performed (รายละเอียดงาน)", height=150)
    note = st.text_area("Note (หมายเหตุ)")
    eng_name = st.text_input("Engineer Name (Prepared By)")

with tab2:
    st.subheader("Part 3: Photo Report")
    st.write("อัปโหลดรูปภาพและใส่คำบรรยายประกอบงาน")
    
    # ส่วนของรูปภาพเพียง 1 ชุดตามที่ต้องการ
    col_img, col_txt = st.columns([1, 1])
    with col_img:
        uploaded_photo = st.file_uploader("Upload Photo", type=['jpg', 'jpeg', 'png'])
        if uploaded_photo:
            st.image(uploaded_photo, caption="Preview", width=300)
            
    with col_txt:
        photo_description = st.text_area("Photo Description (คำบรรยายรูปภาพ)", height=150, placeholder="พิมพ์รายละเอียดรูปภาพที่นี่...")

# --- ปุ่มสร้างไฟล์ Excel ---
st.markdown("---")
if st.button("🚀 Generate Excel Report"):
    try:
        # 1. โหลดเทมเพลต (ต้องชื่อ template.xlsx ใน GitHub)
        wb = openpyxl.load_workbook("template.xlsx")
        sheet = wb.active 

        # 2. เติมข้อมูลลงใน Cell ตามตำแหน่งที่คุณระบุ
        sheet["J5"] = date_issue.strftime('%d/%m/%Y')
        sheet["H7"] = location
        sheet["C9"] = client_name
        sheet["B16"] = project_name
        sheet["D17"] = job_performed
        
        # ตัวอย่างการเติม Description ของรูปภาพลงในหน้า 2 ของ Excel (สมมติช่อง A35)
        # sheet["A35"] = photo_description

        # 3. เตรียมไฟล์ดาวน์โหลด
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
