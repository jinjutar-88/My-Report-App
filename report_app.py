import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from datetime import datetime
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from copy import copy

# --- 1. CONFIGURATION ---
SENDER_EMAIL = "jinjutar.smartdev@gmail.com"
SENDER_PASSWORD = "uzfs bdtc xclz rzsq"
RECEIVER_EMAIL = "jinjutar.smartdev@gmail.com"

# --- 2. HELPERS ---
def copy_style(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def add_image_to_excel(ws, img_file, cell_address):
    if img_file is None: return
    # ใช้ BytesIO แทนไฟล์จริงเพื่อป้องกันปัญหาไฟล์ซ้ำ
    img_data = io.BytesIO(img_file.getvalue())
    img = Image(img_data)
    
    max_w, max_h = 0, 0
    found_range = None
    for m_range in ws.merged_cells.ranges:
        if cell_address in m_range:
            found_range = m_range
            for col in range(m_range.min_col, m_range.max_col + 1):
                col_letter = get_column_letter(col)
                max_w += (ws.column_dimensions[col_letter].width or 8.43) * 7.5
            for row in range(m_range.min_row, m_range.max_row + 1):
                max_h += (ws.row_dimensions[row].height or 15) * 1.33
            break
    
    if not found_range: max_w, max_h = 400, 300 # ปรับให้ใหญ่ขึ้นถ้าไม่เจอช่อง merge
    
    ratio = min((max_w - 10) / img.width, (max_h - 10) / img.height)
    img.width, img.height = int(img.width * ratio), int(img.height * ratio)
    ws.add_image(img, cell_address)

def write_safe(ws, cell_addr, value):
    if value is None: value = ""
    for m_range in ws.merged_cells.ranges:
        if cell_addr in m_range:
            ws.cell(row=m_range.min_row, column=m_range.min_col).value = value
            return
    ws[cell_addr] = value

# --- 3. STREAMLIT UI ---
st.set_page_config(page_title="Smart Dev Report Generator", layout="wide")
if 'photos' not in st.session_state: st.session_state.photos = [0]

st.title("🚀 Smart Dev Report Generator v0.4.2")

# --- ข้อมูลพื้นฐาน ---
with st.expander("📄 Document & Project Details", expanded=True):
    col1, col2, col3 = st.columns(3)
    doc_no = col1.text_input("Doc. No.")
    ref_po_no = col2.text_input("Ref. PO No.")
    date_issue = col3.date_input("Date of Issue", datetime.now())

    p_col1, p_col2 = st.columns(2)
    project_name = p_col1.text_input("Project Name")
    site_location = p_col1.text_input("Site / Location")
    contact_client = p_col2.text_input("Contact Person (Client)")
    contact_co_ltd = p_col2.text_input("Contact (Smart Dev Co., Ltd.)")
    engineer_name = st.text_input("Engineer Name (Prepared By)")

st.header("🛠 Service Details")
job_performed = st.text_area("Job Performed (รายละเอียดงานที่ทำ)")

# --- Photo Report ---
st.header("📸 Photo Report")
final_photo_data = []
for i in list(st.session_state.photos):
    with st.container():
        col_prev, col_input, col_del = st.columns([3, 5, 1])
        with col_input:
            up_img = st.file_uploader(f"Upload Image {i+1}", type=['jpg','png','jpeg'], key=f"f{i}")
            up_desc = st.text_input(f"Description {i+1}", key=f"d{i}")
        with col_prev:
            if up_img: st.image(up_img, use_container_width=True)
        with col_del:
            if st.button("🗑️", key=f"del{i}"):
                st.session_state.photos.remove(i)
                st.rerun()
        final_photo_data.append({"img": up_img, "desc": up_desc})
        st.markdown("---")

if st.button("➕ Add More Photo"):
    new_id = max(st.session_state.photos) + 1 if st.session_state.photos else 0
    st.session_state.photos.append(new_id)
    st.rerun()

# --- 4. GENERATE & SEND ---
if st.button("🚀 Generate & Send Report", type="primary"):
    try:
        wb = load_workbook("template.xlsx")
        ws = wb.active
        ws_temp = wb["ImageTemplate"]

        # --- เขียนข้อมูลลง Excel (เช็คพิกัดเหล่านี้ให้ตรงกับไฟล์คุณ) ---
        write_safe(ws, "B5", doc_no)
        write_safe(ws, "F6", ref_po_no)
        write_safe(ws, "J5", date_issue.strftime('%d/%m/%Y'))
        write_safe(ws, "B16", project_name)
        write_safe(ws, "H7", site_location)
        write_safe(ws, "C9", contact_client)
        write_safe(ws, "A7", contact_co_ltd)
        write_safe(ws, "B42", engineer_name)
        
        # *** สำคัญ: ลองตรวจสอบว่าในไฟล์ Excel ช่อง Job Performed คือ D17 จริงหรือไม่ ***
        write_safe(ws, "D17", job_performed) 

        # --- จัดการรูปภาพ ---
        loc_fixed = ["A49", "A62", "A75", "A92", "A105", "A118"]
        desc_fixed = ["H49", "H62", "H75", "H92", "H105", "H118"]
        
        start_gen_row, row_step, header_h, gap_h, temp_start = 131, 13, 4, 4, 5

        for idx, item in enumerate(final_photo_data):
            if not item["img"]: continue  # ข้ามถ้ารูปไม่ได้อัปโหลด
            
            if idx < 6:
                p_loc, d_loc = loc_fixed[idx], desc_fixed[idx]
            else:
                rel_idx = idx - 6
                num_pages = rel_idx // 3
                curr_row = start_gen_row + (rel_idx * row_step) + (num_pages * header_h) + (num_pages * gap_h)

                # แทรก Header ใหม่ทุก 3 รูป
                if rel_idx % 3 == 0:
                    for hr in range(1, header_h + 1):
                        target_hr = curr_row - header_h + hr - 1
                        ws.row_dimensions[target_hr].height = ws_temp.row_dimensions[hr].height
                        for c in range(1, 12):
                            copy_style(ws_temp.cell(row=hr, column=c), ws.cell(row=target_hr, column=c))

                # Copy บล็อกรูปภาพ
                for r in range(0, row_step):
                    target_r = curr_row + r
                    ws.row_dimensions[target_r].height = ws_temp.row_dimensions[temp_start + r].height
                    for c in range(1, 12):
                        copy_style(ws_temp.cell(row=temp_start + r, column=c), ws.cell(row=target_r, column=c))
                
                # Copy การ Merge Cells
                for m_range in ws_temp.merged_cells.ranges:
                    if m_range.min_row >= 5 and m_range.max_row <= 17:
                        t_o, b_o = m_range.min_row - temp_start, m_range.max_row - temp_start
                        m_min_col = get_column_letter(m_range.min_col)
                        m_max_col = get_column_letter(m_range.max_col)
                        new_m = f"{m_min_col}{curr_row + t_o}:{m_max_col}{curr_row + b_o}"
                        if new_m not in ws.merged_cells: ws.merge_cells(new_m)
                
                p_loc, d_loc = f"A{curr_row}", f"H{curr_row}"

            # วางรูปลงพิกัด
            add_image_to_excel(ws, item["img"], p_loc)
            write_safe(ws, d_loc, item["desc"])

        # บันทึกและส่งเมล
        output = io.BytesIO()
        wb.save(output)
        file_bytes = output.getvalue()

        # ส่งอีเมล
        msg = MIMEMultipart()
        msg['From'], msg['To'], msg['Subject'] = SENDER_EMAIL, RECEIVER_EMAIL, f"Report: {doc_no}"
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="Report_{doc_no}.xlsx"')
        msg.attach(part)
        
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
            
        st.success("✅ สร้างไฟล์และส่งเมลสำเร็จ!")
        st.download_button("📥 Download Excel", file_bytes, f"Report_{doc_no}.xlsx")

    except Exception as e:
        st.error(f"🚨 ข้อผิดพลาด: {e}")
