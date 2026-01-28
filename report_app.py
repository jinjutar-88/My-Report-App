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
    img_data = io.BytesIO(img_file.getvalue())
    img = Image(img_data)
    
    max_w, max_h = 0, 0
    found_range = None
    for m_range in ws.merged_cells.ranges:
        if cell_address in m_range:
            found_range = m_range
            for col in range(m_range.min_col, m_range.max_col + 1):
                max_w += (ws.column_dimensions[get_column_letter(col)].width or 8.43) * 7.5
            for row in range(m_range.min_row, m_range.max_row + 1):
                max_h += (ws.row_dimensions[row].height or 15) * 1.33
            break
    
    if not found_range: max_w, max_h = 350, 250
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

st.title("🚀 Smart Dev Report Generator v0.5")

# --- ส่วนรับข้อมูลที่เคยหายไป ---
with st.expander("📄 Part 1-3: Document & Project Details", expanded=True):
    c1, c2, c3 = st.columns(3)
    doc_no = c1.text_input("Doc. No.")
    ref_po = c2.text_input("Ref. PO No.")
    date_issue = c3.date_input("Date", datetime.now())

    p1, p2 = st.columns(2)
    project_name = p1.text_input("Project Name")
    site_location = p1.text_input("Site / Location")
    contact_client = p2.text_input("Contact Person (Client)")
    contact_co_ltd = p2.text_input("Contact (Smart Dev Co., Ltd.)")
    engineer_name = st.text_input("Engineer Name (Prepared By)")

    job_performed = st.text_area("Job Performed (รายละเอียดงาน)")

# --- Photo Report ---
st.header("📸 Photo Report")
final_photo_data = []
for i in list(st.session_state.photos):
    with st.container():
        c_p, c_i, c_d = st.columns([3, 5, 1])
        with c_i:
            f = st.file_uploader(f"รูปที่ {i+1}", type=['jpg','png','jpeg'], key=f"f{i}")
            d = st.text_input(f"คำบรรยาย {i+1}", key=f"d{i}")
        with c_p:
            if f: st.image(f, use_container_width=True)
        with c_d:
            if st.button("🗑️", key=f"del{i}"):
                st.session_state.photos.remove(i)
                st.rerun()
        final_photo_data.append({"img": f, "desc": d})

if st.button("➕ เพิ่มรูปภาพ"):
    st.session_state.photos.append(max(st.session_state.photos) + 1 if st.session_state.photos else 0)
    st.rerun()

# --- 4. ENGINE ---
if st.button("🚀 Generate & Send Report", type="primary"):
    try:
        wb = load_workbook("template.xlsx")
        ws = wb.active # Sheet1
        ws_temp = wb["ImageTemplate"]

        # --- เขียนข้อมูล Part 1-3 ลงหน้าแรก ---
        write_safe(ws, "B5", doc_no)
        write_safe(ws, "F6", ref_po)
        write_safe(ws, "J5", date_issue.strftime('%d/%m/%Y'))
        write_safe(ws, "B16", project_name)
        write_safe(ws, "H7", site_location)
        write_safe(ws, "C9", contact_client)
        write_safe(ws, "A7", contact_co_ltd)
        write_safe(ws, "B42", engineer_name)
        write_safe(ws, "D17", job_performed) 

        # พิกัดรูปภาพ 1-6
        loc_fixed = ["A49", "A62", "A75", "A92", "A105", "A118"]
        desc_fixed = ["H49", "H62", "H75", "H92", "H105", "H118"]
        
        # --- Logic รูปที่ 7 เป็นต้นไป (Pagination) ---
        start_row = 131 
        header_h = 4
        block_h = 13
        gap_h = 4 # เว้น 4 ช่องก่อนเริ่มหน้าใหม่

        for idx, item in enumerate(final_photo_data):
            if not item["img"]: continue
            
            if idx < 6:
                p_loc, d_loc = loc_fixed[idx], desc_fixed[idx]
            else:
                rel_idx = idx - 6
                page_num = rel_idx // 3
                pos_in_page = rel_idx % 3 # 0, 1, 2
                
                # คำนวณ Page Offset: 1 หน้าประกอบด้วย Header(4) + 3 Blocks(39) + Gap(4)
                # รวมเป็น 47 แถวต่อหน้า
                page_offset = page_num * (header_h + (3 * block_h) + gap_h)
                
                # แทรก Header ใหม่ทุก 3 รูป
                if pos_in_page == 0:
                    h_start = start_row + page_offset
                    for r in range(1, header_h + 1):
                        ws.row_dimensions[h_start + r - 1].height = ws_temp.row_dimensions[r].height
                        for c in range(1, 12):
                            copy_style(ws_temp.cell(row=r, column=c), ws.cell(row=h_start + r - 1, column=c))
                
                # พิกัดวางรูป
                p_row = start_row + page_offset + header_h + (pos_in_page * block_h)
                
                # Copy บล็อกรูป (จาก ImageTemplate แถว 5-17)
                for r in range(0, block_h):
                    target_r = p_row + r
                    ws.row_dimensions[target_r].height = ws_temp.row_dimensions[5 + r].height
                    for c in range(1, 12):
                        copy_style(ws_temp.cell(row=5 + r, column=c), ws.cell(row=target_r, column=c))
                
                # จัดการ Merged Cells
                for m_range in ws_temp.merged_cells.ranges:
                    if m_range.min_row >= 5 and m_range.max_row <= 17:
                        t_o, b_o = m_range.min_row - 5, m_range.max_row - 5
                        new_m = f"{get_column_letter(m_range.min_col)}{p_row + t_o}:{get_column_letter(m_range.max_col)}{p_row + b_o}"
                        if new_m not in ws.merged_cells: ws.merge_cells(new_m)
                
                p_loc, d_loc = f"A{p_row}", f"H{p_row}"

            add_image_to_excel(ws, item["img"], p_loc)
            write_safe(ws, d_loc, item["desc"])

        # บันทึกและส่งเมล
        output = io.BytesIO()
        wb.save(output)
        msg = MIMEMultipart()
        msg['From'], msg['To'], msg['Subject'] = SENDER_EMAIL, RECEIVER_EMAIL, f"Report: {doc_no}"
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(output.getvalue())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="Report_{doc_no}.xlsx"')
        msg.attach(part)
        
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
            
        st.success("✅ สร้างและส่งรายงานสำเร็จ!")
        st.download_button("📥 Download Excel", output.getvalue(), f"Report_{doc_no}.xlsx")

    except Exception as e:
        st.error(f"🚨 ข้อผิดพลาด: {e}")
