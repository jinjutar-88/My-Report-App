import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from copy import copy

# --- 1. CONFIGURATION (ตั้งค่าอีเมล) ---
SENDER_EMAIL = "jinjutar.smartdev@gmail.com"
SENDER_PASSWORD = "uzfs bdtc xclz rzsq"
RECEIVER_EMAIL = "jinjutar.smartdev@gmail.com"

# --- 2. HELPERS (ฟังก์ชันช่วยจัดการ Excel) ---

def copy_style(source_cell, target_cell):
    """ก๊อปปี้รูปแบบจากเซลล์ต้นทางไปยังปลายทาง"""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def add_image_to_excel(ws, img_file, cell_address):
    """วางรูปภาพลงในเซลล์และ Resize ให้พอดีช่องที่ Merge ไว้"""
    if img_file is None: return
    temp_path = f"temp_{cell_address}.png"
    with open(temp_path, "wb") as f:
        f.write(img_file.getbuffer())
    img = Image(temp_path)
    
    max_w, max_h = 0, 0
    found_range = None
    # หาขนาดของช่องที่ Merge ไว้เพื่อคำนวณขนาดรูป
    for m_range in ws.merged_cells.ranges:
        if cell_address in m_range:
            found_range = m_range
            for col in range(m_range.min_col, m_range.max_col + 1):
                col_letter = ws.cell(row=1, column=col).column_letter
                max_w += (ws.column_dimensions[col_letter].width or 8.43) * 7.5
            for row in range(m_range.min_row, m_range.max_row + 1):
                max_h += (ws.row_dimensions[row].height or 15) * 1.33
            break
    
    if not found_range: max_w, max_h = 300, 200 # กรณีไม่ได้ Merge

    ratio = min((max_w - 10) / img.width, (max_h - 10) / img.height)
    img.width, img.height = int(img.width * ratio), int(img.height * ratio)
    ws.add_image(img, cell_address)

def write_safe(ws, cell_addr, value):
    """เขียนข้อมูลลงเซลล์ (รองรับทั้งเซลล์ปกติและ Merge Cells)"""
    for m_range in ws.merged_cells.ranges:
        if cell_addr in m_range:
            ws.cell(row=m_range.min_row, column=m_range.min_col).value = value
            return
    ws[cell_addr] = value

# --- 3. STREAMLIT UI ---

st.set_page_config(page_title="Smart Dev Report v0.4", layout="wide")
if 'photos' not in st.session_state: st.session_state.photos = [0]

st.title("🚀 Smart Dev Report Generator v0.4")

# ส่วนรับข้อมูลเอกสาร
with st.expander("📄 ข้อมูลเอกสาร (Document Info)", expanded=True):
    col1, col2, col3 = st.columns(3)
    doc_no = col1.text_input("เลขที่เอกสาร (Doc No.)")
    ref_po = col2.text_input("Ref. PO No.")
    date_val = col3.date_input("วันที่ (Date)", datetime.now())

# ส่วนจัดการรูปภาพ
st.header("📸 รายการรูปภาพ (Photo List)")
final_photo_data = []

for i in list(st.session_state.photos):
    with st.container():
        c_prev, c_input, c_del = st.columns([3, 5, 1])
        with c_input:
            up_img = st.file_uploader(f"เลือกรูปภาพ", type=['jpg','png','jpeg'], key=f"f{i}")
            up_desc = st.text_input(f"คำบรรยาย", key=f"d{i}")
        with c_prev:
            if up_img: st.image(up_img, use_container_width=True)
        with c_del:
            if st.button("🗑️", key=f"del{i}"):
                st.session_state.photos.remove(i)
                st.rerun()
        final_photo_data.append({"img": up_img, "desc": up_desc})
        st.markdown("---")

if st.button("➕ เพิ่มรูปภาพ"):
    st.session_state.photos.append(max(st.session_state.photos) + 1 if st.session_state.photos else 0)
    st.rerun()

# --- 4. ENGINE (ส่วนสร้างไฟล์) ---

if st.button("🚀 สร้างรายงานและส่งเมล", type="primary"):
    if not doc_no:
        st.warning("กรุณากรอกเลขที่เอกสารก่อนครับ")
        st.stop()

    try:
        # โหลด Template
        wb = load_workbook("template.xlsx")
        ws = wb.active # Sheet หลัก
        ws_temp = wb["ImageTemplate"]

        # เขียน Header หน้าแรก
        write_safe(ws, "B5", doc_no)
        write_safe(ws, "F6", ref_po)
        write_safe(ws, "J5", date_val.strftime("%d/%m/%Y"))

        # ตั้งค่าพิกัด
        loc_fixed = ["A49", "A62", "A75", "A92", "A105", "A118"]
        desc_fixed = ["H49", "H62", "H75", "H92", "H105", "H118"]
        
        start_gen_row = 131   # รูปที่ 7 เริ่มแถวนี้
        row_step = 13         # 1 บล็อกรูปมี 13 แถว
        header_h = 4          # หัวกระดาษมี 4 แถว
        gap_h = 4             # เว้น 4 แถวเมื่อจบกลุ่ม 3 รูป
        temp_row_start = 5    # ใน ImageTemplate บล็อกรูปเริ่มแถว 5

        for idx, item in enumerate(final_photo_data):
            if not item["img"]: continue
            
            # --- รูปที่ 1-6: ใช้ตำแหน่งที่มีอยู่แล้ว ---
            if idx < 6:
                p_loc, d_loc = loc_fixed[idx], desc_fixed[idx]
            
            # --- รูปที่ 7 เป็นต้นไป: สร้างใหม่ตามเงื่อนไข ---
            else:
                rel_idx = idx - 6
                num_pages = rel_idx // 3 # ทุก 3 รูปนับเป็น 1 หน้าใหม่
                
                # คำนวณแถวปัจจุบัน (รวมหัวกระดาษและช่องว่างที่เว้น)
                curr_row = start_gen_row + (rel_idx * row_step) + (num_pages * header_h) + (num_pages * gap_h)

                # ถ้าเป็นรูปแรกของกลุ่ม (7, 10, 13...) ให้ก๊อปปี้หัวกระดาษมาวาง
                if rel_idx % 3 == 0:
                    for h_r in range(1, header_h + 1):
                        target_h_row = curr_row - header_h + h_r - 1
                        ws.row_dimensions[target_h_row].height = ws_temp.row_dimensions[h_r].height
                        for c in range(1, 12):
                            copy_style(ws_temp.cell(row=h_r, column=c), ws.cell(row=target_h_row, column=c))

                # ก๊อปปี้บล็อกรูปภาพ (แถว 5-17 จาก Template)
                for r in range(0, row_step):
                    target_row = curr_row + r
                    ws.row_dimensions[target_row].height = ws_temp.row_dimensions[temp_row_start + r].height
                    for c in range(1, 12):
                        copy_style(ws_temp.cell(row=temp_row_start + r, column=c), ws.cell(row=target_row, column=c))
                
                # ก๊อปปี้ Merge Cells ของบล็อกรูป
                for m_range in ws_temp.merged_cells.ranges:
                    if m_range.min_row >= 5 and m_range.max_row <= 17:
                        t_off = m_range.min_row - temp_row_start
                        b_off = m_range.max_row - temp_row_start
                        new_m = f"{m_range.min_col_letter}{curr_row + t_off}:{m_range.max_col_letter}{curr_row + b_off}"
                        if new_m not in ws.merged_cells: ws.merge_cells(new_m)
                
                p_loc, d_loc = f"A{curr_row}", f"H{curr_row}"

            # วางรูปและข้อความลงไฟล์
            add_image_to_excel(ws, item["img"], p_loc)
            write_safe(ws, d_loc, item["desc"])

        # บันทึกไฟล์ลง Memory
        output = io.BytesIO()
        wb.save(output)
        file_data = output.getvalue()

        # --- 5. EMAIL SENDING (ส่งเมล) ---
        msg = MIMEMultipart()
        msg['From'], msg['To'], msg['Subject'] = SENDER_EMAIL, RECEIVER_EMAIL, f"Service Report: {doc_no}"
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="Report_{doc_no}.xlsx"')
        msg.attach(part)

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)

        st.success("✅ สร้างไฟล์และส่งอีเมลเรียบร้อยแล้ว!")
        st.download_button("📥 ดาวน์โหลดไฟล์ Excel", file_data, f"Report_{doc_no}.xlsx")
        st.balloons()

    except Exception as e:
        st.error(f"🚨 เกิดข้อผิดพลาด: {e}")

