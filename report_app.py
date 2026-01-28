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

# <<< ADDED
from PIL import Image as PILImage
import excel2img
import tempfile
import os

# --- 1. CONFIGURATION ---
SENDER_EMAIL = "jinjutar.smartdev@gmail.com"
SENDER_PASSWORD = "YOUR_APP_PASSWORD"
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

# <<< ADDED: resize รูปอัตโนมัติ
def resize_image(uploaded_file, max_size=1400):
    img = PILImage.open(uploaded_file)

    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")

    img.thumbnail((max_size, max_size))

    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=85)
    buffer.seek(0)
    return buffer

def add_image_to_excel(ws, img_file, cell_address):
    if img_file is None: return

    resized = resize_image(img_file)  # <<< CHANGED
    img = Image(resized)

    max_w, max_h = 0, 0
    found_range = None
    for m_range in ws.merged_cells.ranges:
        if ws[cell_address].coordinate in m_range:
            found_range = m_range
            for col in range(m_range.min_col, m_range.max_col + 1):
                max_w += (ws.column_dimensions[get_column_letter(col)].width or 8.43) * 7.5
            for row in range(m_range.min_row, m_range.max_row + 1):
                max_h += (ws.row_dimensions[row].height or 15) * 1.33
            break
    if not found_range:
        max_w, max_h = 350, 250

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

# <<< ADDED: แปลง Excel → รูป preview
def render_excel_preview(wb):
    temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".png")

    wb.save(temp_excel.name)
    excel2img.export_img(temp_excel.name, temp_img.name, "", "A1:L250")

    return temp_img.name

# --- 3. STREAMLIT UI ---
st.set_page_config(page_title="Smart Dev Report Generator", layout="wide")
if 'photos' not in st.session_state:
    st.session_state.photos = [0]

st.title("🚀 Smart Dev Report Generator v0.7.1")

st.subheader("📄 Part 1: Document Details")
c1, c2, c3 = st.columns(3)
doc_no = c1.text_input("Doc. No.")
ref_po = c2.text_input("Ref. PO No.")
date_issue = c3.date_input("Date", datetime.now())

st.markdown("---")
st.subheader("🏢 Part 2: Project & Contact Information")
p1, p2 = st.columns(2)
project_name = p1.text_input("Project Name")
site_location = p1.text_input("Site / Location")
contact_client = p2.text_input("Contact Person (Client)")
contact_co_ltd = p2.text_input("Contact (Smart Dev Co., Ltd.)")
engineer_name = st.text_input("Engineer Name (Prepared By)")

st.markdown("---")
st.subheader("🛠 Part 3: Service Details")
service_type = st.selectbox("Service Type", ["Project", "Commissioning", "Repairing", "Services", "Training", "Check", "Other"])
job_performed = st.text_area("Job Performed (รายละเอียดงาน)", height=150)

st.markdown("---")
st.subheader("📸 Part 4: Photo Report")
final_photo_data = []

for i in list(st.session_state.photos):
    with st.container():
        col_prev, col_input, col_del = st.columns([3, 5, 1])
        with col_input:
            up_img = st.file_uploader(f"Upload Image {i+1}", type=['jpg','png','jpeg'], key=f"f{i}")
            up_desc = st.text_input(f"Description {i+1}", key=f"d{i}")
        with col_prev:
            if up_img:
                st.image(up_img, use_container_width=True)
        with col_del:
            if st.button("🗑️", key=f"del{i}"):
                st.session_state.photos.remove(i)
                st.rerun()
        final_photo_data.append({"img": up_img, "desc": up_desc})

if st.button("➕ Add More Photo"):
    st.session_state.photos.append(max(st.session_state.photos) + 1)
    st.rerun()

# ---------- PREVIEW ----------
st.markdown("---")
st.subheader("👀 Preview Report (เหมือนหน้า Excel)")

if st.button("Preview หน้า Report"):
    try:
        wb = load_workbook("template.xlsx")
        ws = wb.active
        ws_temp = wb["ImageTemplate"]

        write_safe(ws, "B5", doc_no)
        write_safe(ws, "F6", ref_po)
        write_safe(ws, "J5", date_issue.strftime('%d/%m/%Y'))
        write_safe(ws, "B16", project_name)
        write_safe(ws, "H7", site_location)
        write_safe(ws, "C9", contact_client)
        write_safe(ws, "A7", contact_co_ltd)
        write_safe(ws, "B42", engineer_name)
        write_safe(ws, "D15", service_type)
        write_safe(ws, "D17", job_performed)

        loc_fixed = ["A49", "A62", "A75", "A92", "A105", "A118"]
        desc_fixed = ["H49", "H62", "H75", "H92", "H105", "H118"]

        for idx, item in enumerate(final_photo_data):
            if idx >= 6: break
            if item["img"]:
                add_image_to_excel(ws, item["img"], loc_fixed[idx])
                write_safe(ws, desc_fixed[idx], item["desc"])

        img_path = render_excel_preview(wb)
        st.image(img_path, use_container_width=True)

    except Exception as e:
        st.error(f"Preview error: {e}")

# --- 4. ENGINE (Generate & Send) ---
# >>> โค้ด generate & send ของคุณเดิม ใช้ได้เหมือนเดิมทั้งหมด <<<
