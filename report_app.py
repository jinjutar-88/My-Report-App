import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from datetime import datetime
import io
import smtplib
import os
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from copy import copy

# ---------------- CONFIG ----------------
load_dotenv()

SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")

TEMPLATE_FILE = "template.xlsx"

# ---------------- HELPERS ----------------
def copy_style(source, target):
    if source.has_style:
        target.font = copy(source.font)
        target.border = copy(source.border)
        target.fill = copy(source.fill)
        target.number_format = copy(source.number_format)
        target.protection = copy(source.protection)
        target.alignment = copy(source.alignment)

def write_safe(ws, cell_addr, value):
    value = "" if value is None else value
    for r in ws.merged_cells.ranges:
        if cell_addr in r:
            ws.cell(r.min_row, r.min_col).value = value
            return
    ws[cell_addr] = value

def add_image_to_excel(ws, img_file, cell):
    if not img_file:
        return

    img = Image(io.BytesIO(img_file.getvalue()))

    max_w, max_h = 350, 250
    for r in ws.merged_cells.ranges:
        if ws[cell].coordinate in r:
            max_w = sum((ws.column_dimensions[get_column_letter(c)].width or 8.43) * 7.5 for c in range(r.min_col, r.max_col + 1))
            max_h = sum((ws.row_dimensions[row].height or 15) * 1.33 for row in range(r.min_row, r.max_row + 1))
            break

    ratio = min((max_w - 10) / img.width, (max_h - 10) / img.height)
    img.width = int(img.width * ratio)
    img.height = int(img.height * ratio)

    ws.add_image(img, cell)

# ---------------- UI ----------------
st.set_page_config("Smart Dev Report Generator", layout="wide")

if "photos" not in st.session_state:
    st.session_state.photos = [0]

st.title("🚀 Smart Dev Report Generator v0.8")

st.subheader("📄 Part 1")
c1, c2, c3 = st.columns(3)
doc_no = c1.text_input("Doc. No.")
ref_po = c2.text_input("Ref. PO No.")
date_issue = c3.date_input("Date", datetime.now())

st.subheader("🏢 Part 2")
p1, p2 = st.columns(2)
project_name = p1.text_input("Project Name")
site_location = p1.text_input("Site / Location")
contact_client = p2.text_input("Contact Client")
contact_company = p2.text_input("Contact Smart Dev")
engineer_name = st.text_input("Engineer")

st.subheader("🛠 Part 3")
service_type = st.selectbox("Service Type", ["Project", "Commissioning", "Repairing", "Services", "Training", "Check", "Other"])
job_performed = st.text_area("Job Performed", height=120)

st.subheader("📸 Part 4 Photos")
photo_data = []

for i in list(st.session_state.photos):
    col_prev, col_input, col_del = st.columns([3, 5, 1])

    with col_input:
        img = st.file_uploader(f"Image {i+1}", type=["jpg", "png", "jpeg"], key=f"img{i}")
        desc = st.text_input(f"Description {i+1}", key=f"desc{i}")

    with col_prev:
        if img:
            st.image(img, use_container_width=True)

    with col_del:
        if st.button("🗑️", key=f"del{i}"):
            st.session_state.photos.remove(i)
            st.rerun()

    photo_data.append({"img": img, "desc": desc})

if st.button("➕ Add Photo"):
    st.session_state.photos.append(max(st.session_state.photos) + 1)
    st.rerun()

# ---------------- ENGINE ----------------
if st.button("🚀 Generate Report", use_container_width=True):

    try:
        valid_photos = [p for p in photo_data if p["img"]]

        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active
        ws_temp = wb["ImageTemplate"]

        # write data
        write_safe(ws, "B5", doc_no)
        write_safe(ws, "F6", ref_po)
        write_safe(ws, "J5", date_issue.strftime('%d/%m/%Y'))
        write_safe(ws, "B16", project_name)
        write_safe(ws, "H7", site_location)
        write_safe(ws, "C9", contact_client)
        write_safe(ws, "A7", contact_company)
        write_safe(ws, "B42", engineer_name)
        write_safe(ws, "D15", service_type)
        write_safe(ws, "D17", job_performed)

        loc_fixed = ["A49", "A62", "A75", "A92", "A105", "A118"]
        desc_fixed = ["H49", "H62", "H75", "H92", "H105", "H118"]

        cursor = 174
        header_h, block_h = 4, 13

        for idx, item in enumerate(valid_photos):
            if idx < 6:
                p_loc, d_loc = loc_fixed[idx], desc_fixed[idx]
            else:
                rel = idx - 6

                if rel % 3 == 0:
                    for r in range(1, header_h + 1):
                        for c in range(1, 12):
                            src = ws_temp.cell(r, c)
                            tgt = ws.cell(cursor, c)
                            tgt.value = src.value
                            copy_style(src, tgt)
                        cursor += 1

                p_row = cursor
                for r in range(block_h):
                    for c in range(1, 12):
                        copy_style(ws_temp.cell(5 + r, c), ws.cell(p_row + r, c))

                p_loc, d_loc = f"A{p_row}", f"H{p_row}"
                cursor += block_h

            add_image_to_excel(ws, item["img"], p_loc)
            write_safe(ws, d_loc, item["desc"])

        output = io.BytesIO()
        wb.save(output)

        st.success("✅ Generate สำเร็จ")
        st.download_button("📥 Download Excel", output.getvalue(), f"Report_{doc_no}.xlsx")

    except Exception as e:
        st.error(f"❌ Error: {e}")
