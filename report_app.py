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

# ================= CONFIG =================
SENDER_EMAIL = "jinjutar.smartdev@gmail.com"
SENDER_PASSWORD = "UZFS BDTC XCLZ RZSQ"
RECEIVER_EMAIL = "jinjutar.smartdev@gmail.com"

TEMPLATE_FILE = "template.xlsx"
MAIN_SHEET = "1"
IMAGE_TEMPLATE_SHEET = "ImageTemplate"

# ================= HELPERS =================
def add_image_to_excel(ws, img_file, cell_address):
    if img_file is None:
        return

    img_data = io.BytesIO(img_file.getvalue())
    img = Image(img_data)

    max_w, max_h = 350, 250
    for m_range in ws.merged_cells.ranges:
        if ws[cell_address].coordinate in m_range:
            max_w = sum(
                (ws.column_dimensions[get_column_letter(c)].width or 8.43) * 7.5
                for c in range(m_range.min_col, m_range.max_col + 1)
            )
            max_h = sum(
                (ws.row_dimensions[r].height or 15) * 1.33
                for r in range(m_range.min_row, m_range.max_row + 1)
            )
            break

    ratio = min(max_w / img.width, max_h / img.height)
    img.width = int(img.width * ratio)
    img.height = int(img.height * ratio)

    ws.add_image(img, cell_address)


def write_safe(ws, cell_addr, value):
    if value is None:
        value = ""
    for m_range in ws.merged_cells.ranges:
        if cell_addr in m_range:
            ws.cell(row=m_range.min_row, column=m_range.min_col).value = value
            return
    ws[cell_addr] = value


# ================= UI =================
st.set_page_config(page_title="Smart Dev Report Generator", layout="wide")

if 'photos' not in st.session_state:
    st.session_state.photos = [0]

st.title("📄 Smart Dev Report Generator")

st.subheader("Part 1")
c1, c2, c3 = st.columns(3)
doc_no = c1.text_input("Doc No.")
ref_po = c2.text_input("Ref PO")
date_issue = c3.date_input("Date", datetime.now())

st.subheader("Part 2")
project_name = st.text_input("Project")
site_location = st.text_input("Site")
contact_client = st.text_input("Client")
contact_co_ltd = st.text_input("Smart Dev Contact")
engineer_name = st.text_input("Engineer")

st.subheader("Part 3")
service_type = st.selectbox("Service Type", ["Project","Commissioning","Repairing","Services","Training","Check","Other"])
job_performed = st.text_area("Job Performed")

st.subheader("📸 Photo Upload")
final_photo_data = []

for i in list(st.session_state.photos):
    c1, c2, c3 = st.columns([3,5,1])

    with c2:
        img = st.file_uploader(f"Image {i+1}", type=["jpg","jpeg","png"], key=f"img{i}")
        desc = st.text_input(f"Description {i+1}", key=f"desc{i}")

    with c1:
        if img:
            st.image(img, use_container_width=True)

    with c3:
        if st.button("❌", key=f"del{i}"):
            st.session_state.photos.remove(i)
            st.rerun()

    final_photo_data.append({"img": img, "desc": desc})

if st.button("➕ Add photo"):
    st.session_state.photos.append(max(st.session_state.photos)+1)
    st.rerun()


# ================= GENERATE =================
if st.button("🚀 Generate & Send"):
    try:
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb[MAIN_SHEET]
        ws_temp = wb[IMAGE_TEMPLATE_SHEET]

        # --- text ---
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

        # --- first 6 photos ---
        loc_fixed = ["A49","A62","A75","A92","A105","A118"]
        desc_fixed = ["H49","H62","H75","H92","H105","H118"]

        for i,item in enumerate(final_photo_data[:6]):
            if item["img"]:
                add_image_to_excel(ws, item["img"], loc_fixed[i])
                write_safe(ws, desc_fixed[i], item["desc"])

        # --- photo 7+ using real template ---
        extra = final_photo_data[6:]

        img_cells_template = ["A5", "A18", "A31"]
        desc_cells_template = ["H5", "H18", "H31"]
        rows_template = 43

        if extra:
            start_row = ws.max_row + 2

            for batch in range(0, len(extra), 3):
                group = extra[batch:batch+3]

                # copy template
                for r in range(1, rows_template+1):
                    ws.row_dimensions[start_row+r-1].height = ws_temp.row_dimensions[r].height
                    for c in range(1, 12):
                        src = ws_temp.cell(r,c)
                        dst = ws.cell(start_row+r-1,c)
                        dst.value = src.value
                        if src.has_style:
                            dst.font = copy(src.font)
                            dst.border = copy(src.border)
                            dst.fill = copy(src.fill)
                            dst.alignment = copy(src.alignment)

                # merge
                for m in ws_temp.merged_cells.ranges:
                    new_range = f"{get_column_letter(m.min_col)}{start_row+m.min_row-1}:{get_column_letter(m.max_col)}{start_row+m.max_row-1}"
                    ws.merge_cells(new_range)

                # insert images
                for i,item in enumerate(group):
                    add_image_to_excel(ws, item["img"], f"A{start_row + [5,18,31][i]-1}")
                    write_safe(ws, f"H{start_row + [5,18,31][i]-1}", item["desc"])

                start_row += rows_template + 2

        # save
        output = io.BytesIO()
        wb.save(output)

        # email
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = RECEIVER_EMAIL
        msg["Subject"] = f"Report {doc_no}"

        part = MIMEBase("application", "octet-stream")
        part.set_payload(output.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="Report_{doc_no}.xlsx"')
        msg.attach(part)

        with smtplib.SMTP("smtp.gmail.com",587) as server:
            server.starttls()
            server.login(SENDER_EMAIL,SENDER_PASSWORD)
            server.send_message(msg)

        st.success("✅ สำเร็จ ส่งเมลแล้ว")
        st.download_button("📥 Download Excel", output.getvalue(), f"Report_{doc_no}.xlsx")

    except Exception as e:
        st.error(f"🚨 {e}")
