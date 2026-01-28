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

# ---------------- CONFIG ----------------
TEMPLATE_FILE = "template.xlsx"

SENDER_EMAIL = "jinjutar.smartdev@gmail.com"
SENDER_PASSWORD = "uzfsbdtcxclzrzsq"   # ต้องเป็น app password
RECEIVER_EMAIL = "jinjutar.smartdev@gmail.com"

# ---------------- HELPERS ----------------
def add_image_to_excel(ws, img_file, cell_address):
    if img_file is None:
        return

    img_data = io.BytesIO(img_file.getvalue())
    img = Image(img_data)

    max_w, max_h = 350, 250
    for m in ws.merged_cells.ranges:
        if ws[cell_address].coordinate in m:
            max_w = sum((ws.column_dimensions[get_column_letter(c)].width or 8.43) * 7.5
                        for c in range(m.min_col, m.max_col + 1))
            max_h = sum((ws.row_dimensions[r].height or 15) * 1.33
                        for r in range(m.min_row, m.max_row + 1))
            break

    ratio = min(max_w / img.width, max_h / img.height)
    img.width = int(img.width * ratio)
    img.height = int(img.height * ratio)

    ws.add_image(img, cell_address)


def write_safe(ws, cell_addr, value):
    if value is None:
        value = ""
    for m in ws.merged_cells.ranges:
        if cell_addr in m:
            ws.cell(row=m.min_row, column=m.min_col).value = value
            return
    ws[cell_addr] = value


def copy_template_block(ws, ws_temp, start_row):
    for r in range(1, ws_temp.max_row + 1):
        ws.row_dimensions[start_row + r - 1].height = ws_temp.row_dimensions[r].height
        for c in range(1, 12):
            src = ws_temp.cell(r, c)
            dst = ws.cell(start_row + r - 1, c)
            dst.value = src.value
            if src.has_style:
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.fill = copy(src.fill)
                dst.number_format = copy(src.number_format)
                dst.alignment = copy(src.alignment)

    for m in ws_temp.merged_cells.ranges:
        new_range = f"{get_column_letter(m.min_col)}{start_row + m.min_row - 1}:{get_column_letter(m.max_col)}{start_row + m.max_row - 1}"
        ws.merge_cells(new_range)


# ---------------- UI ----------------
st.set_page_config(page_title="Smart Dev Report", layout="wide")

if "photos" not in st.session_state:
    st.session_state.photos = [0]

st.title("🚀 Smart Dev Report Generator")

# ---- Part 1
st.subheader("📄 Part 1")
c1, c2, c3 = st.columns(3)
doc_no = c1.text_input("Doc No.")
ref_po = c2.text_input("PO No.")
date_issue = c3.date_input("Date", datetime.now())

# ---- Part 2
st.subheader("🏢 Part 2")
p1, p2 = st.columns(2)
project_name = p1.text_input("Project Name")
site_location = p1.text_input("Site Location")
contact_client = p2.text_input("Client Contact")
contact_co = p2.text_input("Smart Dev Contact")
engineer = st.text_input("Engineer")

# ---- Part 3
st.subheader("🛠 Part 3")
service_type = st.selectbox("Service Type", ["Project", "Commissioning", "Repairing", "Service", "Training"])
job_detail = st.text_area("Job Detail", height=120)

# ---- Part 4
st.subheader("📸 Part 4: Photo Report")
final_photo_data = []

for i in list(st.session_state.photos):
    st.markdown(f"### Photo {i+1}")

    img = st.file_uploader(f"Upload image {i+1}", type=["jpg", "png", "jpeg"], key=f"f{i}")
    if img:
        st.image(img, use_container_width=True)

    desc = st.text_input(f"Description {i+1}", key=f"d{i}")

    if st.button("🗑️ Delete photo", key=f"del{i}"):
        st.session_state.photos.remove(i)
        st.rerun()

    st.markdown("---")
    final_photo_data.append({"img": img, "desc": desc})

if st.button("➕ Add more photo"):
    st.session_state.photos.append(max(st.session_state.photos) + 1)
    st.rerun()


# ---------------- GENERATE ----------------
if st.button("🚀 Generate, Send Email & Download"):
    try:
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb["1"]
        ws_temp = wb["ImageTemplate"]

        # Write text
        write_safe(ws, "B5", doc_no)
        write_safe(ws, "F6", ref_po)
        write_safe(ws, "J5", date_issue.strftime("%d/%m/%Y"))
        write_safe(ws, "B16", project_name)
        write_safe(ws, "H7", site_location)
        write_safe(ws, "C9", contact_client)
        write_safe(ws, "A7", contact_co)
        write_safe(ws, "B42", engineer)
        write_safe(ws, "D15", service_type)
        write_safe(ws, "D17", job_detail)

        # Photos 1–6
        loc = ["A49", "A62", "A75", "A92", "A105", "A118"]
        desc_loc = ["H49", "H62", "H75", "H92", "H105", "H118"]

        for i, item in enumerate(final_photo_data[:6]):
            if item["img"]:
                add_image_to_excel(ws, item["img"], loc[i])
                write_safe(ws, desc_loc[i], item["desc"])

        # Photos 7+
        extra = final_photo_data[6:]
        if extra:
            start_row = ws.max_row + 2
            rows_per_page = ws_temp.max_row

            for i in range(0, len(extra), 3):
                group = extra[i:i+3]
                copy_template_block(ws, ws_temp, start_row)

                img_rows = [5, 18, 31]
                desc_rows = [5, 18, 31]

                for j, item in enumerate(group):
                    if item["img"]:
                        add_image_to_excel(ws, item["img"], f"A{start_row + img_rows[j] - 1}")
                        write_safe(ws, f"H{start_row + desc_rows[j] - 1}", item["desc"])

                start_row += rows_per_page + 1

        # Save to memory
        output = io.BytesIO()
        wb.save(output)

        # ---------------- SEND EMAIL ----------------
        email_status = "❌ Email not sent"
        try:
            msg = MIMEMultipart()
            msg["From"] = SENDER_EMAIL
            msg["To"] = RECEIVER_EMAIL
            msg["Subject"] = f"Report: {doc_no}"

            part = MIMEBase("application", "octet-stream")
            part.set_payload(output.getvalue())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="Report_{doc_no}.xlsx"')
            msg.attach(part)

            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.send_message(msg)

            email_status = "✅ Email sent successfully"

        except Exception as mail_err:
            email_status = f"⚠️ Email failed: {mail_err}"

        # ---------------- UI RESULT ----------------
        st.success("✅ Excel generated")
        st.info(email_status)

        st.download_button(
            "📥 Download Excel",
            output.getvalue(),
            f"Report_{doc_no}.xlsx"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
