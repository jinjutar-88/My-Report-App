from openpyxl.drawing.image import Image

def add_image_to_excel(ws, img_path, cell_address):
    """
    ฟังก์ชันวางรูปภาพให้ขนาดพอดีกับช่อง Excel (รองรับทั้งช่องเดี่ยวและช่องที่ Merge)
    """
    img = Image(img_path)
    
    # 1. หาขนาดของช่อง (พิกเซล)
    # กรณีช่องที่ระบุมีการ Merge ไว้ เราต้องบวกความกว้าง/สูงของทุกช่องที่รวมกัน
    target_width = 0
    target_height = 0
    
    # ตรวจสอบว่า cell อยู่ในพื้นที่ Merge หรือไม่
    merged_range = None
    for m_range in ws.merged_cells.ranges:
        if cell_address in m_range:
            merged_range = m_range
            break
            
    if merged_range:
        # ถ้ามีการ Merge: คำนวณความกว้างรวมของทุกคอลัมน์ใน Range
        for col in range(merged_range.min_col, merged_range.max_col + 1):
            col_letter = ws.cell(row=1, column=col).column_letter
            target_width += (ws.column_dimensions[col_letter].width or 8.43) * 7.5
        # คำนวณความสูงรวมของทุกแถวใน Range
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            target_height += (ws.row_dimensions[row].height or 15) * 1.33
    else:
        # ถ้าเป็นช่องเดี่ยว: ใช้ขนาดปกติ
        col_letter = cell_address[0]
        row_num = int(''.join(filter(str.isdigit, cell_address)))
        target_width = (ws.column_dimensions[col_letter].width or 8.43) * 7.5
        target_height = (ws.row_dimensions[row_num].height or 15) * 1.33

    # 2. ตั้งค่าขนาดรูป (ปรับให้เล็กลงกว่าช่องนิดหน่อยเพื่อไม่ให้ทับเส้นขอบ)
    img.width = target_width - 5
    img.height = target_height - 5
    
    # 3. วางรูป
    ws.add_image(img, cell_address)

# --- ตัวอย่างการเรียกใช้ในลูป (ภายใต้ปุ่ม Submit) ---
# สมมติว่ามีรูป 4 รูป (img1, img2, img3, img4) 
# และต้องการวางในตำแหน่งที่เตรียมไว้ใน template
photo_locations = ["B58", "F58", "B75", "F75"] # แก้ไขพิกัดให้ตรงกับไฟล์จริงของคุณ
uploaded_imgs = [img1, img2, img3, img4] # ลิสต์ของรูปที่รับมาจาก file_uploader

for loc, img_file in zip(photo_locations, uploaded_imgs):
    if img_file:
        # บันทึกเป็นไฟล์ชั่วคราวเพื่อส่งให้ openpyxl
        with open(f"temp_{loc}.png", "wb") as f:
            f.write(img_file.getbuffer())
        add_image_to_excel(ws, f"temp_{loc}.png", loc)
