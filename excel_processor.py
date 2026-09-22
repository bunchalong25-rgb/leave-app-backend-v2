import io
import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def clean_dept_name(text: str) -> str:
    """ตัดเครื่องหมายขีด (-) ออก และจัดช่องว่างให้เรียบร้อย"""
    if not text:
        return ""
    cleaned = str(text).replace('-', '').strip()
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip()

def clean_brand_text(text: str) -> str:
    """ตัดคำนำหน้า แบรนด์:, แบรนด์ :, Brand: ออกให้สะอาด"""
    if not text:
        return ""
    cleaned = str(text).strip()
    cleaned = re.sub(r'^(แบรนด์\s*/?\s*สาขา\s*[:：]?\s*|แบรนด์\s*[:：]?\s*|brand\s*[:：]?\s*)', '', cleaned, flags=re.IGNORECASE).strip()
    cleaned = cleaned.strip('-').strip()
    return cleaned

def clean_name_text(text: str) -> str:
    """ตัดคำนำหน้า ชื่อ-นามสกุล:, ชื่อ-สกุล:, ชื่อพนักงาน:, ชื่อ:, Name:, Fullname: ออกให้สะอาด"""
    if not text:
        return ""
    cleaned = str(text).strip()
    cleaned = re.sub(r'^(ชื่อ\s*[-–—]\s*นามสกุล\s*[:：]?\s*|ชื่อ\s*[-–—]\s*สกุล\s*[:：]?\s*|ชื่อพนักงาน\s*[:：]?\s*|ชื่อ\s*[:：]?\s*|fullname\s*[:：]?\s*|name\s*[:：]?\s*)', '', cleaned, flags=re.IGNORECASE).strip()
    cleaned = cleaned.strip('-').strip()
    return cleaned

def create_sample_template(file_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ตารางวันหยุดพนักงาน"
    ws.views.sheetView[0].showGridLines = True

    # Title Banner (Row 1)
    ws.merge_cells('A1:AI1')
    title_cell = ws['A1']
    title_cell.value = "ตารางการทำงานและวันหยุดพนักงานประจำเดือน (Template ต้นฉบับ)"
    title_cell.font = Font(name='Sarabun', size=16, bold=True, color='FFFFFF')
    title_cell.fill = PatternFill(start_color='0F172A', end_color='0F172A', fill_type='solid')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 36

    # Subtitle Legend (Row 2)
    ws.merge_cells('A2:AI2')
    legend_cell = ws['A2']
    legend_cell.value = "สัญลักษณ์: W = วันหยุดประจำเดือน | C = วันเข้าบริษัท | V = วันลาพักร้อน | S = วันลาป่วย | N = วันหยุดไม่มีคนแทน"
    legend_cell.font = Font(name='Sarabun', size=10, italic=True, color='334155')
    legend_cell.fill = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
    legend_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 24

    # Table Column Headers (Row 3)
    headers = ['ลำดับ', 'แบรนด์ / สาขา', 'ชื่อ-นามสกุล'] + [str(d) for d in range(1, 32)]
    ws.row_dimensions[3].height = 28
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='medium', color='0284C7')
    )

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=int(h) if h.isdigit() else h)
        cell.font = Font(name='Sarabun', size=11, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='0284C7', end_color='0284C7', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # Sample Employees
    sample_employees = [
        {"no": 1, "brand": "Brand Alpha (สาขา สยาม)", "name": "สมชาย สายดี"},
        {"no": 2, "brand": "Brand Alpha (สาขา สยาม)", "name": "วิภาวี มั่นคง"},
        {"no": 3, "brand": "Brand Beta (สาขา ชิดลม)", "name": "อภิสิทธิ์ ขยันทำ"},
        {"no": 4, "brand": "Brand Beta (สาขา ชิดลม)", "name": "นภา เพลินตา"},
        {"no": 5, "brand": "Brand Delta (สาขา ไอคอนสยาม)", "name": "กิตติพงษ์ ยอดเยี่ยม"},
        {"no": 6, "brand": "Brand Alpha (สาขา สยาม)", "name": "ธนากร มุ่งมั่น"}
    ]

    cell_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    for idx, emp in enumerate(sample_employees, 4):
        ws.row_dimensions[idx].height = 24
        ws.cell(row=idx, column=1, value=emp["no"]).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=idx, column=2, value=emp["brand"]).alignment = Alignment(horizontal='left', vertical='center')
        
        name_cell = ws.cell(row=idx, column=3, value=emp["name"])
        name_cell.font = Font(bold=True)
        name_cell.alignment = Alignment(horizontal='left', vertical='center')

        for d in range(1, 32):
            c = ws.cell(row=idx, column=d + 3)
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.font = Font(name='Sarabun', size=11, bold=True)
            c.border = cell_border

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 24
    for d in range(1, 32):
        col_letter = openpyxl.utils.get_column_letter(d + 3)
        ws.column_dimensions[col_letter].width = 5

    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    wb.save(file_path)
    print(f"Sample Python openpyxl template created at {file_path}")

def process_excel_template(template_path: str, leave_records: list, period: str = None) -> bytes:
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # 1. Locate Day Columns (1..31) in header rows (Rows 1..10)
    day_col_map = {}
    header_row_index = -1

    for r in range(1, 11):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None:
                val_str = str(val).strip()
                if val_str.isdigit():
                    num = int(val_str)
                    if 1 <= num <= 31:
                        day_col_map[num] = c
                        header_row_index = r
        if len(day_col_map) >= 15:
            break

    # 2. Locate Employee Name Rows
    emp_row_map = {}
    for r in range(header_row_index + 1, ws.max_row + 1):
        for c in range(1, min(ws.max_column + 1, 10)):
            val = ws.cell(row=r, column=c).value
            if val and isinstance(val, str):
                raw_name = val.strip()
                clean_name = clean_name_text(raw_name)
                if len(clean_name) >= 2 and clean_name not in emp_row_map:
                    emp_row_map[clean_name] = r
                if len(raw_name) >= 2 and raw_name not in emp_row_map:
                    emp_row_map[raw_name] = r

    # 3. Fill Leave Codes
    updated_count = 0
    code_colors = {
        'W': '2563EB', # Blue
        'C': '059669', # Emerald
        'V': 'D97706', # Amber
        'S': 'DC2626', # Red
        'N': '475569'  # Slate
    }

    for rec in leave_records:
        day_num = rec.get('dayNumber')
        if not day_num:
            continue

        if period == '1-15' and day_num > 15:
            continue
        if period == '16-end' and day_num < 16:
            continue

        col_idx = day_col_map.get(day_num)
        if not col_idx:
            continue

        rec_name = (rec.get('userName') or '').strip()
        rec_clean_name = clean_name_text(rec_name)
        target_row = emp_row_map.get(rec_clean_name) or emp_row_map.get(rec_name)

        if not target_row and (rec_clean_name or rec_name):
            search_key = rec_clean_name or rec_name
            for k, r_idx in emp_row_map.items():
                if search_key in k or k in search_key:
                    target_row = r_idx
                    break

        if target_row and col_idx:
            cell = ws.cell(row=target_row, column=col_idx)
            code = rec.get('code', '')
            cell.value = code
            
            color_argb = code_colors.get(code, '0F172A')
            cell.font = Font(name='Sarabun', size=11, bold=True, color=color_argb)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            updated_count += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), updated_count

def process_master_data_excel(file_content: bytes, min_staff_threshold: int = 2):
    """
    ฟังก์ชันสำหรับอ่านไฟล์ Excel อัจฉริยะ (รองรับทุก Sheet และรองรับหลายแผนกแทรกใน Sheet เดียวกัน)
    - ตรวจจับแถวแผนก เช่น '--- แผนก 101 :PRESTIGE ชั้น 1 ---' (Col A มีคำว่าแผนก, Col B ว่างเปล่า)
    - ผูกฟิลด์ department เข้ากับพนักงานในแผนกนั้นๆ อย่างถูกต้อง
    - ตัดคำนำหน้า 'แบรนด์:' และ 'ชื่อ-นามสกุล:' ออกให้สะอาด
    - รวบรวมรายชื่อแผนกทั้งหมด และนับจำนวนคนในแบรนด์
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_content), data_only=True)
    
    all_employees = []
    brand_counts = {}
    remarks = []
    departments_list = []

    DEPT_KEYWORDS = ["แผนก", "ฝ่าย", "DEPARTMENT", "SECTION"]
    STOP_KEYWORDS = ["สรุปจำนวน", "รวมทั้งสิ้น", "สรุปยอด", "ยอดรวม", "TOTAL"]
    SHIFT_KEYWORDS = ["คำอธิบายสัญลักษณ์", "ช่วงเวลาปฏิบัติ", "เวลาปฏิบัติงาน", "เวลาทำงาน", "กะการทำงาน", "กะ ", "กะเช้า", "กะบ่าย", "กะดึก", "W =", "C =", "V =", "S =", "N =", "OFF"]

    # 1. วนลูปอ่านข้อมูล "ทุก Sheet" ในไฟล์
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        current_dept = None

        # กรณีชื่อชีตมีชื่อแผนกอยู่ ใช้เป็น fallback
        fallback_dept = None
        if any(k in sheet_name.upper() for k in DEPT_KEYWORDS):
            fallback_dept = clean_dept_name(sheet_name)

        # 2. วนลูปอ่านแถวทีละแถวตั้งแต่บรรทัดแรกจนถึงบรรทัดสุดท้าย
        for row in sheet.iter_rows(values_only=True):
            if not row:
                continue

            # แปลงค่าในแต่ละเซลล์เป็นสตริงเพื่อตรวจจับ
            row_texts = [str(c).strip() if c is not None else "" for c in row]
            if not any(row_texts):
                # แถวว่างเปล่าคั่นระหว่างกลุ่ม/แผนก ให้ข้ามไปแถวถัดไป (ห้าม break หลุดลูป)
                continue

            col_a_raw = row_texts[0] if len(row_texts) > 0 else ""
            col_b_raw = row_texts[1] if len(row_texts) > 1 else ""

            # ข้ามแถวสรุปยอดรวม (เช่น 'รวมทั้งสิ้น ...')
            row_sample_str = " ".join(row_texts[:5])
            if any(stop_word in row_sample_str for stop_word in STOP_KEYWORDS):
                continue

            # -------------------------------------------------------------
            # ก) ตรวจจับแถวหัวข้อแผนก:
            # Col A มีคำว่า "แผนก" (หรือ "ฝ่าย", "DEPARTMENT")
            # และ Col B ว่างเปล่า หรือไม่มีชื่อพนักงาน (หรือเป็น None / -)
            # -------------------------------------------------------------
            is_dept_header = False
            col_a_upper = col_a_raw.upper()
            if any(k in col_a_upper for k in DEPT_KEYWORDS):
                b_name_clean = clean_name_text(col_b_raw)
                if not col_b_raw or col_b_raw.upper() in ["NONE", "NULL", "-"] or not b_name_clean:
                    is_dept_header = True

            if is_dept_header:
                parsed_dept = clean_dept_name(col_a_raw)
                if parsed_dept and len(parsed_dept) >= 2:
                    current_dept = parsed_dept
                    if current_dept not in departments_list:
                        departments_list.append(current_dept)
                continue

            # -------------------------------------------------------------
            # ข) ตรวจจับแถวพนักงาน:
            # Col A มีแบรนด์ (เช่น 'แบรนด์: MAC '), Col B มีชื่อคน (เช่น ' ชื่อ-นามสกุล: กัณฐมณี')
            # -------------------------------------------------------------
            brand_val = clean_brand_text(col_a_raw)
            name_val = clean_name_text(col_b_raw)

            # ข้ามแถวหัวคอลัมน์ทั่วไป (เช่น 'แบรนด์ / สาขา' | 'ชื่อ-นามสกุล')
            is_generic_header = (
                col_a_raw in ["แบรนด์ / สาขา", "แบรนด์/สาขา", "แบรนด์", "สาขา", "BRAND"] or
                col_b_raw in ["ชื่อ-นามสกุล", "ชื่อ - นามสกุล", "ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ", "FULLNAME", "NAME"]
            )
            if is_generic_header:
                continue

            # ถ้าพบทั้งแบรนด์และชื่อพนักงาน
            if brand_val and name_val and brand_val.upper() not in ["NONE", "NULL", "-"] and name_val.upper() not in ["NONE", "NULL", "-"]:
                assigned_dept = current_dept or fallback_dept or sheet_name.strip()
                all_employees.append({
                    "name": name_val,
                    "brand": brand_val,
                    "department": assigned_dept
                })
                brand_counts[brand_val] = brand_counts.get(brand_val, 0) + 1
                if assigned_dept and assigned_dept not in departments_list:
                    departments_list.append(assigned_dept)
                continue

            # เผื่อกรณีตารางมีคอลัมน์คั่น เช่น Col A = ลำดับ, Col B = แบรนด์, Col C = ชื่อ
            if len(row_texts) > 2:
                col_c_raw = row_texts[2]
                alt_brand = clean_brand_text(col_b_raw)
                alt_name = clean_name_text(col_c_raw)
                if alt_brand and alt_name and alt_brand.upper() not in ["NONE", "NULL", "-"] and alt_name.upper() not in ["NONE", "NULL", "-"]:
                    assigned_dept = current_dept or fallback_dept or sheet_name.strip()
                    all_employees.append({
                        "name": alt_name,
                        "brand": alt_brand,
                        "department": assigned_dept
                    })
                    brand_counts[alt_brand] = brand_counts.get(alt_brand, 0) + 1
                    if assigned_dept and assigned_dept not in departments_list:
                        departments_list.append(assigned_dept)
                    continue

        # 3. ดึงข้อมูล "หมายเหตุ" และ "กะการทำงาน" จากชีต
        for row in sheet.iter_rows(values_only=True):
            if not row:
                continue
            for cell in row:
                if cell is not None:
                    cell_str = str(cell).strip()
                    if cell_str and any(k in cell_str for k in SHIFT_KEYWORDS):
                        if cell_str not in remarks:
                            remarks.append(cell_str)

    # 4. กำหนดสิทธิ์เลือกกะ: แบรนด์ไหนมีพนักงานมากกว่าเกณฑ์ (min_staff_threshold)
    for emp in all_employees:
        emp["requires_shift_selection"] = brand_counts.get(emp["brand"], 0) > min_staff_threshold

    # 5. รวบรวมรายชื่อแผนกที่ไม่ซ้ำกันทั้งหมด (คงลำดับที่พบก่อนหลัง)
    unique_departments = []
    for d in departments_list:
        if d and d not in unique_departments:
            unique_departments.append(d)
    for emp in all_employees:
        d = emp.get("department")
        if d and d not in unique_departments:
            unique_departments.append(d)

    return {
        "total_employees": len(all_employees),
        "employees": all_employees,
        "departments": unique_departments,
        "brand_counts": brand_counts,
        "remarks": remarks
    }
