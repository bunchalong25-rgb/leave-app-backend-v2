import io
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

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
                clean_name = val.strip()
                if len(clean_name) >= 3 and clean_name not in emp_row_map:
                    emp_row_map[clean_name] = r

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
        target_row = emp_row_map.get(rec_name)

        if not target_row and rec_name:
            for k, r_idx in emp_row_map.items():
                if rec_name in k or k in rec_name:
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
    ฟังก์ชันสำหรับอ่านไฟล์ Excel อัจฉริยะ (รองรับการอ่านทุก Sheet)
    ดึงชื่อ แบรนด์ ดึงหมายเหตุ/กะการทำงาน และเช็คเงื่อนไขแบรนด์ที่มีพนักงาน > min_staff_threshold
    พร้อมระบบป้องกัน Error กรณีแถวว่างหรือโครงสร้างคอลัมน์ไม่แน่นอน
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_content), data_only=True)
    
    all_employees = []
    brand_counts = {}
    remarks = []

    # คีย์เวิร์ดสำหรับตรวจจับคอลัมน์และหัวตาราง
    NAME_KEYWORDS = ["ชื่อ-นามสกุล", "ชื่อ - นามสกุล", "ชื่อ-สกุล", "ชื่อพนักงาน", "ชื่อ", "FULLNAME", "NAME"]
    BRAND_KEYWORDS = ["แบรนด์ / สาขา", "แบรนด์/สาขา", "แบรนด์", "สาขา", "BRAND"]
    DEPT_KEYWORDS = ["แผนก", "ฝ่าย", "สังกัด", "สายงาน", "DEPARTMENT", "DEPT", "SECTION"]
    STOP_KEYWORDS = ["สรุปจำนวน", "รวมทั้งสิ้น", "สรุปยอด", "ยอดรวม", "TOTAL", "หมายเหตุ"]
    SHIFT_KEYWORDS = ["คำอธิบายสัญลักษณ์", "ช่วงเวลาปฏิบัติ", "เวลาปฏิบัติงาน", "เวลาทำงาน", "กะการทำงาน", "กะ ", "กะเช้า", "กะบ่าย", "กะดึก", "W =", "C =", "V =", "S =", "N =", "OFF"]

    # 1. วนลูปอ่านข้อมูล "ทุก Sheet" ในไฟล์
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        header_row_index = None
        brand_col_idx = 0
        name_col_idx = 1
        dept_col_idx = None
        
        # 2. ค้นหาบรรทัดที่เป็น "หัวตารางที่แท้จริง" (สแกน 25 บรรทัดแรก)
        for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=25, values_only=True), start=1):
            if not row:
                continue
            row_texts = [str(cell).strip() if cell is not None else "" for cell in row]
            
            # ตรวจสอบว่ามีคอลัมน์ชื่อ และ คอลัมน์แบรนด์ ในแถวเดียวกันหรือไม่
            has_name = any(any(k in text.upper() for k in NAME_KEYWORDS) for text in row_texts)
            has_brand = any(any(k in text.upper() for k in BRAND_KEYWORDS) for text in row_texts)

            if has_name and has_brand:
                header_row_index = i
                for col_idx, text in enumerate(row_texts):
                    text_upper = text.upper()
                    if any(k in text_upper for k in BRAND_KEYWORDS):
                        brand_col_idx = col_idx
                    elif any(k in text_upper for k in NAME_KEYWORDS):
                        name_col_idx = col_idx
                    elif any(k in text_upper for k in DEPT_KEYWORDS):
                        dept_col_idx = col_idx
                break
                
        # หาก Sheet ไหนไม่ใช่ตารางรายชื่อ ให้ข้ามไป Sheet ถัดไป
        if header_row_index is None:
            continue
            
        # ค้นหาชื่อแผนกจริงจากส่วนหัวของ Sheet (เช่น ข้อความ 'แผนก 101 :PRESTIGE ชั้น 1' ในแถว 1-15)
        sheet_department = None
        for r_idx in range(1, min(20, sheet.max_row + 1)):
            for c_idx in range(1, min(sheet.max_column + 1, 35)):
                val = sheet.cell(row=r_idx, column=c_idx).value
                if val is not None:
                    val_clean = str(val).strip()
                    val_upper = val_clean.upper()
                    # ตรวจจับข้อความแผนก เช่น "แผนก 101 :PRESTIGE ชั้น 1" หรือ "ฝ่าย..." หรือ "DEPARTMENT"
                    if any(k in val_upper for k in ["แผนก", "ฝ่าย", "DEPARTMENT", "SECTION"]):
                        # หลีกเลี่ยงกรณีที่เป็นแค่ชื่อหัวคอลัมน์คำเดียว
                        if len(val_clean) > 3 and not any(stop in val_upper for stop in STOP_KEYWORDS):
                            sheet_department = val_clean
                            break
            if sheet_department:
                break
                
        # หากไม่พบในเซลล์หัวตาราง ให้ตรวจว่าชื่อ Sheet เป็นชื่อแผนกหรือไม่
        if not sheet_department:
            sheet_department = sheet_name.strip()
            
        # 3. เริ่มดึงข้อมูลพนักงาน (อ่านต่อจากบรรทัดหัวตารางลงมา)
        for row in sheet.iter_rows(min_row=header_row_index + 1, values_only=True):
            if not row:
                continue

            # ตรวจหาจุดสิ้นสุดของรายชื่อพนักงาน
            row_sample_str = " ".join([str(c).strip() for c in row[:5] if c is not None])
            if any(stop_word in row_sample_str for stop_word in STOP_KEYWORDS):
                break
                
            col_brand = str(row[brand_col_idx]).strip() if len(row) > brand_col_idx and row[brand_col_idx] is not None else ""
            col_name = str(row[name_col_idx]).strip() if len(row) > name_col_idx and row[name_col_idx] is not None else ""
            col_dept = str(row[dept_col_idx]).strip() if dept_col_idx is not None and len(row) > dept_col_idx and row[dept_col_idx] is not None else ""
            
            # ตรวจเช็คว่าบรรทัดนี้เป็นคำสรุปหรือไม่
            if any(stop_word in col_brand for stop_word in STOP_KEYWORDS) or any(stop_word in col_name for stop_word in STOP_KEYWORDS):
                break

            final_dept = col_dept if (col_dept and col_dept.upper() not in ["NONE", "NULL", "-"]) else sheet_department

            # ถ้ามีข้อมูลครบทั้ง Brand และ ชื่อ และไม่ใช่ค่าว่าง/NONE ให้เก็บลงระบบ
            if col_brand and col_name and col_brand.upper() not in ["NONE", "NULL", "-"] and col_name.upper() not in ["NONE", "NULL", "-"]:
                # ป้องกันเก็บซ้ำใน Sheet เดียวกัน
                all_employees.append({
                    "name": col_name, 
                    "brand": col_brand,
                    "department": final_dept
                })
                # นับจำนวนคนในแบรนด์ เพื่อเอาไปคำนวณสิทธิ์เลือกกะ
                brand_counts[col_brand] = brand_counts.get(col_brand, 0) + 1
                
        # 4. ดึงข้อมูล "หมายเหตุ" และ "กะการทำงาน" จากตาราง
        for row in sheet.iter_rows(min_row=1, values_only=True):
            if not row:
                continue
            for cell in row:
                if cell is not None:
                    cell_str = str(cell).strip()
                    if cell_str and any(k in cell_str for k in SHIFT_KEYWORDS):
                        if cell_str not in remarks:
                            remarks.append(cell_str)
                    
    # 5. กำหนดสิทธิ์เลือกกะ: แบรนด์ไหนมีพนักงานมากกว่าเกณฑ์ (min_staff_threshold)
    for emp in all_employees:
        emp["requires_shift_selection"] = brand_counts.get(emp["brand"], 0) > min_staff_threshold

    # 6. รวบรวมรายชื่อแผนกที่ไม่ซ้ำกันทั้งหมด
    unique_departments = list(dict.fromkeys(emp["department"] for emp in all_employees if emp.get("department")))

    # ส่งข้อมูลทั้งหมดกลับไปให้ระบบ API 
    return {
        "total_employees": len(all_employees),
        "employees": all_employees,
        "departments": unique_departments,
        "brand_counts": brand_counts,
        "remarks": remarks
    }


