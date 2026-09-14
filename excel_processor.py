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
        {"no": 6, "brand": "Brand Alpha (สาขา สยาม)", "name": "ดารินทร์ สุขใจ (หัวหน้า)"}
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
