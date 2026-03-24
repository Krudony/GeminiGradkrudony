import zipfile
import shutil
import os
import re
import argparse
from datetime import date, timedelta
from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# Config for Attendance
STUDENT_ROWS = [8, 9, 10]
MARK_PRESENT = 141 # Shared string ID for '/'
PERIOD_VAL = 1
MONTHS_TH = {
    4: "เมษายน", 5: "พฤษภาคม", 6: "มิถุนายน", 7: "กรกฎาคม", 8: "สิงหาคม", 9: "กันยายน", 
    10: "ตุลาคม", 11: "พฤศจิกายน", 12: "ธันวาคม", 1: "มกราคม", 2: "กุมภาพันธ์", 3: "มีนาคม"
}

def col_to_num(col):
    num = 0
    for c in col:
        if c.isalpha():
            num = num * 26 + (ord(c.upper()) - ord('A') + 1)
    return num

def num_to_col_letter(n):
    res = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res = chr(65 + rem) + res
    return res

def ensure_cell(row_elem, col_letter, row_num):
    target_ref = f'{col_letter}{row_num}'
    cells = row_elem.findall(f'{{{NS}}}c')
    for c in cells:
        if c.get('r') == target_ref:
            return c
    new_c = etree.Element(f'{{{NS}}}c')
    new_c.set('r', target_ref)
    col_num = col_to_num(col_letter)
    insert_pos = len(cells)
    for i, c in enumerate(cells):
        existing_col = ''.join([ch for ch in c.get('r', '') if ch.isalpha()])
        if col_to_num(existing_col) > col_num:
            insert_pos = i
            break
    row_elem.insert(insert_pos, new_c)
    return new_c

def set_val(row_elem, col_letter, row_num, value, val_type=None):
    if row_elem is None: return
    c = ensure_cell(row_elem, col_letter, row_num)
    if val_type == 'str':
        c.set('t', 'str')
    elif val_type == 's':
        c.set('t', 's')
    elif 't' in c.attrib:
        del c.attrib['t']
    for child in list(c):
        tag = child.tag.split('}')[-1]
        if tag in ['v', 'is', 'f']:
            if tag != 'f':
                c.remove(child)
    v = c.find(f'{{{NS}}}v')
    if v is None:
        v = etree.SubElement(c, f'{{{NS}}}v')
    v.text = str(value)

def clear_val(row_elem, col_letter, row_num):
    if row_elem is None: return
    target_ref = f'{col_letter}{row_num}'
    c = row_elem.find(f'{{{NS}}}c[@r="{target_ref}"]')
    if c is not None:
        v = c.find(f'{{{NS}}}v')
        if v is not None: v.text = ""

def _repack(file_path, extract_dir):
    calc = os.path.join(extract_dir, 'xl', 'calcChain.xml')
    if os.path.exists(calc):
        os.remove(calc)
        ct_path = os.path.join(extract_dir, '[Content_Types].xml')
        if os.path.exists(ct_path):
            with open(ct_path, 'r', encoding='utf-8') as f:
                ct = f.read()
            ct = re.sub(r'<Override[^>]+calcChain[^>]+/>', '', ct)
            with open(ct_path, 'w', encoding='utf-8') as f:
                f.write(ct)

    def zipdir(path, ziph):
        for rd, dd, ff in os.walk(path):
            for f in ff:
                fp = os.path.join(rd, f)
                ziph.write(fp, os.path.relpath(fp, path))

    backup_path = file_path.replace('.xlsx', '_backup.xlsx')
    if not os.path.exists(backup_path):
        shutil.copy2(file_path, backup_path)

    with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipdir(extract_dir, zipf)
    shutil.rmtree(extract_dir)

def detect_students(file_path, sheet_file='sheet8.xml'):
    students = []
    with zipfile.ZipFile(file_path, 'r') as z:
        try:
            data = z.read(f'xl/worksheets/{sheet_file}')
            tree = etree.fromstring(data)
            ns = {'x': NS}
            rows = tree.findall('.//x:row', ns)
            for row in rows:
                rn = int(row.get('r', 0))
                if rn < 8: continue
                for c in row.findall('.//x:c', ns):
                    col = ''.join([ch for ch in c.get('r', '') if ch.isalpha()])
                    if col == 'C':
                        f = c.find(f'{{{NS}}}f')
                        v = c.find(f'{{{NS}}}v')
                        if f is not None and v is not None and v.text:
                            students.append(rn)
        except Exception as e:
            print(f"Error detecting students in {sheet_file}: {e}")
    return students

def update_main_sheet(file_path, updates):
    extract_dir = 'xlsx_tmp_main'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z: z.extractall(extract_dir)
    sheet_path = f'{extract_dir}/xl/worksheets/sheet1.xml'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}

    def is_number(val):
        try: float(str(val)); return True
        except: return False

    for cell_ref, new_val in updates.items():
        rn = int(''.join([c for c in cell_ref if c.isdigit()]))
        col = ''.join([c for c in cell_ref if c.isalpha()])
        row = row_by_num.get(rn)
        if row is None: continue
        set_val(row, col, rn, new_val, val_type=None if is_number(new_val) else 'str')

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)
    print(f'✅ Main sheet updated')

def fill_score_sem2(file_path, scores_list):
    """Primary Sem 2 Score filling"""
    extract_dir = 'xlsx_tmp_score2'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z: z.extractall(extract_dir)
    sheet_path = f'{extract_dir}/xl/worksheets/sheet8.xml'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}
    indicator_cols = ['BJ','BK','BL','BM','BN','BO','BP','BQ']

    row7 = row_by_num.get(7)
    if row7:
        for col in indicator_cols: set_val(row7, col, 7, 10)
        set_val(row7, 'DH', 7, 80); set_val(row7, 'DN', 7, 100)
        set_val(row7, 'DP', 7, 100); set_val(row7, 'DQ', 7, 200)

    for idx, excel_row in enumerate(range(8, 8 + len(scores_list))):
        row = row_by_num.get(excel_row)
        if not row or idx >= len(scores_list): continue
        data = scores_list[idx]
        bh, di, inds = data['bh'], data['di'], data['indicators']
        dh = sum(inds)
        bg = 0
        for c in row.findall(f'{{{NS}}}c'):
            if c.get('r') == f'BG{excel_row}':
                v = c.find(f'{{{NS}}}v')
                if v is not None: bg = int(float(v.text))
        dj = di; dm = dj; dn = dh + dm; do_ = bg + bh; dq = do_ + dn
        dr = round(dq / 200 * 100)

        set_val(row, 'BH', excel_row, bh)
        for ci, col in enumerate(indicator_cols): set_val(row, col, excel_row, inds[ci])
        set_val(row, 'DI', excel_row, di)
        set_val(row, 'DH', excel_row, dh); set_val(row, 'DJ', excel_row, dj)
        set_val(row, 'DM', excel_row, dm); set_val(row, 'DN', excel_row, dn)
        set_val(row, 'DO', excel_row, do_); set_val(row, 'DP', excel_row, dn)
        set_val(row, 'DQ', excel_row, dq); set_val(row, 'DR', excel_row, dr)

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)
    print(f'✅ Scores Sem 2 updated')

def _fill_sheet_matrix(file_path, sheet_file, input_cols, scores_matrix):
    extract_dir = 'xlsx_tmp_matrix'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z: z.extractall(extract_dir)
    sheet_path = f'{extract_dir}/xl/worksheets/{sheet_file}'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}

    for idx, excel_row in enumerate(range(8, 8 + len(scores_matrix))):
        row = row_by_num.get(excel_row)
        if not row or idx >= len(scores_matrix): continue
        for ci, col in enumerate(input_cols):
            set_val(row, col, excel_row, scores_matrix[idx][ci])

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)

def fill_kun_sheet(file_path, kun_scores, level='primary'):
    sheet = 'sheet9.xml' if level=='primary' else 'sheet7.xml'
    cols = ['H','I','J','K','L','M','N','O']
    _fill_sheet_matrix(file_path, sheet, cols, kun_scores)
    print(f'✅ Kun sheet updated')

def fill_read_sheet(file_path, read_scores, level='primary'):
    sheet = 'sheet10.xml' if level=='primary' else 'sheet8.xml'
    cols = ['H','I','J','K','L']
    _fill_sheet_matrix(file_path, sheet, cols, read_scores)
    print(f'✅ Read sheet updated')

def fill_cap_sheet(file_path, cap_sem1, cap_sem2=None, level='primary'):
    sheet = 'sheet11.xml' if level=='primary' else 'sheet9.xml'
    extract_dir = 'xlsx_tmp_cap'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z: z.extractall(extract_dir)
    sheet_path = f'{extract_dir}/xl/worksheets/{sheet}'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    rows = root.findall('.//x:row', ns)
    row_by_num = {int(r.get('r', 0)): r for r in rows}

    sem1_cols = ['H','L','P','T','X'] if level=='primary' else ['H','J','L','N','P']
    sem2_cols = ['I','M','Q','U','Y'] if level=='primary' else []

    for idx, excel_row in enumerate(range(8, 8 + len(cap_sem1))):
        row = row_by_num.get(excel_row)
        if not row: continue
        for ci, col in enumerate(sem1_cols):
            set_val(row, col, excel_row, cap_sem1[idx][ci])
        if cap_sem2 and sem2_cols:
            for ci, col in enumerate(sem2_cols):
                set_val(row, col, excel_row, cap_sem2[idx][ci])

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)
    print(f'✅ Cap sheet updated')

def fill_attendance_surgical(file_path, sheet_xml_name, start_monday, term_start, term_end, teach_weekday_idx, period_val, mark_id, holidays, student_rows):   
    """Advanced Surgical Attendance Fix: Full Week Sequencing"""
    extract_dir = 'xlsx_tmp_att'
    if os.path.exists(extract_dir): shutil.rmtree(extract_dir)
    with zipfile.ZipFile(file_path, 'r') as z: z.extractall(extract_dir)
    
    sheet_path = f'{extract_dir}/xl/worksheets/{sheet_xml_name}'
    tree = etree.parse(sheet_path, etree.XMLParser(remove_blank_text=False))
    root = tree.getroot()
    ns = {'x': NS}
    
    rows = {rn: root.find(f'.//x:row[@r="{rn}"]', ns) for rn in [4, 6, 7] + student_rows}
    
    curr_mon = start_monday
    for w in range(22):
        start_col_idx = 8 + (w * 6) # Col H starts at index 8
        # Set Month string in the first column of the week
        set_val(rows[4], num_to_col_letter(start_col_idx), 4, MONTHS_TH[curr_mon.month], val_type='str')
        
        # 5-day week sequencing
        for d_off in range(5):
            col_let = num_to_col_letter(start_col_idx + d_off)
            d_obj = curr_mon + timedelta(days=d_off)
            
            # Write Date
            set_val(rows[6], col_let, 6, d_obj.day)
            
            # Write Marks ONLY on the teaching day, inside term, and not a holiday
            is_teaching_day = (d_off == teach_weekday_idx)
            is_in_term = (term_start <= d_obj <= term_end)
            is_holiday = (d_obj in holidays)
            
            if is_teaching_day and is_in_term and not is_holiday:
                set_val(rows[7], col_let, 7, period_val)
                for rn in student_rows:
                    set_val(rows[rn], col_let, rn, mark_id, val_type='s')
            else:
                # Clear existing marks to prevent legacy data jumps
                clear_val(rows[7], col_let, 7)
                for rn in student_rows:
                    clear_val(rows[rn], col_let, rn)
                    
        curr_mon += timedelta(days=7)

    tree.write(sheet_path, encoding='utf-8', xml_declaration=True)
    _repack(file_path, extract_dir)
    print(f'✅ Attendance for {sheet_xml_name} updated surgically.')

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Safe Edit Excel Grade")
    parser.add_argument('file', help="Path to the Excel file")
    args = parser.parse_args()
    print("Use this module by importing its functions.")
