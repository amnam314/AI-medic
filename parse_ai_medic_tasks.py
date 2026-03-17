#!/usr/bin/env python3
"""Parse Задачи ИИ медик.xlsx and output JSON for reports."""
import xml.etree.ElementTree as ET
import json
from datetime import datetime

def excel_date(num):
    """Convert Excel serial date to ISO string."""
    if not num or num == '':
        return None
    try:
        n = float(num)
        # Excel epoch is 1899-12-30
        from datetime import timedelta
        d = datetime(1899, 12, 30) + timedelta(days=n)
        return d.strftime('%Y-%m-%d')
    except:
        return None

def parse_xlsx(path):
    import os
    base_dir = os.path.dirname(path)
    extract_dir = os.path.join(base_dir, 'xlsx_ai_medic')
    
    # Parse shared strings
    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    tree = ET.parse(f'{extract_dir}/xl/sharedStrings.xml')
    root = tree.getroot()
    strings = []
    for si in root.findall('.//main:si', ns):
        full = ''
        for elem in si.iter():
            if elem.text:
                full += elem.text
            if elem.tail:
                full += elem.tail
        strings.append(full.strip() or '-')
    
    # Parse sheet
    tree = ET.parse(f'{extract_dir}/xl/worksheets/sheet1.xml')
    root = tree.getroot()
    
    def col_to_idx(ref):
        col = ''.join(c for c in ref if c.isalpha())
        idx = 0
        for c in col:
            idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
        return idx - 1
    
    tasks = []
    for row in root.findall('.//main:row', ns):
        r = int(row.get('r', 0))
        row_data = {}
        for c in row.findall('main:c', ns):
            ref = c.get('r', '')
            col_idx = col_to_idx(ref)
            val_el = c.find('main:v', ns)
            if val_el is not None:
                v = val_el.text
                t = c.get('t')
                if t == 's':
                    idx = int(v)
                    row_data[col_idx] = strings[idx] if idx < len(strings) else v
                else:
                    row_data[col_idx] = v
        
        # Columns: 0=ID, 3=Title, 5=Created, 6=Updated, 7=Completed, 10=State
        task_id = row_data.get(0, '')
        if not task_id or 'MCAI' not in str(task_id):
            continue
        
        state = row_data.get(10, '')
        created = excel_date(row_data.get(5, ''))
        updated = excel_date(row_data.get(6, ''))
        completed = excel_date(row_data.get(7, ''))
        
        tasks.append({
            'id': task_id,
            'title': row_data.get(3, ''),
            'state': state if state and state != '-' else 'Unknown',
            'created': created,
            'updated': updated,
            'completed': completed,
            'type': row_data.get(8, ''),
            'subsystem': row_data.get(12, ''),
        })
    
    return tasks

if __name__ == '__main__':
    tasks = parse_xlsx('/Users/macbook/Cursor/Задачи ИИ медик.xlsx')
    print(json.dumps(tasks, ensure_ascii=False, indent=2))
