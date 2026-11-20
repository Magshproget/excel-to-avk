import os
import sys
import csv
import re
from typing import List, Dict, Any, Tuple

def clean_number(val: Any) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(' ', '').replace('\xa0', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0

def clean_text(val: Any) -> str:
    if val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r'\s+', ' ', s)
    s = s.replace('*', '')  # AVK-5 does not allow * in names
    return s

def normalize_unit(unit: str) -> str:
    u = clean_text(unit).lower().strip('.')
    mapping = {
        'м пог': 'м',
        'м.пог': 'м',
        'м.п': 'м',
        'м п': 'м',
        'пог. м': 'м',
        'м куб': 'м3',
        'м.куб': 'м3',
        'м3': 'м3',
        'м кв': 'м2',
        'м.кв': 'м2',
        'м2': 'м2',
        'шт': 'шт',
        'кг': 'кг',
        'л': 'л',
        'т': 'т',
        'уп': 'компл',
        'упак': 'компл',
        'компл': 'компл',
        'бал': 'шт',
        'балон': 'шт',
        'бандл': 'шт',
        'пляш': 'шт',
        'бутилка': 'шт',
    }
    return mapping.get(u, u if u else 'шт')

def detect_columns(header_row: List[str]) -> Dict[str, int]:
    indices = {}
    for idx, cell in enumerate(header_row):
        c = str(cell).lower().strip()
        if not c:
            continue
        if ('товар' in c or 'найменування' in c or 'назва' in c or 'матеріал' in c or 'номенклатура' in c) and 'name' not in indices:
            indices['name'] = idx
        elif ('од' in c or 'вимір' in c) and 'unit' not in indices:
            indices['unit'] = idx
        elif ('кіл' in c or 'обсяг' in c or 'кол' in c or 'qty' in c) and 'облік' not in c and 'qty' not in indices:
            indices['qty'] = idx
        elif ('кіл' in c or 'кол' in c) and 'qty' not in indices:
            indices['qty'] = idx
        elif ('ціна' in c or 'цена' in c or 'price' in c) and 'ціна за' not in c and 'price' not in indices:
            indices['price'] = idx
        elif ('сума' in c or 'стоимость' in c or 'сумма' in c) and 'облік' not in c and 'sum' not in indices:
            indices['sum'] = idx
        elif '№' in c or 'n ' in c or 'nп' in c or 'код' in c or 'арт' in c:
            if 'num' not in indices:
                indices['num'] = idx
    return indices

def parse_rows(rows: List[List[Any]]) -> Tuple[List[Dict[str, Any]], Dict[str, int]]:
    header_indices = {}
    header_row_idx = -1
    
    # 1. Look for header row
    for r_idx, row in enumerate(rows):
        if not row:
            continue
        str_row = " ".join([str(c).lower() for c in row if c is not None])
        if ('товар' in str_row or 'найменування' in str_row or 'назва' in str_row or 'матеріал' in str_row) and \
           ('кіл' in str_row or 'од' in str_row or 'ціна' in str_row or 'сума' in str_row):
            header_indices = detect_columns([str(c) if c is not None else "" for c in row])
            if 'name' in header_indices and ('qty' in header_indices or 'price' in header_indices):
                header_row_idx = r_idx
                break
                
    if header_row_idx == -1:
        # Fallback heuristic: check row with most non-empty string headers
        for r_idx in range(min(15, len(rows))):
            cols = detect_columns([str(c) if c is not None else "" for c in rows[r_idx]])
            if 'name' in cols:
                header_indices = cols
                header_row_idx = r_idx
                break
                
    if header_row_idx == -1 or 'name' not in header_indices:
        raise ValueError("Не вдалося автоматично знайти стовпчики таблиці (Товар/Найменування, Кількість, Ціна). Перевірте заголовок таблиці.")
        
    items = []
    ignore_words = {'відпустив', 'прийняв', 'всього', 'разом', 'сума', 'організація:', 'склад:', 'підсумок', 'передав'}
    
    for row in rows[header_row_idx + 1:]:
        if not row:
            continue
        if len(row) <= header_indices['name']:
            continue
            
        name_val = row[header_indices['name']]
        name = clean_text(name_val)
        if not name:
            continue
        if name.lower() in ignore_words or any(name.lower().startswith(w) for w in ignore_words):
            continue
        # Skip rows that are purely numeric or summary markers
        if name.replace(' ', '').replace('.', '').replace(',', '').isdigit():
            continue
            
        qty = 0.0
        if 'qty' in header_indices and len(row) > header_indices['qty']:
            qty = clean_number(row[header_indices['qty']])
            
        unit = 'шт'
        if 'unit' in header_indices and len(row) > header_indices['unit']:
            raw_unit = str(row[header_indices['unit']]) if row[header_indices['unit']] is not None else ""
            unit = normalize_unit(raw_unit)
            
        price = 0.0
        if 'price' in header_indices and len(row) > header_indices['price']:
            price = clean_number(row[header_indices['price']])
            
        total = 0.0
        if 'sum' in header_indices and len(row) > header_indices['sum']:
            total = clean_number(row[header_indices['sum']])
        if total == 0.0 and qty > 0 and price > 0:
            total = round(qty * price, 2)
            
        if qty > 0 or price > 0:
            items.append({
                'name': name,
                'unit': unit,
                'qty': qty,
                'price': price,
                'sum': total
            })
            
    return items, header_indices

def read_file(filepath: str) -> List[Dict[str, Any]]:
    ext = os.path.splitext(filepath)[1].lower()
    
    if ext in ['.xlsx', '.xlsm']:
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active
        rows = []
        for r in sheet.iter_rows(values_only=True):
            rows.append(list(r))
        items, _ = parse_rows(rows)
        return items
        
    elif ext in ['.csv', '.txt']:
        # Try UTF-8 then CP1251
        content = None
        for enc in ['utf-8-sig', 'utf-8', 'cp1251']:
            try:
                with open(filepath, 'r', encoding=enc) as f:
                    content = f.read()
                break
            except UnicodeDecodeError:
                continue
        if content is None:
            raise ValueError("Не вдалося визначити кодування текстового файлу (спробуйте UTF-8 або Windows-1251)")
            
        reader = csv.reader(content.splitlines())
        rows = [list(r) for r in reader]
        items, _ = parse_rows(rows)
        return items
    else:
        raise ValueError(f"Непідтримуваний формат файлу: {ext}. Використовуйте .xlsx або .csv")

def export_to_avk_excel(items: List[Dict[str, Any]], output_path: str, cipher_prefix: str = "С1-ПР1-"):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "АВК_Матеріали"
    
    # AVK standard columns for Local Estimate import:
    headers = [
        "Шифр",
        "Найменування",
        "Одиниця виміру",
        "Кількість",
        "Ціна, грн",
        "Маса, т",
        "Сума, грн"
    ]
    
    ws.append(headers)
    
    # Styling headers
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for idx, item in enumerate(items, 1):
        cipher = f"{cipher_prefix}{idx}" if cipher_prefix else str(idx)
        ws.append([
            cipher,
            item['name'],
            item['unit'],
            item['qty'],
            item['price'],
            0.000, # Mass in tons (mandatory for AVK)
            item['sum']
        ])
        
    # Auto-fit column widths
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 60)
        
    wb.save(output_path)

def export_to_avk_csv(items: List[Dict[str, Any]], output_path: str, cipher_prefix: str = "С1-ПР1-"):
    headers = [
        "Шифр",
        "Найменування",
        "Одиниця виміру",
        "Кількість",
        "Ціна",
        "Маса",
        "Сума"
    ]
    with open(output_path, 'w', encoding='cp1251', newline='', errors='replace') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(headers)
        for idx, item in enumerate(items, 1):
            cipher = f"{cipher_prefix}{idx}" if cipher_prefix else str(idx)
            writer.writerow([
                cipher,
                item['name'],
                item['unit'],
                str(item['qty']).replace('.', ','),
                str(item['price']).replace('.', ','),
                '0,000',
                str(item['sum']).replace('.', ',')
            ])

if __name__ == '__main__':
    if len(sys.argv) > 1:
        inp = sys.argv[1]
        out_xlsx = sys.argv[2] if len(sys.argv) > 2 else "avk_output.xlsx"
        print(f"Reading {inp}...")
        parsed = read_file(inp)
        print(f"Found {len(parsed)} items.")
        export_to_avk_excel(parsed, out_xlsx)
        print(f"Saved to {out_xlsx} successfully!")
