"""
Excel filler — reads extracted JSON records and fills template Excel
"""
import openpyxl, json, os, re, yaml
from copy import copy


def load_records(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)


def fill_template(records, template_path, output_path, config_path=None):
    """
    Fill extracted records into the template Excel.
    Matches by 野外编号 (column C) — if existing row found, update it.
    Otherwise, append as new row.
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb["Sheet1"]
    
    # Column mapping from config
    col_map = {
        "野外编号": "C", "统一编号": "B", "路线编号": "D", "日期": "E",
        "地理位置": "F", "天气": "G",
        "井台高度": "P", "井深": "Q", "测点距地面高度": "U",
        "地下水位埋深": "V", "测点距水面距离": "W",
        "井点类型": "T", "所属统测区类型": "S",
        "统测期": "AB", "统测期次类型": "AC",
        "含水介质": "Y", "埋藏深度": "Z", "承压性": "AA",
        "测量人": "AH", "记录人": "AI", "审核人": "AJ",
        "统测实施单位": "AG", "项目名称": "AF",
        "审核人": "AJ",
    }
    
    updated = 0
    added = 0
    
    for record in records:
        field_id = record.get("野外编号")
        if not field_id:
            continue
        
        # Find existing row by 野外编号 (column C)
        target_row = None
        for row in range(2, ws.max_row + 1):
            if ws[f"C{row}"].value and str(ws[f"C{row}"].value).strip() == field_id:
                target_row = row
                break
        
        if target_row is None:
            # Find first empty row
            target_row = ws.max_row + 1
            added += 1
        else:
            updated += 1
        
        # Fill cells
        for field_key, col_letter in col_map.items():
            value = record.get(field_key)
            if value:
                cell = ws[f"{col_letter}{target_row}"]
                
                # Convert date format
                if field_key == "日期":
                    value = _normalize_date(value)
                
                # Convert numeric
                if field_key in ("井台高度", "井深", "测点距地面高度", "地下水位埋深", "测点距水面距离"):
                    try:
                        value = float(value)
                    except:
                        pass
                
                cell.value = value
    
    wb.save(output_path)
    print(f"Updated: {updated} rows, Added: {added} rows")
    print(f"Saved to: {output_path}")
    return updated, added


def _normalize_date(val):
    """Normalize various date formats to YYYYMMDD"""
    val = str(val).strip()
    # 2026-05-21 → 20260521
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})', val)
    if m:
        return m.group(1) + m.group(2) + m.group(3)
    # 20260521 → 20260521
    m = re.match(r'(\d{8})', val)
    if m:
        return val
    # 2026年5月21日 → 20260521
    m = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', val)
    if m:
        return f"{m.group(1)}{int(m.group(2)):02d}{int(m.group(3)):02d}"
    return val


if __name__ == "__main__":
    import sys
    
    json_path = sys.argv[1] if len(sys.argv) > 1 else r"D:\hermes-workspace\projects\well-extractor\output\extracted.json"
    template = sys.argv[2] if len(sys.argv) > 2 else r"D:\地下水\大同朔州忻州_地下水统测表_补全_最终版.xlsx"
    output = sys.argv[3] if len(sys.argv) > 3 else r"D:\hermes-workspace\projects\well-extractor\output\filled_output.xlsx"
    
    records = load_records(json_path)
    print(f"Loaded {len(records)} records")
    fill_template(records, template, output)
