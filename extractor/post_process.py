"""
Post-processor — cleans OCR noise from extracted records
"""
import re, json, os
from datetime import datetime

KNOWN_CORRUPTIONS = {
    # OCR noise → correct value
    "一般统测区国": "一般统测区",
    "一航统测区国": "一般统测区",
    "一般统测区": "一般统测区",
    "国": "",
    "日": "",
    "吴": "",
    "口": "",
    "园": "",
    "阴一": "阴",
    "阴国": "阴",
    "杨光国": "杨光",
    "杨升伟国": "杨升伟",
    "杨升伟园": "杨升伟",
    "eo\"\"": "地下水平水位期",
    "eo\"": "地下水平水位期",
    "类型": "一期统测",
    "请选择审核人": "",
    "Aes": "山西省地质调查院有限公司",
    "226-5-21": "20260521",
    "期统测团": "一期统测",
    "期统测\\": "一期统测",
    "期统测\"": "一期统测",
    "杨光~": "杨光",
    "杨光-": "杨光",
    "杨光\\": "杨光",
    "杨升伟~": "杨升伟",
    "杨升伟\\": "杨升伟",
    "一般统测区\\": "一般统测区",
    '一般统测区"': "一般统测区",
}

def clean_value(val):
    """Clean a single value"""
    if not val:
        return val
    
    # Direct replacements
    for bad, good in KNOWN_CORRUPTIONS.items():
        val = val.replace(bad, good)
    
    # Remove standalone garbage chars
    val = re.sub(r'[国日吴口园]', '', val)
    
    # Clean project name (common pattern)
    if '山西省' in val and '水位统测' in val:
        # Extract: "山西省地下水水位统测" + optional suffix
        match = re.search(r'(山西省地下水水位统测)', val)
        if match:
            val = match.group(1)
    
    # Fix trailing garbage — but only after CJK chars, NOT after digits
    val = re.sub(r'([\u4e00-\u9fff])[eEoO0]+[\"\'\"\']*\s*$', r'\1', val)
    val = re.sub(r'士也下7J[^一]*$', '', val)
    
    # Strip
    val = val.strip()
    
    # Remove if empty after cleaning
    if not val or val in ('/', '无', '请选择'):
        return None
    
    return val


def normalize_record(record):
    """Clean an entire record"""
    cleaned = {}
    for key, val in record.items():
        if key.startswith("_"):
            cleaned[key] = val
            continue
        
        if isinstance(val, str):
            new_val = clean_value(val)
            if new_val:
                cleaned[key] = new_val
        elif val is not None:
            cleaned[key] = val
    
    # Fallback: try to extract 野外编号 from image filename or context
    # The screenshots have file IDs — we may need manual mapping
    
    return cleaned


def process_json(input_path, output_path=None):
    with open(input_path, "r", encoding="utf-8") as f:
        records = json.load(f)
    
    cleaned = [normalize_record(r) for r in records]
    
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(cleaned, f, ensure_ascii=False, indent=2)
    
    return cleaned


if __name__ == "__main__":
    in_path = r"D:\hermes-workspace\projects\well-extractor\output\extracted.json"
    out_path = in_path.replace(".json", "_cleaned.json")
    
    records = process_json(in_path, out_path)
    
    print("Cleaned records:")
    for i, r in enumerate(records):
        print(f"\n[{i+1}]")
        for k, v in r.items():
            if not k.startswith("_"):
                print(f"  {k}: {v}")
