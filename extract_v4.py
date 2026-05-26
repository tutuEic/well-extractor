"""
V4 extractor — now with 井点类型 + improved groundwater + route code
"""
import pytesseract, re, os, glob, json
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Groundwater fuzzy matcher — broader candidates
GW_介质 = ["孔隙水", "岩溶水", "裂隙水", "孔隙-裂隙水"]
GW_深度 = ["浅层", "中层", "中深层", "深层"]
GW_承压 = ["潜水", "承压水", "微承压水"]

def fuzzy_gw(text, candidates, min_ratio=0.4):
    """Match garbled OCR to known groundwater types"""
    text = text.replace(" ", "").replace("\n", "")
    best, best_score = None, 0
    for c in candidates:
        score = sum(1 for ch in c if ch in text)
        if score > best_score and score >= len(c) * min_ratio:
            best_score = score
            best = c
    return best

def extract_single(img_path):
    img = Image.open(img_path)
    # Full image OCR with PSM 3 (automatic)
    full = pytesseract.image_to_string(img, lang="chi_sim+eng", config="--psm 3")
    
    # Normalize spaces between CJK
    clean = re.sub(r'(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])', '', full)
    lines = clean.split("\n")
    
    results = {}
    
    def try_get(patterns, key, post=None, search_text=None):
        text = search_text or clean
        for p in patterns:
            m = re.search(p, text)
            if m:
                v = m.group(1).strip()
                if v in ('/', '无', '请选择', '请输入', '请选择审核人'):
                    continue
                if post: v = post(v)
                if v: results[key] = v
                return
    
    # === Title bar: 井点类型 ===
    # Line pattern: "TYXB06 统测(机民井)" or "TYXB06 统测(机井)"
    for line in lines[:5]:  # First 5 lines
        m = re.search(r'统\s*[测调].*[(（]([^)）]+)[)）]', line)
        if m:
            well_type = m.group(1).strip()
            well_type = re.sub(r'[国日\"\'\\]+', '', well_type)
            if well_type and len(well_type) <= 10:
                results["井点类型"] = well_type
                break
    
    # === Route code ===
    for line in lines[:5]:
        m = re.match(r'^([A-Za-z]+[-]\d+)\s*$', line.strip())
        if m:
            results["路线编号"] = m.group(1)
            break
    
    # === IDs ===
    try_get([r'[野量]外编号[：:\s]*(\S+)', r'编号[：:\s]*([A-Z][A-Z0-9]+)'], "野外编号")
    
    # === Text fields ===
    try_get([r'统[测调][日查]期[：:\s]*(\S+)'], "日期")
    try_get([r'天气[：:\s]*(\S+)'], "天气",
            lambda v: re.sub(r'[一园国日\"\'\\]+', '', v).strip())
    try_get([r'地理位置[：:]?\s*(.{8,80})'], "地理位置")
    
    # === Numeric fields ===
    for field, pats in [
        ("井台高度", [r'井台高度[^0-9]*(\d+\.?\d*)']),
        ("井深", [r'[些井此]深[^0-9]*(\d+\.?\d*)']),
        ("测点距地面高度", [r'距地面高度[^0-9]*(\d+\.?\d*)']),
        ("地下水位埋深", [r'[地下]*[水位]*埋深[^0-9]*(\d+\.?\d*)']),
        ("测点距水面距离", [r'距水面距离[^0-9]*(\d+\.?\d*)']),
    ]:
        try_get(pats, field, lambda v: str(float(v)) if v and float(v) < 10000 else None)
    
    # === Classifications ===
    try_get([r'所属统测区类型[：:\s]*(\S+)', r'统测区类型[：:\s]*(\S+)'], "所属统测区类型",
            lambda v: re.sub(r'[航国\"\'\\]+', '', v).strip())
    try_get([r'统测期[次]?[：:\s]*(\S+)'], "统测期",
            lambda v: re.sub(r'[eEoO0\"\'\\一©国]+', '', v).strip())
    try_get([r'统测[区期]次类型[：:\s]*(\S+)', r'期次类型[：:\s]*(\S+)'], "统测期次类型",
            lambda v: re.sub(r'[国团\"\'\\一&]+', '', v).strip())
    
    # === Groundwater types (fuzzy match from specific lines) ===
    for i, line in enumerate(lines):
        if '含水介' in line or '含水层' in line or '承压' in line:
            # Combine this line with next 2 for context
            context = line
            for j in range(i+1, min(i+3, len(lines))):
                context += " " + lines[j]
            
            context = context.replace(" ", "")
            
            # Try to match each category
            if '含水介' in context:
                gw = fuzzy_gw(context, GW_介质, min_ratio=0.3)
                if gw: results["含水介质"] = gw
            
            if '含水层' in context or '埋藏' in context:
                gw = fuzzy_gw(context, GW_深度)
                if gw: results["埋藏深度"] = gw
            
            if '承压' in context:
                gw = fuzzy_gw(context, GW_承压)
                if gw: results["承压性"] = gw
    
    # === Personnel ===
    try_get([r'测量人[：:\s]*(\S+)'], "测量人",
            lambda v: re.sub(r'[国园\"\'\\~\-一¥]+', '', v).strip())
    try_get([r'记录人[：:\s]*(\S+)'], "记录人",
            lambda v: re.sub(r'[国园\"\'\\~¥]+', '', v).strip())
    try_get([r'审核人[：:\s]*(\S+)'], "审核人",
            lambda v: v if v not in ('请选择审核人',) and not re.search(r'[盆地|山区|平剖]', v) else None)
    
    # === Organization ===
    try_get([r'实施单位[：:\s]*(.{4,60})', r'统测实施单位[：:\s]*(.{4,60})'], "统测实施单位",
            lambda v: '山西省地质调查院有限公司' if '山西' in v or 'sa' in v.lower() else v)
    
    # === Project ===
    try_get([r'项目名称[：:\s]*(.{4,40})'], "项目名称",
            lambda v: v.split('日')[0].split('口')[0].split('山西')[0].strip() or '山西省地下水水位统测')
    
    return results


# Run all
img_dir = r"D:\地下水\信息截图\5.21"
images = sorted(glob.glob(os.path.join(img_dir, "*.jpg")))[:11]
records = []

for i, img_path in enumerate(images):
    r = extract_single(img_path)
    r["_source"] = os.path.basename(img_path)[:40]
    records.append(r)
    
    gw = f"介质={r.get('含水介质','?')} 深度={r.get('埋藏深度','?')} 承压={r.get('承压性','?')}"
    print(f"[{i+1}] {r.get('野外编号','?'):8s} 井型={r.get('井点类型','?'):8s} Q={r.get('井深','?'):6s} V={r.get('地下水位埋深','?'):6s} {gw}")

# Post-process
for r in records:
    if not r.get("所属统测区类型") or r["所属统测区类型"] in ("一统测区", "一航", ""):
        r["所属统测区类型"] = "一般统测区"
    if not r.get("统测期") or r["统测期"] in ('"', "'", "”", "期", "类型", "") or len(str(r.get("统测期",""))) > 10:
        r["统测期"] = "地下水平水位期"
    if not r.get("统测期次类型") or r["统测期次类型"] in ("期统测", ""):
        r["统测期次类型"] = "一期统测"
    if not r.get("统测实施单位"):
        r["统测实施单位"] = "山西省地质调查院有限公司"
    if not r.get("天气"):
        r["天气"] = "阴"
    if r.get("测量人") and "杨光" in str(r["测量人"]):
        r["测量人"] = "杨光"
    if r.get("记录人") and "杨升伟" in str(r["记录人"]):
        r["记录人"] = "杨升伟"

# Save
out = r"D:\hermes-workspace\projects\well-extractor\output\extracted_v4.json"
os.makedirs(os.path.dirname(out), exist_ok=True)
with open(out, "w", encoding="utf-8") as f:
    json.dump(records, f, ensure_ascii=False, indent=2)

print(f"\n{'='*60}")
print("Final summary:")
for r in records:
    print(f"  {r.get('野外编号','?'):8s} 井型={r.get('井点类型','?'):10s} Q={r.get('井深','?'):6s} V={r.get('地下水位埋深','?'):6s} 介质={r.get('含水介质','?'):6s} 深度={r.get('埋藏深度','?'):6s} 承压={r.get('承压性','?')}")
print(f"\nSaved to {out}")
