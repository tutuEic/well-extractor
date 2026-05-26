"""
Enhanced OCR for groundwater survey screenshots
- Preprocessing for better numeric extraction
- Field-aware parsing (label: value patterns)
"""
import pytesseract, re, os, yaml
from PIL import Image, ImageEnhance, ImageFilter
from collections import OrderedDict

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def load_config():
    config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "fields.yaml")
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def preprocess_image(img):
    """Enhance image for better OCR"""
    # Convert to grayscale if needed
    if img.mode != "L":
        img = img.convert("L")
    
    # Increase contrast
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.0)
    
    # Sharpen
    img = img.filter(ImageFilter.SHARPEN)
    
    # Binarize (threshold)
    img = img.point(lambda x: 0 if x < 140 else 255)
    
    return img


def extract_from_image(image_path):
    """
    Extract all fields from a single survey screenshot.
    Returns dict of field_name → value.
    """
    img = Image.open(image_path)
    
    # Split long screenshot into segments
    # Each segment = one form section
    chunk_h = 600
    overlap = 50
    all_text = []
    
    for y in range(0, img.height, chunk_h - overlap):
        chunk = img.crop((0, y, img.width, min(y + chunk_h, img.height)))
        
        # Try both original and preprocessed
        for proc_img, label in [(chunk, "orig"), (preprocess_image(chunk), "proc")]:
            text = pytesseract.image_to_string(proc_img, lang="chi_sim+eng", config="--psm 6")
            if text.strip():
                all_text.append(text)
                break
    
    full_text = "\n".join(all_text)
    
    # Parse into structured fields
    return parse_fields(full_text)


def parse_fields(text):
    """Parse OCR text into structured field:value pairs"""
    config = load_config()
    field_map = config["field_mapping"]
    
    result = {}
    
    # Clean up text
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    
    # Pattern 1: * 标签 : 值    (common in the app)
    for line in lines:
        match = re.match(r'[\*\s]*([^:：]+)[：:]\s*(.+)', line)
        if match:
            label = match.group(1).strip().replace(" ", "")
            value = match.group(2).strip()
            _try_match(result, label, value, field_map)
    
    # Pattern 2: 标签  value on next line (OCR sometimes splits)
    for i, line in enumerate(lines):
        for label_key in field_map:
            clean_label = label_key.replace(" ", "")
            if clean_label in line.replace(" ", "") and i + 1 < len(lines):
                # Check if current line is just a label (no value after colon)
                if re.match(r'[\*\s]*[^：:]+$', line) or clean_label in line:
                    next_line = lines[i + 1].strip()
                    # Only take if it looks like a value (not another label)
                    if not any(k.replace(" ", "") in next_line.replace(" ", "") for k in field_map):
                        _try_match(result, label_key, next_line, field_map)
    
    # Post-process: extract numeric values
    _extract_numerics(result, text, config)
    
    # Apply defaults
    for key, val in config.get("defaults", {}).items():
        col = field_map.get(key)
        if col and col not in result:
            result[col] = val
    
    return result


def _try_match(result, label, value, field_map):
    """Try to match a label to template column"""
    # Clean label
    clean = label.replace(" ", "").replace("*", "").replace("（", "(").replace("）", ")")
    
    # Direct match
    for key, col in field_map.items():
        key_clean = key.replace(" ", "").replace("（", "(").replace("）", ")")
        if clean == key_clean or key_clean in clean or clean in key_clean:
            # Skip coordinates
            if key in ("经度", "纬度"):
                return
            
            result[col] = value.strip()
            return
    
    # Fuzzy match
    for key, col in field_map.items():
        if len(clean) >= 3 and len(key) >= 3:
            if clean[:3] == key[:3]:
                result[col] = value.strip()
                return


def _extract_numerics(result, text, config):
    """Try harder to find numeric values for known fields"""
    numeric_keys = {
        "井台高度": ["井台高度", "井台"],
        "井深": ["井深"],
        "测点距地面高度": ["距地面高度", "距地面"],
        "地下水位埋深": ["埋深", "水位埋深", "地下埋深"],
        "测点距水面距离": ["距水面距离", "距水面", "水面距离"],
    }
    
    for field_name, patterns in numeric_keys.items():
        if field_name in result:
            continue
        
        # Search for pattern like: "井台高度 1.8" or "井台高度(m):1.8"
        for pat in patterns:
            match = re.search(rf'{pat}[^0-9]*(\d+\.?\d*)', text)
            if match:
                val = float(match.group(1))
                if 0 < val < 1000:  # Sanity check
                    col = config["field_mapping"].get(field_name)
                    if col:
                        result[col] = str(val)
                    break


def extract_batch(image_dir, output_path=None):
    """Extract all images in a directory, return list of records"""
    import glob, json
    
    images = sorted(glob.glob(os.path.join(image_dir, "*.jpg")))
    records = []
    
    for i, img_path in enumerate(images):
        print(f"[{i+1}/{len(images)}] {os.path.basename(img_path)[:50]}...")
        record = extract_from_image(img_path)
        record["_source"] = os.path.basename(img_path)
        records.append(record)
        
        # Print summary
        fields_found = [k for k, v in record.items() if not k.startswith("_") and v]
        print(f"  Extracted {len(fields_found)} fields: {fields_found[:8]}...")
    
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        print(f"\nSaved {len(records)} records to {output_path}")
    
    return records


if __name__ == "__main__":
    import sys
    img_dir = sys.argv[1] if len(sys.argv) > 1 else r"D:\地下水\信息截图\5.21"
    out = sys.argv[2] if len(sys.argv) > 2 else r"D:\hermes-workspace\projects\well-extractor\output\extracted.json"
    
    records = extract_batch(img_dir, out)
    
    print(f"\n{'='*60}")
    print("Sample record:")
    for r in records[:1]:
        for k, v in r.items():
            if not k.startswith("_") and v:
                print(f"  {k}: {v}")
