"""
OCR Engine v2 - PP-OCR (RapidOCR) extraction with smart post-processing.
Also supports Xiaomi MiMo vision API for high-accuracy form extraction.
"""
import os, sys, base64, json, re, functools
import numpy as np
from PIL import Image

_rapid_ocr = None

@functools.lru_cache(maxsize=1)
def _cached_rapidocr(lazy_key: int = 0):
    from rapidocr_onnxruntime import RapidOCR
    return RapidOCR()


def get_reader():
    global _rapid_ocr
    if _rapid_ocr is None:
        _rapid_ocr = _cached_rapidocr()
    return _rapid_ocr

LABEL_MAP = {
    "野外编号": "野外编号", "统一编号": "统一编号", "路线编号": "路线编号",
    "统测日期": "调查日期", "地理位置": "地理位置", "天气": "天气",
    "经度": "经度_decimal", "纬度": "纬度_decimal",
    "经度": "经度_dms", "纬度": "纬度_dms",
    "地面高程": "地面高程", "测点高程": "测点高程",
    "井台高度": "井台高度", "井深": "井深",
    "测点距水面距离": "测点距水面距离",
    "测点距地面高度": "测点距地面高度",
    "地下水位埋深": "地下水位埋深",
    "地下水埋深": "地下水位埋深",
    "所属统测区类型": "所属统测区类型",
    "统测期": "统测期",
    "统测区统测期次类型": "统测期次类型",
    "地下水类型(按含水介": "含水介质",
    "地下水类型(按含水介": "含水介质",
    "地下水类型 (按含水层": "埋藏深度",
    "地下水类型(按含水层": "埋藏深度",
    "地下水类型(按含水层": "埋藏深度",
    "地下水类型(按承压": "承压性",
    "地下水类型(按承压": "承压性",
    "地下水资源区名称": "地下水资源区名称",
    "统测实施单位": "统测实施单位",
    "测量人": "测量人", "记录人": "记录人", "审核人": "审核人",
    "井点类型": "井点类型",
    "地面高程获取方法": "地面高程获取方法",
    "测点高程获取方法": "测点高程获取方法",
}

# ── MiMo API config ─────────────────────────────────────────────────
MIMO_API_URL = "https://api.xiaomimimo.com/v1/chat/completions"
MIMO_MODEL = "mimo-v2-omni"

MIMO_EXTRACTION_PROMPT = """请仔细识别这张地下水统测调查微信小程序截图,提取所有字段和对应的值。

返回严格的 JSON 格式,键名使用以下字段名(中文字段名):

需要提取的字段:
- 野外编号
- 统一编号
- 路线编号
- 调查日期(格式: YYYY-MM-DD)
- 地理位置
- 天气
- 经度(十进制度)
- 纬度(十进制度)
- 经度_dms(度分秒格式,如 111°25'24.98")
- 纬度_dms(度分秒格式)
- 地面高程(数字,m)
- 地面高程获取方法
- 测点高程(数字,m)
- 测点高程获取方法
- 井点类型(机井/民井/大口井等)
- 井台高度(m)
- 井深(m)
- 地下水位埋深(m)
- 测点距水面距离(m)
- 测点距地面高度(m)
- 含水介质(孔隙水/裂隙水/岩溶水等)
- 埋藏深度(浅层/中层/深层等)
- 承压性(潜水/承压水等)
- 井壁结构
- 成井时间(格式: YYYY-MM-DD)
- 取水用途
- 项目名称
- 调查单位(统测实施单位)
- 调查人(测量人)
- 记录人
- 审核人
- 所属统测区类型
- 统测期
- 统测期次类型
- 地下水资源区名称
- 地理位置描述
- 县级行政编码及县名
- 图幅(1:50w)
- 图幅(1:25w)
- 图幅(1:20w)
- 图幅(1:10w)
- 图幅(1:5w)
- 备注

规则:
1. 以截图中的实际值为准,空字段返回空字符串 ""
2. 调查日期/成井时间格式化为 YYYY-MM-DD
3. 经度/纬度十进制度值保留6位小数
4. 高度/埋深/距离等数值只返回数字,不含单位
5. 只返回 JSON,不要任何解释文字

JSON 示例:
{
  "野外编号": "TDS001-1",
  "地理位置": "岚县大蛇头乡条子沟村西北700m",
  ...
}"""


def _get_mimo_api_key():
    """Read MiMo API key from local config or environment."""
    # Priority: local config file (always has the correct key for this app)
    base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
    local_cfg = os.path.join(base_dir, "mimo_config.json")
    try:
        if os.path.exists(local_cfg):
            with open(local_cfg) as f:
                cfg = json.load(f)
            api_key = cfg.get("api_key", "")
            if api_key:
                return api_key
    except Exception:
        pass

    # Fallback: environment variables
    api_key = os.environ.get("XIAOMI_API_KEY", "") or os.environ.get("MIMO_VISION_API_KEY", "")
    if api_key:
        return api_key

    # Fallback: Hermes config
    try:
        import yaml
        config_path = os.path.expanduser("~/.hermes/config.yaml")
        if os.path.exists(config_path):
            with open(config_path) as f:
                cfg = yaml.safe_load(f)
            api_key = cfg.get("auxiliary", {}).get("vision", {}).get("api_key", "")
    except Exception:
        pass
    return api_key


def extract_from_image_mimo(image_path, api_key=None, model=None):
    """
    Extract groundwater survey data from screenshot using Xiaomi MiMo vision API.

    Returns:
        dict with field:value pairs, same format as extract_from_image_easyocr.
        None on failure.
    """
    if model is None:
        model = MIMO_MODEL
    if api_key is None:
        api_key = _get_mimo_api_key()
    # Load config
    base_url = "https://api.xiaomimimo.com/v1"
    base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
    local_cfg = os.path.join(base_dir, "mimo_config.json")
    try:
        if os.path.exists(local_cfg):
            with open(local_cfg) as f:
                cfg = json.load(f)
            if model == MIMO_MODEL:
                model = cfg.get("model", model)
            base_url = cfg.get("base_url", base_url).rstrip("/")
    except Exception:
        pass
    if not api_key:
        print("[MiMo] ERROR: 未找到 API Key")
        return None

    # Read and resize image (MiMo has image size limits — tall screenshots fail)
    try:
        img = Image.open(image_path)
        w, h = img.size
        MAX_HEIGHT = 2400  # Resize tall screenshots to fit MiMo context window
        if h > MAX_HEIGHT:
            ratio = MAX_HEIGHT / h
            new_size = (int(w * ratio), MAX_HEIGHT)
            img = img.resize(new_size, Image.LANCZOS)
            print(f"[MiMo] Resized {image_path}: {w}x{h} -> {new_size[0]}x{new_size[1]}")
        # Save to buffer
        import io
        buf = io.BytesIO()
        ext = os.path.splitext(image_path)[1].lower()
        fmt = "JPEG" if ext in (".jpg", ".jpeg") else "PNG" if ext == ".png" else "JPEG"
        save_ext = ".jpg" if fmt == "JPEG" else ".png"
        img.save(buf, format=fmt, quality=85)
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception as e:
        print(f"[MiMo] ERROR reading image: {e}")
        return None

    mime_map = {".jpg": "jpeg", ".jpeg": "jpeg", ".png": "png", ".bmp": "bmp", ".webp": "webp"}
    mime = mime_map.get(save_ext, "jpeg")

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key, base_url=base_url)
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/{mime};base64,{b64}"}},
                {"type": "text", "text": MIMO_EXTRACTION_PROMPT}
            ]}],
            max_completion_tokens=16384,
            temperature=0.1,
        )
        content = resp.choices[0].message.content
        record = _parse_mimo_response(content)
        return record

    except Exception as e:
        print(f"[MiMo] API error: {e}")
        return None


def _recover_truncated_json(text):
    """Try to salvage a truncated JSON object by closing missing braces/quotes."""
    # Count open vs close braces
    open_braces = text.count('{') - text.count('}')
    # Remove trailing incomplete key-value pair
    # Find last complete comma/brace
    text = text.rstrip()
    # Try removing last incomplete line (usually a key or value cut off mid-string)
    last_quote = text.rfind('"')
    if last_quote > 0:
        # Check if we have an odd number of quotes after last newline
        after_last_nl = text[text.rfind('\n'):]
        if after_last_nl.count('"') % 2 == 1:
            # Incomplete string — chop it
            text = text[:last_quote]
    # Strip trailing comma
    text = text.rstrip().rstrip(',').rstrip()
    # Close JSON
    text += '}' * open_braces
    try:
        return json.loads(text)
    except:
        return None


def _parse_keyvalue_lines(text):
    """Last-resort: parse 'key': 'value' or 'key': value lines from text."""
    result = {}
    import re as _re
    # Match "key": "value" or "key": value patterns
    for m in _re.finditer(r'"([^"]+)"\s*:\s*"([^"]*)"|"([^"]+)"\s*:\s*([^,\n}]+)', text):
        if m.group(1):
            result[m.group(1)] = m.group(2)
        elif m.group(3):
            val = m.group(4).strip().rstrip(',')
            result[m.group(3)] = val
    return result if result else None


def _parse_mimo_response(text):
    """Parse MiMo response text into a field:value dict."""
    # Try to extract JSON block
    # MiMo sometimes wraps JSON in markdown fences
    text = text.strip()
    if text.startswith("```"):
        lines = text.split("\n")
        text = "\n".join(lines[1:])
        if text.endswith("```"):
            text = text[:-3].strip()

    try:
        raw = json.loads(text)
    except json.JSONDecodeError:
        # Try to find JSON in the text
        import re as _re
        m = _re.search(r'\{[\s\S]*\}', text)
        if m:
            try:
                raw = json.loads(m.group())
            except json.JSONDecodeError:
                # Try truncated JSON recovery
                truncated = m.group()
                raw = _recover_truncated_json(truncated)
                if raw is None:
                    print(f"[MiMo] Failed to parse JSON from: {text[:200]}...")
                    return None
        else:
            # Try line-by-line key:value parsing as last resort
            raw = _parse_keyvalue_lines(text)
            if raw is None:
                print(f"[MiMo] No JSON found in response: {text[:200]}...")
                return None

    # Normalize keys and values
    record = {}
    for k, v in raw.items():
        if v is None:
            v = ""
        v = str(v).strip()
        # Clean noisy key names (MiMo sometimes adds units or notes in parens)
        # "井台高度(m)" -> "井台高度", "调查人(测量人)" -> "测量人"
        clean_k = k
        # Strip trailing parenthesized annotations (m), (米), (十进制度), (测量人), etc.
        clean_k = re.sub(r'\([^)]*\)$', '', clean_k).strip()
        # Also handle Chinese parentheses （）
        clean_k = re.sub(r'（[^）]*）$', '', clean_k).strip()
        # Unify MiMo-specific key names to FIELDS-compatible names
        key_map = {
            "调查单位": "统测实施单位",
            "调查单位(统测实施单位)": "统测实施单位",
            "调查人": "测量人",
            "调查人(测量人)": "测量人",
            "经度_dms": "经度",
            "纬度_dms": "纬度",
            "经度": "经度",
            "纬度": "纬度",
            "含水介质": "含水介质",
            "埋藏深度": "埋藏深度",
            "承压性": "承压性",
            "统测期次类型": "统测期次类型",
            "统测区统测期次类型": "统测期次类型",
            "井点类型": "井点类型",
            "地下水资源区名称": "地下水资源区名称",
            # Fix common OCR key errors
            "调查日期": "调查日期",
            "统测日期": "调查日期",
            "地理位置": "地理位置",
            "天气": "天气",
            "井台高度": "井台高度",
            "井深": "井深",
            "测点高程": "测点高程",
            "地面高程": "地面高程",
            "地下水位埋深": "地下水位埋深",
            "测点距水面距离": "测点距水面距离",
            "测点距地面高度": "测点距地面高度",
            "水位标高": "水位标高",
            "路线编号": "路线编号",
            "统一编号": "统一编号",
            "野外编号": "野外编号",
            "测量人": "测量人",
            "记录人": "记录人",
            "审核人": "审核人",
            "统测实施单位": "统测实施单位",
            "统测期": "统测期",
            "所属统测区类型": "所属统测区类型",
        }
        mapped = key_map.get(clean_k, clean_k)
        # If the value looks like a DMS coordinate, convert to decimal
        if mapped in ("经度", "纬度") and ("°" in v or "'" in v):
            # Keep original but also ensure we have a clean decimal
            pass
        record[mapped] = v

    # Ensure 野外编号 is present
    if "野外编号" not in record:
        for k in ("统一编号",):
            if k in record and record[k]:
                record["野外编号"] = record[k]
                break

    return record


# ── EasyOCR functions (unchanged) ───────────────────────────────────

def fuzzy_match_label(text):
    text = text.replace(" ", "").replace("*", "").strip()
    if not text: return None
    for key in sorted(LABEL_MAP, key=len, reverse=True):
        if key in text or (len(key) >= 4 and key[:6] in text) or (len(text) >= 4 and text[:6] in key):
            return LABEL_MAP[key]
    return None
def fix_ocr_number(val):
    val = val.strip()
    # Clean trailing noise like ~, -, etc from IDs
    val = re.sub(r'[~`\u2018\u2019]+$', '', val)
    # Remove trailing dash from truncated IDs
    val = re.sub(r'-$', '', val)
    # Fix common OCR prefix confusion: 1X- -> LX-
    val = val.replace('1X-', 'LX-').replace('1x-', 'LX-')
    return val


def fix_date(val):
    """Clean date strings with attached time/garbage."""
    m = re.search(r'(\d{4}-\d{2}-\d{2})', val)
    if m:
        return m.group(1)
    return val

def fix_decimal(val, label):
    """Fix OCR decimal point loss using field-specific value ranges.

    Strategy: if a numeric value exceeds the plausible max for its field,
    try dividing by 10 (most common: 6.8 -> 68) then by 100.
    Only restore when the result falls within the expected range.
    """
    val = val.strip()
    try:
        num = float(val)
    except ValueError:
        return val

    # (min, max) plausible ranges for each field (unit: meters)
    field_ranges = {
        "井台高度":          (0.01, 5.0),
        "测点距水面距离":    (0.01, 50.0),
        "测点距地面高度":    (0.01, 5.0),
        "地下水位埋深":      (0.1, 300.0),
    }
    if label not in field_ranges:
        return val

    lo, hi = field_ranges[label]
    # Already in range - keep as is
    if lo <= num <= hi:
        return val

    # Try /10 (OCR lost one decimal: 6.8 -> 68)
    if lo <= num / 10 <= hi:
        return str(num / 10)

    # Try /100 (OCR lost two decimals: 0.68 -> 68)
    if lo <= num / 100 <= hi:
        return str(num / 100)

    # Nothing worked - return original and let human review
    return val

def fix_text_value(val, label):
    val = val.strip()
    if label == "所属统测区类型" and "般" in val:
        return "一般统测区"
    if label == "统测期":
        if any(w in val for w in ["低水位", "平水位", "水平水位", "枯水"]):
            return "地下水平水位期"
    if label == "统测期次类型" and "期统测" in val:
        return "一期统测"
    if label == "承压性":
        if "承压" in val: return "承压水"
        if "潜水" in val: return "潜水"
    val = re.sub(r'[;::,,。]$', '', val)
    return val

# ── Split label handling ────────────────────────────────────────────
# Labels that OCR commonly splits across lines (header + value + suffix)
def _try_split_label(blocks, start_i):
    """Match split label: header line + value line + suffix line."""
    if start_i + 2 >= len(blocks):
        return None

    y1, x1, t1, c1 = blocks[start_i]
    t1_clean = t1.replace(" ", "").replace("*", "")

    # Must start with "地下水类型"
    if "地下水类型" not in t1_clean:
        return None

    # The pattern variants:
    # A: header(i), header_cont(i+1), value(i+2), suffix(i+3)  [chunked OCR]
    # B: header(i), value(i+1), suffix(i+2)                     [original]
    # C: header(i), suffix(i+1), value(i+2)

    # Try all suffix positions (1..3)
    for sfx_offset in [1, 2, 3]:
        if start_i + sfx_offset >= len(blocks):
            continue
        _, _, sfx_text, _ = blocks[start_i + sfx_offset]
        sfx_clean = sfx_text.replace(" ", "").replace("*", "").strip()

        # Match suffix to subtype
        if "水介质" in sfx_clean:
            sub_label = "含水介质"
        elif "水层深度" in sfx_clean:
            sub_label = "埋藏深度"
        elif "压性" in sfx_clean:
            sub_label = "承压性"
        else:
            continue

        # Value is between header and suffix
        # Skip any header continuation lines like "(按含", "(按承"
        for vi in range(start_i + 1, start_i + sfx_offset):
            _, _, v_text, _ = blocks[vi]
            v_clean = v_text.strip()
            # Skip header continuations
            if v_clean in ["(按含", "(按承", "(按含水", "按含", "按承"]:
                continue
            if v_clean.startswith("(") and len(v_clean) <= 6:
                continue
            if not v_clean:
                continue
            if any(k in v_clean for k in ["地下水", "类型"]):
                continue
            # This looks like a value
            val = fix_ocr_number(v_clean)
            val = fix_text_value(val, sub_label)
            if val and len(val) > 1:
                return (sub_label, val, start_i + sfx_offset + 1)

    return None


def pair_labels_values(text_blocks, img_height):
    blocks = []
    seen = set()
    for bbox, text, conf in text_blocks:
        top_y = min(p[1] for p in bbox)
        left_x = min(p[0] for p in bbox)
        # Deduplicate: same text at similar Y
        key = (text.strip(), top_y // 10)  # 10px tolerance
        if key in seen:
            continue
        seen.add(key)
        blocks.append((top_y, left_x, text, conf))
    blocks.sort()

    pairs = {}
    i = 0
    while i < len(blocks):
        y1, x1, text, conf = blocks[i]
        text_clean = text.replace(" ", "").replace("*", "")

        # ── Handle SPLIT labels (e.g. "地下水类型 (按含" + "水介质)") ──
        # These have the value sandwiched between the two label parts
        if not any(key in text_clean for key in LABEL_MAP):
            split_result = _try_split_label(blocks, i)
            if split_result:
                pairs[split_result[0]] = split_result[1]
                i = split_result[2]  # skip to after the split group
                continue

        # Check for "label: value" on same line
        label = None
        for key in sorted(LABEL_MAP, key=len, reverse=True):
            if key in text_clean:
                label = LABEL_MAP[key]
                # Extract value after the label
                idx = text_clean.find(key)
                after = text_clean[idx + len(key):].lstrip("::))((mM")
                if after and (after[0].isdigit() or after[0] in '一二三四五六七八九十'):
                    val = fix_ocr_number(after)
                    val = fix_decimal(val, label)
                    val = fix_text_value(val, label)
                    if label == "调查日期":
                        val = fix_date(val)
                    pairs[label] = val
                break

        # Check PREVIOUS line as value FIRST (value above label, e.g. 地理位置)
        value_above_labels = {"地理位置", "统测实施单位", "项目名称"}
        if label and i > 0 and label in value_above_labels:
            y0, x0, prev_text, prev_conf = blocks[i - 1]
            prev_label = fuzzy_match_label(prev_text.replace(" ", "").replace("*", ""))
            if not prev_label and (y1 - y0) < 80 and len(prev_text.strip()) > 1:
                # For 地理位置, the value must contain Chinese characters
                if label == "地理位置" and not re.search(r'[\u4e00-\u9fff]', prev_text):
                    pass  # Skip - not a location name
                else:
                    val = fix_ocr_number(prev_text)
                    # Also append next line if it's a distance/number continuation
                    if i + 1 < len(blocks):
                        _, _, nt, _ = blocks[i + 1]
                        if re.match(r'^\d+', nt.strip()) and not fuzzy_match_label(nt):
                            val = val + nt.strip()
                    val = fix_text_value(val, label)
                    # Fix OCR O/0 confusion in the final value
                    val = fix_ocr_number(val)
                    pairs[label] = val
                    i += 1
                    continue

        # Check next line as value
        if label and i + 1 < len(blocks):
            y2, x2, next_text, next_conf = blocks[i + 1]
            next_label = fuzzy_match_label(next_text.replace(" ", "").replace("*", ""))
            if not next_label and (y2 - y1) < 80:
                val = fix_ocr_number(next_text)
                val = fix_decimal(val, label)
                val = fix_text_value(val, label)
                if label == "调查日期":
                    val = fix_date(val)
                pairs[label] = val
                i += 2
                continue

        # Check PREVIOUS line as value (value above label, e.g. 地理位置)
        if label and i > 0 and label not in pairs:
            y0, x0, prev_text, prev_conf = blocks[i - 1]
            prev_label = fuzzy_match_label(prev_text.replace(" ", "").replace("*", ""))
            if not prev_label and (y1 - y0) < 80 and len(prev_text.strip()) > 1:
                # For 地理位置, the value must contain Chinese characters
                if label == "地理位置" and not re.search(r'[\u4e00-\u9fff]', prev_text):
                    pass  # Skip - not a location name
                else:
                    val = fix_ocr_number(prev_text)
                    # For 地理位置, also append next line if it's a distance
                    if label == "地理位置" and i + 1 < len(blocks):
                        _, _, nt, _ = blocks[i + 1]
                        if re.match(r'^\d+', nt.strip()):
                            val = val + nt.strip()
                    val = fix_text_value(val, label)
                    if label == "调查日期":
                        val = fix_date(val)
                    pairs[label] = val

        # Try same-Y different-X (inline value)
        if label:
            for j in range(i + 1, min(i + 4, len(blocks))):
                yj, xj, jt, jc = blocks[j]
                if abs(yj - y1) < 30 and xj > x1 + 50:
                    if not fuzzy_match_label(jt):
                        val = fix_ocr_number(jt)
                        val = fix_decimal(val, label)
                        val = fix_text_value(val, label)
                        if label == "调查日期":
                            val = fix_date(val)
                        if val and val.strip() and not all(c.isalpha() for c in val.strip().replace(" ", "")):
                            pairs[label] = val
                        break

        i += 1

    return pairs

def dms_to_decimal_str(dms):
    """Convert various DMS formats to decimal string."""
    dms = dms.strip()
    dms = dms.replace(chr(0x2033), "").replace('"', "").replace(chr(0x00B0), " ")
    dms = dms.replace("''", "").replace('\u201d', "").replace('\u2033', "")
    dms = dms.replace("'", " ")  # Minute marker → space separator

    # Format 1: DDD MM SS.SS (e.g. "111 49 27.38")
    m = re.match(r"(\d{2,3})\s+(\d{1,2})\s+(\d{1,2}(?:\.\d+)?)", dms)
    if m:
        d, mi, s = float(m.group(1)), float(m.group(2)), float(m.group(3))
        if mi < 60 and s < 60:
            return "{:.6f}".format(d + mi / 60 + s / 3600)

    # Format 2: DDDMM SS.SS (e.g. "11149 27.38")
    m = re.match(r"(\d{3})(\d{2})\s+(\d{1,2}(?:\.\d+)?)", dms)
    if m:
        d, mi, s = float(m.group(1)), float(m.group(2)), float(m.group(3))
        if mi < 60 and s < 60:
            return "{:.6f}".format(d + mi / 60 + s / 3600)

    # Format 3: DDMM SS.SS (e.g. "3828 30.64")
    m = re.match(r"(\d{2})(\d{2})\s+(\d{1,2}(?:\.\d+)?)", dms)
    if m:
        d, mi, s = float(m.group(1)), float(m.group(2)), float(m.group(3))
        if mi < 60 and s < 60:
            return "{:.6f}".format(d + mi / 60 + s / 3600)

    # Format 4: DD MM SS.SS (e.g. "38 4 52.8")
    m = re.match(r"(\d{1,2})\s+(\d{1,2})\s+(\d{1,2}(?:\.\d+)?)", dms)
    if m:
        d, mi, s = float(m.group(1)), float(m.group(2)), float(m.group(3))
        if mi < 60 and s < 60:
            return "{:.6f}".format(d + mi / 60 + s / 3600)

    return dms

def _dedup_chunks(results):
    """Remove duplicate detections from overlapping image chunks.

    Two results are considered duplicates when they share the same text
    and their vertical center positions are within 200 px (less than the
    300 px chunk overlap).  The result with higher confidence is kept.
    """
    if not results:
        return results
    deduped = []
    for bbox, text, conf in results:
        y_center = (min(p[1] for p in bbox) + max(p[1] for p in bbox)) / 2
        merged = False
        for j, (db, dt, dc) in enumerate(deduped):
            dy_center = (min(p[1] for p in db) + max(p[1] for p in db)) / 2
            if dt == text and abs(y_center - dy_center) < 200:
                if conf > dc:
                    deduped[j] = (bbox, text, conf)
                merged = True
                break
        if not merged:
            deduped.append((bbox, text, conf))
    return deduped

def _rapidocr_to_easyocr(raw_result):
    """Convert RapidOCR output to EasyOCR-compatible format."""
    if not raw_result:
        return []
    converted = []
    for item in raw_result:
        bbox = [[float(p[0]), float(p[1])] for p in item[0]]
        text = str(item[1])
        conf = float(item[2])
        converted.append((bbox, text, conf))
    return converted


def extract_from_image_easyocr(image_path):
    engine = get_reader()
    img = Image.open(image_path)
    img_array = np.array(img)
    h, w = img_array.shape[:2]

    # Long images: process in chunks to avoid resizing loss
    if h > 3000:
        all_results = []
        chunk_h = 2000
        overlap = 300
        for y in range(0, h, chunk_h - overlap):
            y_end = min(y + chunk_h, h)
            chunk = img_array[y:y_end, :, :]
            raw, _ = engine(chunk)
            chunk_results = _rapidocr_to_easyocr(raw)
            # Adjust Y coordinates back to full image space
            for bbox, text, conf in chunk_results:
                adjusted_bbox = [[p[0], p[1] + y] for p in bbox]
                all_results.append((adjusted_bbox, text, conf))
        # Deduplicate overlapping chunk results
        results = _dedup_chunks(all_results)
    else:
        raw, _ = engine(img_array)
        results = _rapidocr_to_easyocr(raw)

    if not results:
        return None

    pairs = pair_labels_values(results, h)

    record = {}
    for label, val in pairs.items():
        if label in ("经度_dms", "纬度_dms"):
            base = label.replace("_dms", "")
            record[base] = dms_to_decimal_str(val)
        elif label == "经度_decimal":
            record["经度"] = val
        elif label == "纬度_decimal":
            record["纬度"] = val
        else:
            record[label] = val

    # Fallback: extract field ID from title bar
    if "野外编号" not in record:
        for _, text, _ in results:
            m = re.search(r'(TYXB\d+|TSK\d+)', text.replace('O', '0').replace('o', '0'))
            if m:
                record["野外编号"] = m.group(1)
                break

    if "井点类型" not in record:
        for _, text, _ in results:
            m = re.search(r'(机民井|民井|监测井|大口井|钻孔)', text)
            if m:
                record["井点类型"] = m.group(1)
                break

    if "调查日期" not in record:
        for _, text, _ in results:
            m = re.search(r'(\d{4}-\d{2}-\d{2})', text)
            if m:
                record["调查日期"] = m.group(1).replace("-", "")
                break

    # ── Keyword fallback for groundwater type fields ──
    gw_media = {"孔隙水": "含水介质", "裂隙水": "含水介质", "岩溶水": "含水介质",
                 "孔隙-裂隙水": "含水介质"}
    gw_depth = {"浅层": "埋藏深度", "中层": "埋藏深度", "中深层": "埋藏深度", "深层": "埋藏深度"}
    gw_press = {"潜水": "承压性", "承压水": "承压性", "微承压水": "承压性"}

    for _, text, _ in results:
        t = text.strip()
        for kw, field in gw_media.items():
            if kw in t and (field not in record or record[field] in ('', '压性)', '水介质)', '水层深度)')):
                record[field] = kw
        for kw, field in gw_depth.items():
            if kw in t and (field not in record or record[field] in ('', '压性)', '水介质)', '水层深度)')):
                record[field] = kw
        for kw, field in gw_press.items():
            if kw in t and (field not in record or record[field] in ('', '压性)', '水介质)', '水层深度)')):
                record[field] = kw

    return record
