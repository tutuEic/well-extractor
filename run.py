"""
Complete pipeline runner: OCR extraction → Excel fill
"""
import sys, os, json, subprocess, time

BASE = r"D:\hermes-workspace\projects\well-extractor"
OS_SEP = "\\"

def step1_extract(image_dir, output_json):
    """Run OCR extraction"""
    print("=" * 50)
    print("Step 1: OCR Extraction")
    print("=" * 50)
    
    # Import and run directly
    sys.path.insert(0, os.path.join(BASE, "extractor"))
    from ocr_extract import extract_batch
    
    records = extract_batch(image_dir, output_json)
    return records


def step2_fill(json_path, template_path, output_xlsx):
    """Fill template Excel with extracted data"""
    print("\n" + "=" * 50)
    print("Step 2: Fill Excel Template")
    print("=" * 50)
    
    sys.path.insert(0, os.path.join(BASE, "extractor"))
    from excel_filler import load_records, fill_template
    
    records = load_records(json_path)
    updated, added = fill_template(records, template_path, output_xlsx)
    return updated, added


def main():
    import argparse
    p = argparse.ArgumentParser(description="地下水统测信息提取器")
    p.add_argument("--images", default=r"D:\地下水\信息截图\5.21", help="图片目录")
    p.add_argument("--template", default=r"D:\地下水\大同朔州忻州_地下水统测表_补全_最终版.xlsx", help="模板Excel")
    p.add_argument("--output", default=os.path.join(BASE, "output", "filled.xlsx"), help="输出Excel")
    p.add_argument("--json", default=os.path.join(BASE, "output", "extracted.json"), help="中间JSON")
    args = p.parse_args()
    
    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    
    t0 = time.time()
    records = step1_extract(args.images, args.json)
    updated, added = step2_fill(args.json, args.template, args.output)
    
    print(f"\n{'='*50}")
    print(f"Done in {time.time()-t0:.0f}s")
    print(f"Records: {len(records)}, Updated: {updated}, Added: {added}")
    print(f"Output: {args.output}")


if __name__ == "__main__":
    main()
