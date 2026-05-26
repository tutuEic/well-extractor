"""
Generate Word docs from template + Excel data.
For each row 806-836 in E2E_v2_FULL.xlsx, fill template and save.
"""
import os, shutil
import openpyxl
from docx import Document
from docx.shared import Pt

TEMPLATE = r'D:\地下水\模版.docx'
EXCEL = r'D:\地下水\E2E_v2_FULL.xlsx'
OUT_DIR = r'D:\地下水\word'

os.makedirs(OUT_DIR, exist_ok=True)

# Read Excel
wb = openpyxl.load_workbook(EXCEL, data_only=False)
ws = wb['Sheet1']

# Column mapping from FIELDS (name -> Excel column letter)
col_names = {}  # col_letter -> value from row 1
for col in range(1, ws.max_column + 1):
    h = ws.cell(row=1, column=col).value
    if h:
        col_names[h] = col

def get_val(row, field):
    """Get value from Excel row by field name."""
    if field not in col_names:
        return ''
    val = ws.cell(row=row, column=col_names[field]).value
    if val is None:
        return ''
    # If it's a formula string like "=L825-V825", return as is
    s = str(val).strip()
    return s

# Process rows 806-836
for row in range(806, 837):
    fid = get_val(row, '野外编号')
    if not fid:
        continue
    
    doc = Document(TEMPLATE)
    table = doc.tables[0]
    
    # ── Fill the table ──────────────────────────────────────────
    # Helper: write to cell [row_idx, col_idx]
    def w(r, c, val):
        if val:
            cell = table.cell(r, c)
            # Clear existing paragraphs
            for p in cell.paragraphs:
                p.clear()
            # Add our text
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.name = '仿宋'
            run.font.size = Pt(10)
    
    # Row 0: 统一编号 | 野外编号
    w(0, 1, get_val(row, '统一编号'))
    w(0, 4, fid)
    w(0, 5, fid)
    
    # Row 1: 路线编号 | 调查日期
    w(1, 1, get_val(row, '路线编号'))
    w(1, 2, get_val(row, '路线编号'))
    dt = get_val(row, '调查日期')
    w(1, 4, dt)
    w(1, 5, dt)
    
    # Row 2: 地理位置 | 天气
    w(2, 1, get_val(row, '地理位置'))
    w(2, 2, get_val(row, '地理位置'))
    weather = get_val(row, '天气')
    w(2, 4, weather)
    w(2, 5, weather)
    
    # Row 3: 经度 | 纬度
    w(3, 1, get_val(row, '经度'))
    w(3, 2, get_val(row, '经度'))
    w(3, 4, get_val(row, '纬度'))
    w(3, 5, get_val(row, '纬度'))
    
    # Row 4: 地面高程 | 地面高程获取方法
    w(4, 1, get_val(row, '地面高程'))
    w(4, 2, get_val(row, '地面高程'))
    method = get_val(row, '地面高程获取方法')
    w(4, 4, method)
    w(4, 5, method)
    
    # Row 5: 测点高程 | 测点高程获取方法
    elev = get_val(row, '测点高程')
    w(5, 1, elev)
    w(5, 2, elev)
    method2 = get_val(row, '测点高程获取方法')
    w(5, 4, method2)
    w(5, 5, method2)
    
    # Row 6: 井台高度 | 井深
    w(6, 1, get_val(row, '井台高度'))
    w(6, 2, get_val(row, '井台高度'))
    w(6, 4, get_val(row, '井深'))
    w(6, 5, get_val(row, '井深'))
    
    # Row 7: 地下水资源区名称 | 所属统测区类型
    w(7, 1, get_val(row, '地下水资源区名称'))
    w(7, 2, get_val(row, '地下水资源区名称'))
    zone = get_val(row, '所属统测区类型')
    w(7, 4, zone)
    w(7, 5, zone)
    
    # Row 8: 统测期 | 统测期次类型
    period = get_val(row, '统测期')
    w(8, 1, period)
    w(8, 2, period)
    sub = get_val(row, '统测期次类型')
    w(8, 4, sub)
    w(8, 5, sub)
    
    # Row 9: 含水介质
    media = get_val(row, '含水介质')
    w(9, 1, media)
    w(9, 2, media)
    w(9, 3, media)
    w(9, 4, media)
    w(9, 5, media)
    
    # Row 10: 埋藏深度 | 承压性
    depth = get_val(row, '埋藏深度')
    w(10, 1, depth)
    w(10, 2, depth)
    press = get_val(row, '承压性')
    w(10, 4, press)
    w(10, 5, press)
    
    # Row 11: 测点距地面高度 | 测点距水面距离
    w(11, 1, get_val(row, '测点距地面高度'))
    w(11, 2, get_val(row, '测点距地面高度'))
    w(11, 4, get_val(row, '测点距水面距离'))
    w(11, 5, get_val(row, '测点距水面距离'))
    
    # Row 12: 地下水位埋深 | 水位标高
    w(12, 1, get_val(row, '地下水位埋深'))
    w(12, 2, get_val(row, '地下水位埋深'))
    w(12, 4, get_val(row, '水位标高'))
    w(12, 5, get_val(row, '水位标高'))
    
    # Row 15: 照片编号 | 视频编号
    photo_id = fid + '文件夹'
    w(15, 1, photo_id)
    w(15, 2, photo_id)
    w(15, 4, get_val(row, '视频编号'))
    w(15, 5, get_val(row, '视频编号'))
    
    # Row 16: 备注
    note = get_val(row, '备注')
    w(16, 1, note)
    w(16, 2, note)
    w(16, 3, note)
    
    # Row 17: 项目名称
    proj = '山西省地下水水位统测'
    w(17, 1, proj)
    w(17, 2, proj)
    w(17, 4, proj)
    
    # Row 18: 统测实施单位
    unit = get_val(row, '统测实施单位')
    w(18, 1, unit)
    w(18, 2, unit)
    w(18, 4, unit)
    
    # Row 19: 测量人 | 记录人 | 审核人
    w(19, 1, get_val(row, '测量人'))
    w(19, 3, get_val(row, '记录人'))
    w(19, 5, get_val(row, '审核人'))
    
    # Save
    out_path = os.path.join(OUT_DIR, f'{fid}.docx')
    doc.save(out_path)
    print(f'Saved: {fid}.docx')

wb.close()
print(f'\nDone! {row - 805} files in {OUT_DIR}')
