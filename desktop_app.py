"""
地下水统测信息提取器 — WPF 风格桌面应用
支持：图片 OCR 提取 + Excel 文件读取 + 表格编辑 + 导出
"""
import sys, os, re, json, shutil, tempfile
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QToolBar, QStatusBar, QMenuBar, QMenu, QSplitter,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QTreeWidget, QTreeWidgetItem,
    QPushButton, QLabel, QFileDialog, QMessageBox,
    QProgressBar, QTabWidget, QGroupBox, QStyle,
    QStyleFactory, QAbstractItemView, QSizePolicy,
)
from PySide6.QtCore import Qt, QThread, Signal, QSize, QTimer
from PySide6.QtGui import QAction, QIcon, QFont, QColor, QPalette

import pytesseract
from PIL import Image
import openpyxl

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = r"D:\地下水\大同朔州忻州_地下水统测表_补全_最终版.xlsx"
COORD_PATH = r"D:\地下水\坐标.xlsx"

# ===== Coordinate Index =====
_coord_index = None

def load_coord_index():
    """Build index from 坐标.xlsx: Col1=编号, Col3=X, Col4=Y, Col5=高程, Col6=经度, Col7=纬度"""
    global _coord_index
    if _coord_index is not None:
        return _coord_index
    
    _coord_index = {}
    if not os.path.exists(COORD_PATH):
        print(f"坐标文件不存在: {COORD_PATH}")
        return _coord_index
    
    wb = openpyxl.load_workbook(COORD_PATH, data_only=True)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Headers: Col1=编号, Col2=?, Col3=X, Col4=Y, Col5=高程, Col6=经度, Col7=纬度
        for row in range(2, ws.max_row + 1):
            fid = ws.cell(row=row, column=1).value  # 野外编号
            if not fid: continue
            fid = str(fid).strip()
            
            def _clean(v):
                if v is None: return None
                if isinstance(v, str): return v.strip().replace('\t', '').replace('\n', '')
                return v
            _coord_index[fid] = {
                "X": _clean(ws.cell(row=row, column=3).value),
                "Y": _clean(ws.cell(row=row, column=4).value),
                "地面高程": _clean(ws.cell(row=row, column=5).value),
                "经度": _clean(ws.cell(row=row, column=6).value),
                "纬度": _clean(ws.cell(row=row, column=7).value),
            }
    wb.close()
    print(f"坐标索引: {len(_coord_index)} 条记录")
    return _coord_index


def merge_coordinates(record):
    """坐标按顺序匹配：第1条坐标→第1条TYXB记录"""
    # Already handled in batch mode below
    return record


def batch_merge_coordinates(records):
    """
    Merge coordinates from 坐标.xlsx into records.
    Coordinates are auto-loaded — match by index position.
    """
    coords = load_coord_index()
    if not coords:
        return records
    
    # Get coordinate records in order
    coord_list = list(coords.values())
    
    # Find records that have valid TYXB/TSK IDs (from screenshots)
    valid_records = [r for r in records if r.get("野外编号") and re.search(r'(TYXB\d+|TSK\d+)', str(r.get("野外编号","")))]
    
    # Match by position
    for i, r in enumerate(valid_records):
        if i < len(coord_list):
            c = coord_list[i]
            for k, v in c.items():
                if v is not None and v != '':
                    r[k] = str(v).replace('\t', '') if isinstance(v, str) else v
            r["_has_coords"] = True
    return records

# ===== Dark Theme Stylesheet =====
DARK_STYLE = """
QMainWindow { background: #1e1e2e; }
QMenuBar { background: #181825; color: #cdd6f4; border-bottom: 1px solid #313244; padding: 2px; }
QMenuBar::item:selected { background: #45475a; }
QMenu { background: #1e1e2e; color: #cdd6f4; border: 1px solid #313244; }
QMenu::item:selected { background: #45475a; }
QToolBar { background: #181825; border-bottom: 1px solid #313244; spacing: 6px; padding: 4px; }
QToolBar QToolButton { color: #cdd6f4; padding: 4px 10px; border-radius: 4px; }
QToolBar QToolButton:hover { background: #45475a; }
QStatusBar { background: #181825; color: #a6adc8; border-top: 1px solid #313244; }
QSplitter::handle { background: #313244; width: 2px; }
QTreeWidget { background: #181825; color: #cdd6f4; border: 1px solid #313244; border-radius: 6px; }
QTreeWidget::item { padding: 4px; }
QTreeWidget::item:selected { background: #45475a; }
QTableWidget { background: #1e1e2e; color: #cdd6f4; gridline-color: #313244; border: 1px solid #313244; border-radius: 6px; }
QTableWidget::item { padding: 4px; }
QTableWidget::item:selected { background: #45475a; }
QHeaderView::section { background: #181825; color: #cdd6f4; border: none; border-bottom: 2px solid #313244; padding: 6px; font-weight: bold; }
QPushButton { background: #45475a; color: #cdd6f4; border: none; padding: 6px 14px; border-radius: 4px; }
QPushButton:hover { background: #585b70; }
QPushButton:pressed { background: #313244; }
QPushButton#btnExtract { background: #89b4fa; color: #1e1e2e; font-weight: bold; }
QPushButton#btnExtract:hover { background: #b4d0fb; }
QPushButton#btnExport { background: #a6e3a1; color: #1e1e2e; font-weight: bold; }
QPushButton#btnExport:hover { background: #c3f0c0; }
QProgressBar { background: #313244; border: none; border-radius: 3px; height: 6px; text-align: center; }
QProgressBar::chunk { background: #89b4fa; border-radius: 3px; }
QGroupBox { color: #cdd6f4; border: 1px solid #313244; border-radius: 6px; margin-top: 8px; padding-top: 16px; }
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
QTabWidget::pane { border: 1px solid #313244; background: #1e1e2e; }
QTabBar::tab { background: #181825; color: #a6adc8; padding: 6px 16px; border: none; }
QTabBar::tab:selected { color: #cdd6f4; border-bottom: 2px solid #89b4fa; }
"""

FIELDS = [
    ("野外编号", "C"), ("统一编号", "B"), ("路线编号", "D"),
    ("调查日期", "E"), ("地理位置", "F"), ("天气", "G"),
    ("经度", "H"), ("纬度", "I"), ("X", "J"), ("Y", "K"),
    ("地面高程", "L"), ("地面高程获取方法", "M"),
    ("测点高程", "N"), ("测点高程获取方法", "O"),
    ("井点类型", "T"), ("所属统测区类型", "S"),
    ("井台高度", "P"), ("井深", "Q"),
    ("地下水资源区名称", "R"),
    ("测点距地面高度", "U"), ("地下水位埋深", "V"), ("测点距水面距离", "W"),
    ("水位标高", "X"),
    ("含水介质", "Y"), ("埋藏深度", "Z"), ("承压性", "AA"),
    ("统测期", "AB"), ("统测期次类型", "AC"),
    ("照片编号", "AD"), ("视频编号", "AE"), ("备注", "AF"),
    ("测量人", "AH"), ("记录人", "AI"), ("审核人", "AJ"),
    ("统测实施单位", "AG"),
]

# Fixed values for every record
FIXED_VALUES = {
    "路线编号": "LX-3",
    "地面高程获取方法": "RTK",
    "测点高程获取方法": "RTK",
    "统测实施单位": "山西省地质调查院有限公司",
    "测量人": "杨光",
    "记录人": "杨升伟",
    "审核人": "杨升伟",
    "照片编号": "/",
    "视频编号": "/",
    "备注": "/",
    "所属统测区类型": "一般统测区",
    "统测期": "地下水平水位期",
    "统测期次类型": "一期统测",
    "地下水资源区名称": "天桥岩溶",
}

# ===== OCR Extraction (same V4 logic) =====
GW_介质 = ["孔隙水", "岩溶水", "裂隙水", "孔隙-裂隙水"]
GW_深度 = ["浅层", "中层", "中深层", "深层"]
GW_承压 = ["潜水", "承压水", "微承压水"]

def fuzzy_gw(text, candidates, min_ratio=0.4):
    text = text.replace(" ", "").replace("\n", "")
    best, best_score = None, 0
    for c in candidates:
        score = sum(1 for ch in c if ch in text)
        if score > best_score and score >= len(c) * min_ratio:
            best_score, best = score, c
    return best

def extract_from_image(img_path):
    img = Image.open(img_path)
    full = pytesseract.image_to_string(img, lang="chi_sim+eng", config="--psm 3")
    clean = re.sub(r'(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])', '', full)
    lines = clean.split("\n")
    results = {}
    
    def try_get(patterns, key, post=None, search_text=None):
        text = search_text or clean
        for p in patterns:
            m = re.search(p, text)
            if m:
                v = m.group(1).strip()
                if v in ('/', '无', '请选择', '请输入', '请选择审核人'): continue
                if post: v = post(v)
                if v: results[key] = v
                return
    
    # === ID extraction: title bar is most reliable ===
    # Pattern: "TYXB06 统测(机民井)" — ID in large font at top
    for line in lines[:5]:
        m = re.search(r'(TYXB\d+|TSK\d+)', line)
        if m:
            results["野外编号"] = m.group(1)
            break
    
    # Well type from same title bar
    for line in lines[:5]:
        m = re.search(r'统\s*[测调].*[(（]([^)）]+)[)）]', line)
        if m:
            wt = re.sub(r'[国日\"\'\\]+', '', m.group(1).strip())
            if wt and len(wt) <= 10: results["井点类型"] = wt; break
        # Also try simpler pattern
        m2 = re.search(r'[(（]([机民钻探观测]+井)[)）]', line)
        if m2: results["井点类型"] = m2.group(1); break
    
    # Fallback: try "野外编号: TYXB##" pattern
    if "野外编号" not in results:
        try_get([r'[野量]外编号[：:\s]*(\S+)', r'编号[：:\s]*([A-Z][A-Z0-9]+)'], "野外编号")
    try_get([r'统[测调][日查]期[：:\s]*(\S+)'], "调查日期", lambda v: v.replace('闻','').strip())
    try_get([r'天气[：:\s]*(\S+)'], "天气", lambda v: re.sub(r'[一园国日\"\'\\]+', '', v).strip())
    try_get([r'地理位置[：:]?\s*(.{8,80})'], "地理位置")
    
    for field, pats in [
        ("井台高度", [r'井台高度[^0-9]*(\d+\.?\d*)']),
        ("井深", [r'[些井此]深[^0-9]*(\d+\.?\d*)']),
        ("测点距地面高度", [r'距地面高度[^0-9]*(\d+\.?\d*)']),
        ("地下水位埋深", [r'[地下]*[水位]*埋深[^0-9]*(\d+\.?\d*)']),
        ("测点距水面距离", [r'距水面距离[^0-9]*(\d+\.?\d*)']),
    ]:
        try_get(pats, field, lambda v: str(float(v)) if v and float(v) < 10000 else None)
    
    try_get([r'所属统测区类型[：:\s]*(\S+)', r'统测区类型[：:\s]*(\S+)'], "所属统测区类型",
            lambda v: re.sub(r'[航国\"\'\\]+', '', v).strip())
    try_get([r'统测期[次]?[：:\s]*(\S+)'], "统测期", lambda v: re.sub(r'[eEoO0\"\'\\一©国]+', '', v).strip())
    try_get([r'统测[区期]次类型[：:\s]*(\S+)', r'期次类型[：:\s]*(\S+)'], "统测期次类型",
            lambda v: re.sub(r'[国团\"\'\\一&]+', '', v).strip())
    
    for i, line in enumerate(lines):
        if '含水介' in line or '含水层' in line or '承压' in line:
            context = line
            for j in range(i+1, min(i+3, len(lines))): context += " " + lines[j]
            context = context.replace(" ", "")
            if '含水介' in context:
                gw = fuzzy_gw(context, GW_介质, min_ratio=0.3)
                if gw: results["含水介质"] = gw
            if '含水层' in context or '埋藏' in context:
                gw = fuzzy_gw(context, GW_深度)
                if gw: results["埋藏深度"] = gw
            if '承压' in context:
                gw = fuzzy_gw(context, GW_承压)
                if gw: results["承压性"] = gw
    
    try_get([r'测量人[：:\s]*(\S+)'], "测量人", lambda v: re.sub(r'[国园\"\'\\~\-一¥]+', '', v).strip())
    try_get([r'记录人[：:\s]*(\S+)'], "记录人", lambda v: re.sub(r'[国园\"\'\\~¥]+', '', v).strip())
    try_get([r'审核人[：:\s]*(\S+)'], "审核人",
            lambda v: v if v not in ('请选择审核人',) and not re.search(r'[盆地|山区|平剖]', v) else None)
    try_get([r'实施单位[：:\s]*(.{4,60})', r'统测实施单位[：:\s]*(.{4,60})'], "统测实施单位",
            lambda v: '山西省地质调查院有限公司')
    
    # === Post-processing cleanup ===
    for r in [results]:
        # Apply all fixed values FIRST (always override OCR)
        for key, val in FIXED_VALUES.items():
            r[key] = val
        
        # CRITICAL: validate 野外编号 — must be TYXB## or TSK### format
        fid = str(r.get("野外编号", "")).strip()
        if not fid or not re.search(r'(TYXB\d+|TSK\d+)', fid):
            # Invalid ID — try to extract from title bar again harder
            for line in lines[:5]:
                m = re.search(r'(TYXB\d+|TSK\d+)', line)
                if m:
                    r["野外编号"] = m.group(1)
                    break
        
        # Date: always use today
        r["调查日期"] = datetime.now().strftime("%Y%m%d")
        
        # 测点高程 = 地面高程 (same value)
        if r.get("地面高程") and not r.get("测点高程"):
            r["测点高程"] = r["地面高程"]
        
        # Weather: default to 晴 if not recognized
        if not r.get("天气") or r["天气"] not in ("晴", "阴"):
            r["天气"] = "晴"
        
        # Clean personnel names
        if r.get("测量人") and "杨光" in str(r["测量人"]): r["测量人"] = "杨光"
        if r.get("记录人") and "杨升伟" in str(r["记录人"]): r["记录人"] = "杨升伟"
        
        # Normalize date
        dt = r.get("调查日期", "")
        if dt:
            dt = str(dt).replace("-", "").replace(":", "").replace(" ", "").replace("000000", "")
            if len(dt) >= 8: r["调查日期"] = dt[:8]
    
    return results


def read_xlsx_file(file_path):
    """Read xlsx and extract rows as records"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    records = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        headers = [c.value for c in ws[1]]
        if not headers: continue
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            record = {}
            for h, v in zip(headers, row):
                if h and v is not None:
                    record[str(h)] = str(v) if not isinstance(v, (int, float)) else v
            if record:
                record["_source"] = f"{os.path.basename(file_path)} [{sheet_name}]"
                records.append(record)
    wb.close()
    return records


# ===== Worker Thread =====
class ExtractWorker(QThread):
    progress = Signal(int)
    finished = Signal(list)
    error = Signal(str)
    
    def __init__(self, files):
        super().__init__()
        self.files = files
    
    def run(self):
        records = []
        total = len(self.files)
        for i, f in enumerate(self.files):
            try:
                ext = os.path.splitext(f)[1].lower()
                if ext in ('.jpg', '.jpeg', '.png', '.bmp'):
                    r = extract_from_image(f)
                elif ext in ('.xlsx', '.xls'):
                    # Skip coordinate file — it's auto-loaded separately
                    if '坐标' in os.path.basename(f):
                        continue
                    recs = read_xlsx_file(f)
                    records.extend(recs)
                    self.progress.emit(int((i+1)/total*100))
                    continue
                else:
                    continue
                r["_source"] = os.path.basename(f)
                records.append(r)
            except Exception as e:
                records.append({"_source": os.path.basename(f), "_error": str(e)})
            self.progress.emit(int((i+1)/total*100))
        
        # Batch merge coordinates by position
        batch_merge_coordinates(records)
        self.finished.emit(records)


# ===== Main Window =====
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("地下水统测信息提取器")
        self.resize(1400, 850)
        self.setMinimumSize(1000, 600)
        self.records = []
        self.files = []
        
        self._setup_ui()
        
    def _setup_ui(self):
        # Menu bar
        menubar = self.menuBar()
        file_menu = menubar.addMenu("文件(&F)")
        file_menu.addAction("打开图片...", self._open_images, "Ctrl+O")
        file_menu.addAction("打开Excel...", self._open_excel, "Ctrl+E")
        file_menu.addSeparator()
        file_menu.addAction("导出填充表...", self._export, "Ctrl+S")
        file_menu.addSeparator()
        file_menu.addAction("退出", self.close, "Alt+F4")
        
        help_menu = menubar.addMenu("帮助(&H)")
        help_menu.addAction("关于", self._about)
        
        # Toolbar
        toolbar = QToolBar("工具栏")
        toolbar.setIconSize(QSize(20, 20))
        self.addToolBar(toolbar)
        
        self.btn_images = QPushButton("📷 添加图片")
        self.btn_images.clicked.connect(self._open_images)
        toolbar.addWidget(self.btn_images)
        
        self.btn_excel = QPushButton("📊 添加Excel")
        self.btn_excel.clicked.connect(self._open_excel)
        toolbar.addWidget(self.btn_excel)
        
        toolbar.addSeparator()
        
        self.btn_extract = QPushButton("🔍 开始提取")
        self.btn_extract.setObjectName("btnExtract")
        self.btn_extract.clicked.connect(self._extract)
        toolbar.addWidget(self.btn_extract)
        
        toolbar.addSeparator()
        
        self.btn_clear = QPushButton("清空")
        self.btn_clear.clicked.connect(self._clear)
        toolbar.addWidget(self.btn_clear)
        
        # Spacer and export
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        toolbar.addWidget(spacer)
        
        self.btn_export = QPushButton("📥 导出Excel")
        self.btn_export.setObjectName("btnExport")
        self.btn_export.clicked.connect(self._export)
        self.btn_export.setVisible(False)
        toolbar.addWidget(self.btn_export)
        
        # Central widget with splitter
        splitter = QSplitter(Qt.Horizontal)
        
        # Left panel: file list
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(4, 4, 4, 4)
        
        file_group = QGroupBox("文件列表")
        file_layout = QVBoxLayout(file_group)
        self.file_tree = QTreeWidget()
        self.file_tree.setHeaderLabels(["文件"])
        self.file_tree.setRootIsDecorated(False)
        file_layout.addWidget(self.file_tree)
        
        self.lbl_count = QLabel("就绪")
        self.lbl_count.setStyleSheet("color: #a6adc8; padding: 4px;")
        file_layout.addWidget(self.lbl_count)
        left_layout.addWidget(file_group)
        
        splitter.addWidget(left)
        
        # Center: data table
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(4, 4, 4, 4)
        
        # Tabs
        self.tabs = QTabWidget()
        
        # Table tab
        self.table = QTableWidget()
        self.table.setColumnCount(len(FIELDS) + 3)
        self.table.setHorizontalHeaderLabels(["状态", "模板行", "来源"] + [f[0] for f in FIELDS])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.cellChanged.connect(self._cell_changed)
        self.tabs.addTab(self.table, "📋 数据表格")
        
        # Template preview tab
        self.template_preview = QLabel("模板：大同朔州忻州_地下水统测表_补全_最终版.xlsx\n\n上传文件后点击「开始提取」")
        self.template_preview.setAlignment(Qt.AlignCenter)
        self.template_preview.setStyleSheet("color: #a6adc8; padding: 40px;")
        self.tabs.addTab(self.template_preview, "📄 模板预览")
        
        right_layout.addWidget(self.tabs)
        
        splitter.addWidget(right)
        splitter.setSizes([250, 1150])
        
        self.setCentralWidget(splitter)
        
        # Status bar
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)
        self.progress = QProgressBar()
        self.progress.setMaximumWidth(200)
        self.progress.hide()
        self.statusbar.addPermanentWidget(self.progress)
        self.lbl_status = QLabel("就绪")
        self.statusbar.addWidget(self.lbl_status)
    
    def _open_images(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择截图", "", "Images (*.jpg *.jpeg *.png *.bmp)")
        if files:
            self.files.extend(files)
            self._refresh_file_list()
    
    def _open_excel(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", "", "Excel (*.xlsx *.xls)")
        if files:
            self.files.extend(files)
            self._refresh_file_list()
    
    def _refresh_file_list(self):
        self.file_tree.clear()
        for f in self.files:
            icon = "📷" if os.path.splitext(f)[1].lower() in ('.jpg','.jpeg','.png','.bmp') else "📊"
            QTreeWidgetItem(self.file_tree, [f"{icon} {os.path.basename(f)}"])
        self.lbl_count.setText(f"{len(self.files)} 个文件")
    
    def _extract(self):
        if not self.files:
            QMessageBox.warning(self, "提示", "请先添加文件")
            return
        
        self.btn_extract.setEnabled(False)
        self.btn_extract.setText("⏳ 提取中...")
        self.progress.show()
        self.progress.setValue(0)
        self.lbl_status.setText("正在提取...")
        
        self.worker = ExtractWorker(self.files)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(lambda e: QMessageBox.critical(self, "错误", e))
        self.worker.start()
    
    def _on_progress(self, val):
        self.progress.setValue(val)
    
    def _on_finished(self, records):
        self.records = records
        # DEBUG
        coords_count = sum(1 for r in records if r.get("_has_coords"))
        first = records[0] if records else {}
        print(f"[DEBUG] _on_finished: {len(records)} records, {coords_count} with coords")
        print(f"[DEBUG] First record keys with non-None values: {[(k,v) for k,v in first.items() if v is not None and not k.startswith('_')][:10]}")
        self._check_matches()  # Check which records match template rows
        self._refresh_table()
        self.btn_extract.setEnabled(True)
        self.btn_extract.setText("🔍 开始提取")
        self.progress.hide()
        self.btn_export.setVisible(True)
        
        matched = sum(1 for r in records if r.get("_match_row"))
        new = len(records) - matched
        self.lbl_status.setText(f"完成：{len(records)} 条记录 — {matched} 条匹配现有行，{new} 条将新增")
    
    def _check_matches(self):
        """Check each record against template to find matching rows"""
        try:
            wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=True)
            ws = wb["Sheet1"]
            
            # Build index: 野外编号 → row number
            index = {}
            for row in range(2, ws.max_row + 1):
                c_val = ws[f"C{row}"].value
                if c_val:
                    index[str(c_val).strip()] = row
            wb.close()
            
            for r in self.records:
                fid = r.get("野外编号")
                if fid and str(fid).strip() in index:
                    r["_match_row"] = index[str(fid).strip()]
                    r["_match_status"] = "已匹配"
                else:
                    r["_match_row"] = None
                    r["_match_status"] = "新增" if fid else "无编号"
        except Exception as e:
            for r in self.records:
                r["_match_row"] = None
                r["_match_status"] = "?"
    
    def _refresh_table(self):
        self.table.setRowCount(0)
        self.table.setRowCount(len(self.records))
        for i, r in enumerate(self.records):
            # Status column
            status = r.get("_match_status", "")
            status_item = QTableWidgetItem(status)
            if status == "已匹配":
                status_item.setForeground(QColor("#a6e3a1"))
            elif status == "新增":
                status_item.setForeground(QColor("#89b4fa"))
            elif status == "无编号":
                status_item.setForeground(QColor("#f38ba8"))
            self.table.setItem(i, 0, status_item)
            
            # Match row
            match_row = str(r.get("_match_row", "") or "")
            match_item = QTableWidgetItem(match_row)
            self.table.setItem(i, 1, match_item)
            
            # Source file
            src = str(r.get("_source", "") or "")
            src_item = QTableWidgetItem(src[:25])
            src_item.setForeground(QColor("#6c7086"))
            self.table.setItem(i, 2, src_item)
            
            # Data fields
            for j, (field, _) in enumerate(FIELDS):
                val = str(r.get(field, "") or "")
                item = QTableWidgetItem(val)
                if not val:
                    item.setForeground(QColor("#6c7086"))
                self.table.setItem(i, j + 3, item)
    
    def _cell_changed(self, row, col):
        if row < len(self.records) and col >= 3:
            field = FIELDS[col - 3][0]
            val = self.table.item(row, col).text() if self.table.item(row, col) else ""
            self.records[row][field] = val
    
    def _clear(self):
        self.files = []
        self.records = []
        self.file_tree.clear()
        self._refresh_table()
        self.btn_export.setVisible(False)
        self.lbl_status.setText("已清空")
        self.lbl_count.setText("就绪")
    
    def _export(self):
        if not self.records:
            QMessageBox.warning(self, "提示", "没有数据可导出")
            return
        
        path, _ = QFileDialog.getSaveFileName(
            self, "导出填充表", f"地下水统测表_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "Excel (*.xlsx)")
        if not path:
            return
        
        try:
            from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
            from copy import copy as copy_style
            
            col_map = {f[0]: f[1] for f in FIELDS}
            wb = openpyxl.load_workbook(TEMPLATE_PATH)
            ws = wb["Sheet1"]
            
            # Get reference style from row 2
            ref_cell = ws.cell(row=2, column=3)  # C2 as template
            template_font = Font(name='仿宋', size=10)
            template_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            template_border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            
            updated = 0
            added = 0
            with_coords = 0
            seen_ids = set()
            debug_log = []
            
            for record in self.records:
                fid = str(record.get("野外编号", "")).strip()
                # Skip garbage IDs
                if not fid or fid in seen_ids or fid == 'None': continue
                if not re.search(r'(TYXB\d+|TSK\d+)', fid):
                    print(f"[EXPORT] Skipping invalid ID: {fid}")
                    continue
                seen_ids.add(fid)
                has_coords = record.get("_has_coords")
                if has_coords: with_coords += 1
                
                target = record.get("_match_row")
                if target:
                    updated += 1
                else:
                    target = ws.max_row + 1
                    added += 1
                    ws.row_dimensions[target].height = 36
                
                lng_val = record.get("经度")
                x_val = record.get("X")
                debug_log.append(f"  {fid} row={target} has_coords={has_coords} 经度={str(lng_val)[:20] if lng_val else 'EMPTY'} X={x_val}")
                
                for key, col in col_map.items():
                    val = record.get(key)
                    if val is None or val == '': continue
                    
                    cell = ws[f"{col}{target}"]
                    
                    # Apply template formatting
                    cell.font = template_font
                    cell.alignment = template_align
                    cell.border = template_border
                    
                    # Write value
                    if key == "水位标高":
                        # Formula: =L{row}-V{row}
                        cell.value = f"=L{target}-V{target}"
                    elif key in ("井台高度","井深","测点距地面高度","地下水位埋深","测点距水面距离","地面高程","测点高程","X","Y"):
                        try: cell.value = float(val)
                        except: cell.value = str(val)
                    elif key == "调查日期":
                        cell.value = str(val).replace('-','')[:8]
                    else:
                        cell.value = str(val)
            
            print("[EXPORT DEBUG] Records to export:")
            for line in debug_log:
                print(line)
            print(f"[EXPORT DEBUG] Total: updated={updated} added={added} with_coords={with_coords}")
            
            wb.save(path)
            QMessageBox.information(self, "导出成功", 
                f"已保存到：\n{path}\n\n更新 {updated} 行，新增 {added} 行\n其中 {with_coords} 条已自动填入坐标")
            self.lbl_status.setText(f"已导出 — 更新{updated}行 新增{added}行 坐标{with_coords}条")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))
    
    def _about(self):
        QMessageBox.about(self, "关于",
            "地下水统测信息提取器 v1.0\n\n"
            "从调查App截图中自动提取井点数据\n"
            "支持图片OCR识别 + Excel文件读取\n"
            "输出符合统测表模板格式")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(DARK_STYLE)
    
    font = QFont("Microsoft YaHei", 9)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
