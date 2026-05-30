
"""
Well Extractor v2 — Light Luxury Minimalist UI with EasyOCR engine.
"""
import sys, os, re, glob, shutil, tempfile
from datetime import datetime
import numpy as np
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


def normalize_date(val):
    """Normalize any date string to YYYYMMDD (8 digits).

    Handles: YYYY-MM-DD, YYYYMMDD, YYYY/MM/DD, YYYY年MM月DD日,
    and malformed OCR output like 2026-052-1.
    """
    if not val:
        return val
    s = str(val).strip()
    # Already clean 8-digit
    if re.fullmatch(r'\d{8}', s):
        return s
    # Extract 4-digit year, then collect remaining digits
    m = re.match(r'(\d{4})[^\d]*(\d{1,2})[^\d]*(\d{1,2})', s)
    if m:
        return f"{m.group(1)}{int(m.group(2)):02d}{int(m.group(3)):02d}"
    # Last resort: strip all non-digits and take first 8
    digits = re.sub(r'\D', '', s)
    if len(digits) >= 8:
        return digits[:8]
    return s

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QProgressBar, QLabel, QFileDialog, QMessageBox, QSplitter,
    QFrame, QGraphicsDropShadowEffect, QAbstractItemView, QGroupBox,
    QInputDialog, QLineEdit, QRadioButton, QButtonGroup
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QPropertyAnimation, QEasingCurve, QSize
from PySide6.QtGui import QFont, QColor, QPalette, QLinearGradient, QBrush

from ocr_engine import extract_from_image_easyocr, extract_from_image_mimo

# ── Paths ───────────────────────────────────────────────────────────
COORD_PATH = r"D:\地下水\坐标.xlsx"
TEMPLATE_PATH = r"D:\地下水\大同朔州忻州_地下水统测表.xlsx"
SCREENSHOT_DIR = r"D:\\地下水\\信息截图"

# ── Field mapping ───────────────────────────────────────────────────
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

FIXED_VALUES = {
    "路线编号": "LX-3", "地面高程获取方法": "RTK", "测点高程获取方法": "RTK",
    "统测实施单位": "山西省地质调查院有限公司",
    "测量人": "杨光", "记录人": "杨升伟", "审核人": "杨升伟",
    "照片编号": "/", "视频编号": "/", "备注": "/",
    "所属统测区类型": "一般统测区", "统测期": "地下水平水位期",
    "统测期次类型": "一期统测", "地下水资源区名称": "天桥岩溶",
}


# ── Coordinate loading ──────────────────────────────────────────────
_coord_index = None




def decimal_to_dms(dd, is_latitude=True):
    """Convert decimal degrees to DMS string like 111°37'08.7369\" or 39°00'38.0697\""""
    if dd is None or dd == '':
        return None
    try:
        dd = float(dd)
    except (ValueError, TypeError):
        return None
    prefix = 'N' if is_latitude else 'E'
    if dd < 0:
        prefix = 'S' if is_latitude else 'W'
        dd = abs(dd)
    degrees = int(dd)
    minutes_full = (dd - degrees) * 60
    minutes = int(minutes_full)
    seconds = (minutes_full - minutes) * 60
    return '{:s}{:d}°{:02d}\'{:07.4f}\"'.format(prefix, degrees, minutes, seconds)


def load_coord_index():
    global _coord_index
    if _coord_index is not None:
        return _coord_index
    _coord_index = {}
    if not os.path.exists(COORD_PATH):
        return _coord_index
    wb = openpyxl.load_workbook(COORD_PATH, data_only=True)
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in range(1, ws.max_row + 1):
            fid = ws.cell(row=row, column=1).value
            if not fid: continue
            fid = str(fid).strip()
            def _c(v):
                if v is None: return None
                if isinstance(v, str): return v.strip().replace('\t','').replace('\n','')
                return v
            _coord_index[fid] = {
                "X": _c(ws.cell(row=row, column=3).value),
                "Y": _c(ws.cell(row=row, column=4).value),
                "地面高程": _c(ws.cell(row=row, column=5).value),
                "经度": _c(ws.cell(row=row, column=6).value),
                "纬度": _c(ws.cell(row=row, column=7).value),
            }
    wb.close()
    return _coord_index

def fuzzy_match_id(ocr_id, coord_ids):
    """Match OCR field ID to coordinate file ID."""
    if ocr_id in coord_ids:
        return ocr_id
    # Strip trailing -digit suffix and retry
    base = re.sub(r'-\d+$', '', ocr_id)
    if base in coord_ids:
        return base
    # Check if any coord ID starts with the OCR ID (e.g. "THK008" matches "THK008-1")
    for cid in coord_ids:
        if cid.startswith(ocr_id) or ocr_id.startswith(cid):
            return cid
    # Strip suffix from both sides
    for cid in coord_ids:
        cid_base = re.sub(r'-\d+$', '', cid)
        if base == cid_base:
            return cid
    return None

def batch_merge_coordinates(records, coord_path=None):
    """Match coordinates by 野外编号. OCR坐标作废，只用坐标文件数据。"""
    if coord_path:
        global _coord_index
        _coord_index = None
        global COORD_PATH
        COORD_PATH = coord_path
    coords = load_coord_index()
    if not coords:
        return records
    coord_ids = list(coords.keys())
    matched = 0
    for r in records:
        fid = r.get("野外编号", "").strip()
        if not fid:
            continue
        match_id = fuzzy_match_id(fid, coord_ids)
        if match_id:
            c = coords[match_id]
            for k, v in c.items():
                if v is not None and v != '':
                    r[k] = v
            r["_has_coords"] = True
            matched += 1
    print(f"坐标匹配: {matched}/{len(records)}")
    return records

# ── Neon Cyberpunk Stylesheet ───────────────────────────────────────
NEON_STYLE = """
QMainWindow {
    background: #FAFAFA;
}
QMenuBar {
    background: #FFFFFF;
    color: #333333;
    border-bottom: 1px solid #E8E8E8;
    padding: 4px;
    font-size: 13px;
}
QMenuBar::item:selected {
    background: #F0F0F0;
    border-radius: 4px;
}
QMenu {
    background: #FFFFFF;
    color: #333333;
    border: 1px solid #E0E0E0;
    border-radius: 8px;
    padding: 4px;
}
QMenu::item:selected {
    background: #F5F5F5;
    border-radius: 4px;
}
QPushButton {
    background: #FFFFFF;
    color: #555555;
    border: 1px solid #D0D0D0;
    border-radius: 6px;
    padding: 7px 18px;
    font-size: 13px;
}
QPushButton:hover {
    background: #F5F5F5;
    border: 1px solid #B0B0B0;
}
QPushButton:pressed {
    background: #EBEBEB;
}
QPushButton#btnExtract {
    background: #2C2C2C;
    color: #FFFFFF;
    border: none;
    font-size: 14px;
    padding: 9px 28px;
    font-weight: bold;
    letter-spacing: 1px;
}
QPushButton#btnExtract:hover {
    background: #444444;
}
QPushButton#btnExtract:disabled {
    background: #C0C0C0;
    color: #888888;
}
QPushButton#btnExport {
    background: #1A1A1A;
    color: #FFFFFF;
    border: none;
    font-size: 14px;
    padding: 9px 28px;
    font-weight: bold;
}
QPushButton#btnExport:hover {
    background: #333333;
}
QPushButton#btnCoord {
    background: #FFFFFF;
    color: #666666;
    border: 1px solid #D0D0D0;
    border-radius: 6px;
    padding: 7px 18px;
    font-size: 13px;
}
QPushButton#btnCoord:hover {
    background: #F5F5F5;
    border: 1px solid #B0B0B0;
}
QPushButton#btnImport {
    background: #FFFFFF;
    color: #888888;
    border: 1.5px dashed #C0C0C0;
    font-size: 15px;
    padding: 18px;
    border-radius: 10px;
}
QPushButton#btnImport:hover {
    border: 1.5px dashed #999999;
    color: #555555;
    background: #FAFAFA;
}
QProgressBar {
    background: #F0F0F0;
    border: none;
    border-radius: 4px;
    text-align: center;
    color: #666666;
    font-size: 11px;
    height: 22px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #2C2C2C, stop:1 #555555);
    border-radius: 4px;
}
QTableWidget {
    background: #FFFFFF;
    color: #333333;
    border: 1px solid #E8E8E8;
    border-radius: 8px;
    gridline-color: #F0F0F0;
    font-size: 12px;
    alternate-background-color: #FAFAFA;
}
QTableWidget::item {
    padding: 5px 10px;
    border-bottom: 1px solid #F0F0F0;
}
QTableWidget::item:selected {
    background: #E8E8E8;
    color: #1A1A1A;
}
QHeaderView::section {
    background: #F8F8F8;
    color: #555555;
    border: none;
    border-bottom: 2px solid #E0E0E0;
    border-right: 1px solid #E8E8E8;
    padding: 8px 10px;
    font-weight: bold;
    font-size: 12px;
}
QGroupBox {
    color: #555555;
    border: 1px solid #E8E8E8;
    border-radius: 10px;
    margin-top: 16px;
    padding-top: 20px;
    font-weight: bold;
    font-size: 13px;
    background: #FFFFFF;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 16px;
    padding: 0 8px;
    color: #666666;
}
QLabel {
    color: #666666;
    font-size: 12px;
}
QLabel#statusLabel {
    color: #2C2C2C;
    font-size: 13px;
    font-weight: bold;
}
QLabel#titleLabel {
    color: #1A1A1A;
    font-size: 22px;
    font-weight: bold;
    letter-spacing: 2px;
}
QLabel#subtitleLabel {
    color: #999999;
    font-size: 13px;
}
QSplitter::handle {
    background: #E8E8E8;
    width: 1px;
}
QScrollBar:vertical {
    background: #FAFAFA;
    width: 7px;
    border-radius: 3px;
}
QScrollBar::handle:vertical {
    background: #CCCCCC;
    border-radius: 3px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover {
    background: #AAAAAA;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QRadioButton {
    color: #555555;
    font-size: 12px;
    spacing: 6px;
}
QRadioButton::indicator {
    width: 16px;
    height: 16px;
    border-radius: 9px;
    border: 2px solid #CCCCCC;
    background: #FFFFFF;
}
QRadioButton::indicator:checked {
    background: #2C2C2C;
    border: 2px solid #2C2C2C;
}
QFrame[frameShape="4"] {
    background: #E8E8E8;
    max-height: 1px;
}
"""

# ── Worker Thread ───────────────────────────────────────────────────
class ExtractWorker(QThread):
    progress = Signal(int)
    status = Signal(str)
    finished = Signal(object)
    error = Signal(str)

    def __init__(self, files, coord_path=None, engine="easyocr"):
        super().__init__()
        self.files = files
        self.coord_path = coord_path
        self.engine = engine  # "easyocr" or "mimo"

    def run(self):
        records = []
        total = len(self.files)
        
        if self.engine == "mimo":
            self.status.emit(f"正在连接 MiMo 视觉引擎...")
        else:
            self.status.emit(f"正在加载 OCR 引擎...")
        
        for i, f in enumerate(self.files):
            try:
                self.status.emit(f"识别中 [{i+1}/{total}]: {os.path.basename(f)[:30]}...")
                if self.engine == "mimo":
                    r = extract_from_image_mimo(f)
                else:
                    r = extract_from_image_easyocr(f)
                if r:
                    r["_source"] = os.path.basename(f)
                    # Apply fixed values
                    for k, v in FIXED_VALUES.items():
                        if k not in r or not r.get(k):
                            r[k] = v
                    # Set survey date to today if missing
                    if "调查日期" not in r or not r.get("调查日期"):
                        r["调查日期"] = datetime.now().strftime("%Y%m%d")
                    # Discard ALL OCR/screenshot coordinates — coordinate file is authoritative
                    for k in ("经度", "纬度", "地面高程", "测点高程", "X", "Y"):
                        r.pop(k, None)
                    records.append(r)
                else:
                    records.append({"_source": os.path.basename(f), "_error": "MiMo 识别失败，请检查 API Key 和网络"})
            except Exception as e:
                records.append({"_source": os.path.basename(f), "_error": str(e)})
            self.progress.emit(int((i + 1) / total * 100))
        
        # Merge coordinates
        self.status.emit("正在合并坐标数据...")
        batch_merge_coordinates(records, self.coord_path)
        
        # Template matching (moved off main thread to avoid UI freeze)
        self.status.emit("正在匹配模板...")
        matches = self._do_match_template(records)
        
        self.finished.emit((records, matches))
    
    def _do_match_template(self, records):
        """Match records to template rows. Runs in worker thread."""
        match_info = {}
        try:
            wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=True)
            ws = wb["Sheet1"]
            idx = {}
            for row in range(2, ws.max_row + 1):
                v = ws[f"C{row}"].value
                if v: idx[str(v).strip()] = row
            wb.close()
            for i, r in enumerate(records):
                fid = r.get("野外编号")
                if fid and str(fid).strip() in idx:
                    match_info[i] = (idx[str(fid).strip()], "已匹配")
                else:
                    match_info[i] = (None, "新增" if fid else "无编号")
        except Exception:
            for i in range(len(records)):
                match_info[i] = (None, "?")
        return match_info


# ── Main Window ─────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("地下水统测信息提取系统 v2.0")
        self.resize(1500, 900)
        self.setMinimumSize(1100, 650)
        self.records = []
        self.files = []
        self.coord_path = None
        self.engine = "easyocr"  # default engine
        self._setup_ui()
        self._apply_glow_effects()
    
    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(20, 16, 20, 16)
        main_layout.setSpacing(12)
        
        # ── Header ──
        header = QHBoxLayout()
        title_block = QVBoxLayout()
        title = QLabel("地下水统测信息提取器")
        title.setObjectName("titleLabel")
        subtitle = QLabel("AI-Powered OCR Engine  ·  轻奢简约版")
        subtitle.setObjectName("subtitleLabel")
        title_block.addWidget(title)
        title_block.addWidget(subtitle)
        header.addLayout(title_block)
        header.addStretch()
        
        self.status_label = QLabel("就绪 // READY")
        self.status_label.setObjectName("statusLabel")
        header.addWidget(self.status_label)
        
        main_layout.addLayout(header)
        
        # ── Toolbar ──
        toolbar = QHBoxLayout()
        toolbar.setSpacing(12)
        
        self.btn_import = QPushButton("📁 导入截图文件夹")
        self.btn_import.setObjectName("btnImport")
        self.btn_import.clicked.connect(self._import_folder)
        toolbar.addWidget(self.btn_import)
        
        # ── Engine selector ──
        engine_frame = QFrame()
        engine_frame.setStyleSheet("QFrame { background: #F8F8F8; border: 1px solid #E8E8E8; border-radius: 8px; padding: 4px 8px; }")
        engine_layout = QHBoxLayout(engine_frame)
        engine_layout.setContentsMargins(8, 2, 8, 2)
        engine_layout.setSpacing(8)
        
        engine_label = QLabel("识别引擎")
        engine_label.setStyleSheet("color: #999999; font-size: 11px; background: transparent; border: none;")
        engine_layout.addWidget(engine_label)
        
        self.engine_group = QButtonGroup(self)
        
        self.rb_easyocr = QRadioButton("PP-OCR")
        self.rb_easyocr.setChecked(True)
        
        self.rb_mimo = QRadioButton("MiMo")
        
        self.engine_group.addButton(self.rb_easyocr, 0)
        self.engine_group.addButton(self.rb_mimo, 1)
        self.engine_group.buttonClicked.connect(self._on_engine_changed)
        
        engine_layout.addWidget(self.rb_easyocr)
        engine_layout.addWidget(self.rb_mimo)
        toolbar.addWidget(engine_frame)
        
        self.btn_extract = QPushButton("⚡ 开始提取")
        self.btn_extract.setObjectName("btnExtract")
        self.btn_extract.clicked.connect(self._extract)
        self.btn_extract.setEnabled(False)
        toolbar.addWidget(self.btn_extract)
        
        self.btn_coord = QPushButton("📍 导入坐标")
        self.btn_coord.setObjectName("btnCoord")
        self.btn_coord.clicked.connect(self._import_coord)
        toolbar.addWidget(self.btn_coord)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        toolbar.addWidget(self.progress_bar, 1)
        
        self.btn_export = QPushButton("📥 导出 Excel")
        self.btn_export.setObjectName("btnExport")
        self.btn_export.clicked.connect(self._export)
        self.btn_export.setVisible(False)
        toolbar.addWidget(self.btn_export)
        
        self.btn_clear = QPushButton("🗑 清空")
        self.btn_clear.clicked.connect(self._clear)
        toolbar.addWidget(self.btn_clear)
        
        main_layout.addLayout(toolbar)
        
        # ── Content: splitter ──
        splitter = QSplitter(Qt.Horizontal)
        
        # Left: stats panel
        left = QFrame()
        left.setStyleSheet("QFrame { background: #FFFFFF; border-radius: 12px; border: 1px solid #E8E8E8; }")
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(16, 16, 16, 16)
        
        stats_group = QGroupBox("统计面板")
        stats_layout = QVBoxLayout(stats_group)
        
        self.lbl_total = QLabel("待处理文件: 0")
        self.lbl_total.setStyleSheet("font-size: 28px; font-weight: bold; color: #1A1A1A;")
        stats_layout.addWidget(self.lbl_total)
        
        self.lbl_extracted = QLabel("已提取记录: 0")
        self.lbl_extracted.setStyleSheet("font-size: 15px; color: #555555;")
        stats_layout.addWidget(self.lbl_extracted)
        
        self.lbl_coords = QLabel("含坐标记录: 0")
        self.lbl_coords.setStyleSheet("font-size: 15px; color: #555555;")
        stats_layout.addWidget(self.lbl_coords)
        
        self.lbl_matched = QLabel("模板匹配: 0")
        self.lbl_matched.setStyleSheet("font-size: 15px; color: #555555;")
        stats_layout.addWidget(self.lbl_matched)
        
        stats_layout.addStretch()
        left_layout.addWidget(stats_group)
        
        # Recent files list
        files_group = QGroupBox("文件列表")
        files_layout = QVBoxLayout(files_group)
        self.lbl_files = QLabel("拖拽或导入截图文件夹")
        self.lbl_files.setWordWrap(True)
        self.lbl_files.setStyleSheet("color: #999999; font-size: 13px;")
        files_layout.addWidget(self.lbl_files)
        left_layout.addWidget(files_group)
        
        splitter.addWidget(left)
        
        # Right: data table
        right = QFrame()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        self.table = QTableWidget()
        self.table.setColumnCount(len(FIELDS) + 3)
        self.table.setHorizontalHeaderLabels(
            ["状态", "模板行", "来源"] + [f[0] for f in FIELDS]
        )
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        right_layout.addWidget(self.table)
        
        splitter.addWidget(right)
        splitter.setSizes([280, 1220])
        
        main_layout.addWidget(splitter, 1)
    
    def _apply_glow_effects(self):
        """Apply subtle shadow to primary buttons."""
        for btn_name in ["btnExtract", "btnExport"]:
            btn = self.findChild(QPushButton, btn_name)
            if btn:
                shadow = QGraphicsDropShadowEffect()
                shadow.setBlurRadius(16)
                shadow.setColor(QColor(0, 0, 0, 40))
                shadow.setOffset(0, 2)
                btn.setGraphicsEffect(shadow)
    
    def _import_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择截图文件夹", SCREENSHOT_DIR)
        if not folder:
            return
        self.files = []
        for ext in ('*.jpg', '*.jpeg', '*.png', '*.bmp'):
            self.files.extend(glob.glob(os.path.join(folder, ext)))
        self.files = sorted(self.files)
        
        # Update UI
        file_list = "<br>".join(f"▸ {os.path.basename(f)[:50]}..." for f in self.files[:20])
        if len(self.files) > 20:
            file_list += f"<br><br>... 及其他 {len(self.files) - 20} 个文件"
        self.lbl_files.setText(file_list)
        self.lbl_total.setText(f"待处理文件: {len(self.files)}")
        self.btn_extract.setEnabled(len(self.files) > 0)
        self.btn_export.setVisible(False)
        self.status_label.setText(f"已加载 {len(self.files)} 个文件 // LOADED")
    
    def _import_coord(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择坐标文件", "D:/地下水", "Excel (*.xlsx *.xls)")
        if not path:
            return
        self.coord_path = path
        self.lbl_coords.setText("坐标文件: {}".format(os.path.basename(path)))
        self.status_label.setText("坐标已加载 // {}".format(os.path.basename(path)))
    
    def _on_engine_changed(self, btn):
        """Handle engine radio button selection."""
        if btn == self.rb_easyocr:
            self.engine = "easyocr"
            self.status_label.setText("引擎: PP-OCR // 本地离线")
        else:
            self.engine = "mimo"
            self.status_label.setText("引擎: MiMo // 云端视觉")
    
    def _extract(self):
        if not self.files:
            return
        self.btn_extract.setEnabled(False)
        self.btn_import.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("提取中... // PROCESSING")
        
        self.worker = ExtractWorker(self.files, self.coord_path, engine=self.engine)
        self.worker.progress.connect(self._on_progress)
        self.worker.status.connect(lambda s: self.status_label.setText(s))
        self.worker.finished.connect(self._on_finished)
        self.worker.start()
    
    def _on_progress(self, val):
        self.progress_bar.setValue(val)
    
    def _on_finished(self, result):
        records, matches = result
        self.records = records
        # Apply template match results from worker thread
        for i, (row, status) in matches.items():
            records[i]["_match_row"] = row
            records[i]["_match_status"] = status
        self._refresh_table()
        
        self.btn_extract.setEnabled(True)
        self.btn_import.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.btn_export.setVisible(True)
        
        coords = sum(1 for r in records if r.get("_has_coords"))
        matched = sum(1 for r in records if r.get("_match_row"))
        
        self.lbl_extracted.setText(f"已提取记录: {len(records)}")
        self.lbl_coords.setText(f"含坐标记录: {coords}")
        self.lbl_matched.setText(f"模板匹配: {matched}")
        
        self.status_label.setText(f"完成: {len(records)} 条 // {coords} 有坐标 // DONE")
    
    def _check_matches(self):
        try:
            wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=True)
            ws = wb["Sheet1"]
            idx = {}
            for row in range(2, ws.max_row + 1):
                v = ws[f"C{row}"].value
                if v: idx[str(v).strip()] = row
            wb.close()
            for r in self.records:
                fid = r.get("野外编号")
                if fid and str(fid).strip() in idx:
                    r["_match_row"] = idx[str(fid).strip()]
                    r["_match_status"] = "已匹配"
                else:
                    r["_match_row"] = None
                    r["_match_status"] = "新增" if fid else "无编号"
        except:
            for r in self.records:
                r["_match_row"] = None
                r["_match_status"] = "?"

    def _refresh_table(self):
        self.table.setUpdatesEnabled(False)  # batch mode — no per-item repaint
        self.table.setRowCount(len(self.records))
        for i, r in enumerate(self.records):
            # Status
            st = r.get("_match_status", "")
            si = QTableWidgetItem(st)
            if st == "已匹配": si.setForeground(QColor("#2E7D32"))
            elif st == "新增": si.setForeground(QColor("#1565C0"))
            elif st == "无编号": si.setForeground(QColor("#C62828"))
            self.table.setItem(i, 0, si)
            # Row
            mr = str(r.get("_match_row", "") or "")
            self.table.setItem(i, 1, QTableWidgetItem(mr))
            # Source
            src = str(r.get("_source", "") or "")[:25]
            sri = QTableWidgetItem(src)
            sri.setForeground(QColor("#999999"))
            self.table.setItem(i, 2, sri)
            # Fields
            for j, (fn, _) in enumerate(FIELDS):
                val = str(r.get(fn, "") or "")
                item = QTableWidgetItem(val)
                if not val:
                    item.setForeground(QColor("#CCCCCC"))
                if fn == "野外编号" and val:
                    item.setForeground(QColor("#1565C0"))
                if fn in ("经度", "纬度") and val:
                    item.setForeground(QColor("#6A1B9A"))
                self.table.setItem(i, j + 3, item)
            if i % 10 == 0:
                QApplication.processEvents()  # keep UI responsive every 10 rows
        self.table.setUpdatesEnabled(True)  # single repaint for all changes
    
    def _clear(self):
        self.records = []
        self.files = []
        self.table.setRowCount(0)
        self.btn_export.setVisible(False)
        self.lbl_total.setText("待处理文件: 0")
        self.lbl_extracted.setText("已提取记录: 0")
        self.lbl_coords.setText("含坐标记录: 0")
        self.lbl_matched.setText("模板匹配: 0")
        self.status_label.setText("就绪 // READY")

    def _export(self):
        if not self.records:
            QMessageBox.warning(self, "\u63d0\u793a", "\u6ca1\u6709\u6570\u636e\u53ef\u5bfc\u51fa")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "\u5bfc\u51fa\u586b\u5145\u8868",
            f"\u5730\u4e0b\u6c34\u7edf\u6d4b\u8868_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "Excel (*.xlsx)")
        if not path: return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Sheet1"

            # Styles
            tf = Font(name='\u4eff\u5b8b', size=10)
            ta = Alignment(horizontal='center', vertical='center', wrap_text=True)
            tb = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))
            header_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

            # Headers (36 columns matching reference format)
            headers = [
                '\u5e8f\u53f7', '\u7edf\u4e00\u7f16\u53f7', '\u91ce\u5916\u7f16\u53f7',
                '\u8def\u7ebf\u7f16\u53f7', '\u8c03\u67e5\u65e5\u671f', '\u5730\u7406\u4f4d\u7f6e',
                '\u5929\u6c14', 'E(\u7ecf\u5ea6)', 'N\uff08\u7eac\u5ea6\uff09',
                'X', 'Y', '\u5730\u9762\u9ad8\u7a0b(m)',
                '\u5730\u9762\u9ad8\u7a0b\u83b7\u53d6\u65b9\u6cd5',
                '\u6d4b\u70b9\u9ad8\u7a0b\uff08m\uff09',
                '\u6d4b\u70b9\u9ad8\u7a0b\u83b7\u53d6\u65b9\u6cd5',
                '\u4e95\u53f0\u9ad8\u5ea6\uff08m\uff09',
                '\u4e95\u6df1\uff08m\uff09',
                '\u5730\u4e0b\u6c34\u8d44\u6e90\u533a\u540d\u79f0',
                '\u6240\u5c5e\u7edf\u6d4b\u533a\u7c7b\u578b',
                '\u4e95\u70b9\u7c7b\u578b',
                '\u6d4b\u70b9\u8ddd\u5730\u9762\u9ad8\u5ea6\uff08m\uff09',
                '\u5730\u4e0b\u6c34\u4f4d\u57cb\u6df1\uff08m\uff09',
                '\u6d4b\u70b9\u8ddd\u6c34\u9762\u8ddd\u79bb\uff08m\uff09',
                '\u6c34\u4f4d\u6807\u9ad8\uff08m\uff09',
                '\u5730\u4e0b\u6c34\u7c7b\u578b\uff08\u6309\u542b\u6c34\u4ecb\u8d28\uff09',
                '\u5730\u4e0b\u6c34\u7c7b\u578b\uff08\u6309\u57cb\u85cf\u6df1\u5ea6\uff09',
                '\u5730\u4e0b\u6c34\u7c7b\u578b\uff08\u6309\u627f\u538b\u6027\uff09',
                '\u7edf\u6d4b\u671f',
                '\u7edf\u6d4b\u533a\u6240\u5c5e\u7edf\u6d4b\u671f\u6b21\u7c7b\u578b',
                '\u7167\u7247\u7f16\u53f7', '\u89c6\u9891\u7f16\u53f7', '\u5907\u6ce8',
                '\u7edf\u6d4b\u5b9e\u65bd\u5355\u4f4d',
                '\u6d4b\u91cf\u4eba', '\u8bb0\u5f55\u4eba', '\u5ba1\u6838\u4eba',
            ]

            # Write headers with formatting
            for ci, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=ci, value=h)
                cell.font = Font(name='\u4eff\u5b8b', size=10)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
                cell.fill = header_fill
            ws.row_dimensions[1].height = 26

            # Column widths
            widths = {
                'A': 5, 'B': 12, 'C': 19, 'D': 12, 'E': 11, 'F': 51,
                'G': 6, 'H': 26, 'I': 21, 'J': 15, 'K': 18, 'L': 12,
                'M': 9, 'N': 15, 'O': 9, 'P': 9, 'Q': 10, 'R': 20,
                'S': 11, 'T': 9, 'U': 13, 'V': 12, 'W': 12, 'X': 13,
                'Y': 20, 'Z': 18, 'AA': 14, 'AB': 14, 'AC': 16,
                'AD': 11, 'AE': 11, 'AF': 11, 'AG': 22, 'AH': 10,
                'AI': 10, 'AJ': 10,
            }
            for letter, w in widths.items():
                ws.column_dimensions[letter].width = w

            # Field -> column mapping
            col_map = {
                '\u7edf\u4e00\u7f16\u53f7': 'B', '\u91ce\u5916\u7f16\u53f7': 'C',
                '\u8def\u7ebf\u7f16\u53f7': 'D', '\u8c03\u67e5\u65e5\u671f': 'E',
                '\u5730\u7406\u4f4d\u7f6e': 'F', '\u5929\u6c14': 'G',
                '\u7ecf\u5ea6': 'H', '\u7eac\u5ea6': 'I',
                'X': 'J', 'Y': 'K',
                '\u5730\u9762\u9ad8\u7a0b': 'L',
                '\u5730\u9762\u9ad8\u7a0b\u83b7\u53d6\u65b9\u6cd5': 'M',
                '\u6d4b\u70b9\u9ad8\u7a0b': 'N',
                '\u6d4b\u70b9\u9ad8\u7a0b\u83b7\u53d6\u65b9\u6cd5': 'O',
                '\u4e95\u53f0\u9ad8\u5ea6': 'P', '\u4e95\u6df1': 'Q',
                '\u5730\u4e0b\u6c34\u8d44\u6e90\u533a\u540d\u79f0': 'R',
                '\u6240\u5c5e\u7edf\u6d4b\u533a\u7c7b\u578b': 'S',
                '\u4e95\u70b9\u7c7b\u578b': 'T',
                '\u6d4b\u70b9\u8ddd\u5730\u9762\u9ad8\u5ea6': 'U',
                '\u5730\u4e0b\u6c34\u4f4d\u57cb\u6df1': 'V',
                '\u6d4b\u70b9\u8ddd\u6c34\u9762\u8ddd\u79bb': 'W',
                '\u6c34\u4f4d\u6807\u9ad8': 'X',
                '\u542b\u6c34\u4ecb\u8d28': 'Y', '\u57cb\u85cf\u6df1\u5ea6': 'Z',
                '\u627f\u538b\u6027': 'AA',
                '\u7edf\u6d4b\u671f': 'AB',
                '\u7edf\u6d4b\u671f\u6b21\u7c7b\u578b': 'AC',
                '\u7167\u7247\u7f16\u53f7': 'AD', '\u89c6\u9891\u7f16\u53f7': 'AE',
                '\u5907\u6ce8': 'AF',
                '\u7edf\u6d4b\u5b9e\u65bd\u5355\u4f4d': 'AG',
                '\u6d4b\u91cf\u4eba': 'AH', '\u8bb0\u5f55\u4eba': 'AI',
                '\u5ba1\u6838\u4eba': 'AJ',
            }
            numeric_keys = ('\u4e95\u53f0\u9ad8\u5ea6', '\u4e95\u6df1',
                            '\u6d4b\u70b9\u8ddd\u5730\u9762\u9ad8\u5ea6',
                            '\u5730\u4e0b\u6c34\u4f4d\u57cb\u6df1',
                            '\u6d4b\u70b9\u8ddd\u6c34\u9762\u8ddd\u79bb',
                            '\u5730\u9762\u9ad8\u7a0b', '\u6d4b\u70b9\u9ad8\u7a0b',
                            'X', 'Y')

            seen_ids = set()
            row_num = 2
            added = 0
            for idx, record in enumerate(self.records):
                fid = str(record.get('\u91ce\u5916\u7f16\u53f7', '')).strip()
                if not fid or fid in seen_ids or fid == 'None':
                    continue
                seen_ids.add(fid)
                added += 1

                ws.row_dimensions[row_num].height = 36
                # 序号
                ws.cell(row=row_num, column=1, value=idx + 1)

                for key, col in col_map.items():
                    if col in ('A',):
                        continue
                    val = record.get(key)
                    col_idx = openpyxl.utils.column_index_from_string(col)
                    cell = ws.cell(row=row_num, column=col_idx)
                    cell.font = Font(name='\u4eff\u5b8b', size=10)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

                    # Longitude -> DMS
                    if key == '\u7ecf\u5ea6' and val:
                        dms = decimal_to_dms(val, is_latitude=False)
                        cell.value = dms if dms else str(val)
                    # Latitude -> DMS
                    elif key == '\u7eac\u5ea6' and val:
                        dms = decimal_to_dms(val, is_latitude=True)
                        cell.value = dms if dms else str(val)
                    # 水位标高 = 地面高程 - 地下水位埋深
                    elif key == '\u6c34\u4f4d\u6807\u9ad8':
                        elev = record.get('\u5730\u9762\u9ad8\u7a0b')
                        depth = record.get('\u5730\u4e0b\u6c34\u4f4d\u57cb\u6df1')
                        if elev and depth:
                            cell.value = f"=L{row_num}-V{row_num}"
                        else:
                            cell.value = None
                    # 测点高程 = 地面高程
                    elif key == '\u6d4b\u70b9\u9ad8\u7a0b':
                        elev = record.get('\u5730\u9762\u9ad8\u7a0b')
                        if elev is not None and elev != '':
                            try: cell.value = float(elev)
                            except: cell.value = str(elev)
                    elif val is None or val == '':
                        cell.value = None
                    elif key == '\u8c03\u67e5\u65e5\u671f':
                        cell.value = normalize_date(val)
                    elif key in numeric_keys:
                        try: cell.value = float(val)
                        except: cell.value = str(val)
                    else:
                        cell.value = str(val)

                row_num += 1

            wb.save(path)
            QMessageBox.information(self, "\u5bfc\u51fa\u6210\u529f",
                f"\u5df2\u4fdd\u5b58\u5230:\n{path}\n\n\u5171 {added} \u6761\u8bb0\u5f55")
            self.status_label.setText(f"\u5df2\u5bfc\u51fa // EXPORTED: {added} rows")
        except Exception as e:
            QMessageBox.critical(self, "\u5bfc\u51fa\u5931\u8d25", str(e))

from PySide6.QtWidgets import QDialog as _QDialog

class LicenseDialog(_QDialog):
    """Custom license dialog with copyable machine code."""
    def __init__(self, machine_code, parent=None):
        super().__init__(parent)
        self.machine_code = machine_code
        self.setWindowTitle("🔐 授权验证")
        self.setFixedSize(520, 340)
        self._result = None
        self._ui()
    
    def _ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(30, 24, 30, 24)
        
        title = QLabel("软件授权验证")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #1A1A1A; font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        
        # Machine code row
        layout.addWidget(QLabel("机器码（发送给管理员获取密钥）："))
        
        mc_row = QHBoxLayout()
        self.mc_box = QLineEdit(self.machine_code)
        self.mc_box.setReadOnly(True)
        self.mc_box.setAlignment(Qt.AlignCenter)
        self.mc_box.setStyleSheet("""
            QLineEdit {
                background: #F8F8F8; color: #333333; border: 1px solid #E0E0E0;
                border-radius: 6px; padding: 10px 14px; font-size: 15px;
                font-weight: bold; letter-spacing: 2px;
            }
        """)
        mc_row.addWidget(self.mc_box)
        
        btn_mc_copy = QPushButton("📋 复制")
        btn_mc_copy.setFixedWidth(80)
        btn_mc_copy.clicked.connect(lambda: self._copy_to_clip(self.machine_code, "机器码"))
        mc_row.addWidget(btn_mc_copy)
        layout.addLayout(mc_row)
        
        self._mc_feedback = QLabel("")
        self._mc_feedback.setStyleSheet("color: #2E7D32; font-size: 11px;")
        layout.addWidget(self._mc_feedback)
        
        layout.addSpacing(8)
        
        # Key input
        layout.addWidget(QLabel("请输入授权密钥："))
        
        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("XXXX-XXXX-XXXX-XXXX")
        self.key_input.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.key_input)
        
        layout.addSpacing(10)
        
        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        
        btn_activate = QPushButton("⚡ 激活")
        btn_activate.setStyleSheet("""
            QPushButton {
                background: #2C2C2C;
                color: #FFFFFF; border: none;
                border-radius: 8px; padding: 10px 28px;
                font-size: 15px; font-weight: bold;
            }
            QPushButton:hover {
                background: #444444;
            }
        """)
        btn_activate.clicked.connect(self._activate)
        btn_row.addWidget(btn_activate)
        btn_row.addStretch()
        layout.addLayout(btn_row)
        
        layout.addStretch()
    
    def _copy_to_clip(self, text, label):
        QApplication.clipboard().setText(text)
        self._mc_feedback.setText(f"✅ {label}已复制到剪贴板!")
    
    def _activate(self):
        self._result = self.key_input.text().strip()
        if not self._result:
            QMessageBox.warning(self, "提示", "请输入授权密钥")
            return
        self.accept()
    
    def get_result(self):
        return self._result


# ── Entry ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    from license import get_machine_code, check_license, is_licensed, save_license
    
    app = QApplication(sys.argv)
    
    # License check
    if not is_licensed():
        mc = get_machine_code()
        dlg = LicenseDialog(mc)
        if dlg.exec() == _QDialog.Accepted:
            key = dlg.get_result()
            if key and check_license(key, mc):
                save_license(key)
                QMessageBox.information(None, "成功", "授权成功！")
            else:
                QMessageBox.critical(None, "错误", "授权密钥无效，程序退出。")
                sys.exit(1)
        else:
            sys.exit(1)
    
    app.setStyle("Fusion")
    app.setStyleSheet(NEON_STYLE)
    font = QFont("Microsoft YaHei", 9)
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
