
"""
Well Extractor v2 — Futuristic Cyberpunk UI with EasyOCR engine.
"""
import sys, os, re, glob, shutil, tempfile
from datetime import datetime
import numpy as np
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

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
        for row in range(2, ws.max_row + 1):
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
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #0a0a1a, stop:0.5 #0d0d2b, stop:1 #0a0a1a);
}
QMenuBar {
    background: rgba(10, 10, 30, 0.95);
    color: #00f0ff;
    border-bottom: 1px solid #00f0ff44;
    padding: 4px;
    font-size: 13px;
}
QMenuBar::item:selected {
    background: #00f0ff22;
    border: 1px solid #00f0ff66;
    border-radius: 4px;
}
QMenu {
    background: #0d0d2b;
    color: #e0e0ff;
    border: 1px solid #00f0ff44;
    border-radius: 8px;
    padding: 4px;
}
QMenu::item:selected {
    background: #00f0ff22;
    border-radius: 4px;
}
QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00f0ff22, stop:1 #ff00ff22);
    color: #00f0ff;
    border: 1px solid #00f0ff55;
    border-radius: 8px;
    padding: 8px 20px;
    font-size: 13px;
    font-weight: bold;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00f0ff44, stop:1 #ff00ff44);
    border: 1px solid #00f0ffaa;
}
QPushButton:pressed {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00f0ff66, stop:1 #ff00ff66);
}
QPushButton#btnExtract {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00ff8844, stop:1 #00f0ff44);
    color: #ffffff;
    font-size: 14px;
    padding: 10px 30px;
    text-shadow: 0 0 8px #00ff88;
}
QPushButton#btnExtract:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00ff8866, stop:1 #00f0ff66);
}
QPushButton#btnExport {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #ff00ff44, stop:1 #ff880044);
    color: #fff;
    font-size: 14px;
    padding: 10px 30px;
}
QPushButton#btnCoord {
    background: #ff880022;
    color: #ff8800;
    border: 1px solid #ff880055;
    border-radius: 8px;
    padding: 8px 20px;
    font-size: 13px;
    font-weight: bold;
}
QPushButton#btnCoord:hover {
    background: #ff880044;
    border: 1px solid #ff8800aa;
}
QPushButton#btnImport {
    background: #00f0ff22;
    color: #00f0ff;
    border: 1px dashed #00f0ff55;
    font-size: 16px;
    padding: 20px;
    border-radius: 12px;
}
QPushButton#btnImport:hover {
    background: #00f0ff33;
    border: 1px solid #00f0ff88;
}
QProgressBar {
    background: #0a0a1a;
    border: 1px solid #00f0ff33;
    border-radius: 6px;
    text-align: center;
    color: #00f0ff;
    font-weight: bold;
    height: 18px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00f0ff, stop:0.5 #00ff88, stop:1 #ff00ff);
    border-radius: 5px;
}
QTableWidget {
    background: rgba(10, 10, 30, 0.8);
    color: #e0e0ff;
    border: 1px solid #00f0ff22;
    border-radius: 8px;
    gridline-color: #00f0ff11;
    font-size: 12px;
    alternate-background-color: rgba(0, 240, 255, 0.03);
}
QTableWidget::item {
    padding: 6px 10px;
    border-bottom: 1px solid #00f0ff08;
}
QTableWidget::item:selected {
    background: #00f0ff22;
    color: #fff;
}
QHeaderView::section {
    background: #0d0d2b;
    color: #00f0ff;
    border: none;
    border-bottom: 2px solid #00f0ff44;
    padding: 8px 10px;
    font-weight: bold;
    font-size: 12px;
}
QGroupBox {
    color: #00f0ff;
    border: 1px solid #00f0ff22;
    border-radius: 10px;
    margin-top: 16px;
    padding-top: 20px;
    font-weight: bold;
    font-size: 13px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 16px;
    padding: 0 8px;
    color: #00f0ff;
}
QLabel {
    color: #c0c0f0;
    font-size: 12px;
}
QLabel#statusLabel {
    color: #00ff88;
    font-size: 13px;
    font-weight: bold;
}
QLabel#titleLabel {
    color: #00f0ff;
    font-size: 22px;
    font-weight: bold;
}
QLabel#subtitleLabel {
    color: #8080c0;
    font-size: 14px;
}
QSplitter::handle {
    background: #00f0ff22;
    width: 2px;
}
QScrollBar:vertical {
    background: #0a0a1a;
    width: 8px;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #00f0ff33;
    border-radius: 4px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover {
    background: #00f0ff55;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
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
                    # 经纬度、XY坐标、地面高程、测点高程 均使用坐标文件数据
                    for k in ("经度", "纬度", "地面高程", "测点高程", "X", "Y"):
                        r.pop(k, None)
                    records.append(r)
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
        self.setWindowTitle("HYDRA EXTRACTOR // 地下水统测信息提取系统 v2.0")
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
        title = QLabel("HYDRA EXTRACTOR")
        title.setObjectName("titleLabel")
        subtitle = QLabel("地下水统测信息提取系统  //  AI-Powered OCR Engine")
        subtitle.setObjectName("subtitleLabel")
        title_block.addWidget(title)
        title_block.addWidget(subtitle)
        header.addLayout(title_block)
        header.addStretch()
        
        self.status_label = QLabel("就绪 // READY")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setStyleSheet("font-size: 13px; color: #00ff88;")
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
        engine_frame.setStyleSheet("QFrame { background: rgba(0,240,255,0.05); border: 1px solid #00f0ff33; border-radius: 10px; padding: 4px 8px; }")
        engine_layout = QHBoxLayout(engine_frame)
        engine_layout.setContentsMargins(8, 2, 8, 2)
        engine_layout.setSpacing(8)
        
        engine_label = QLabel("识别引擎")
        engine_label.setStyleSheet("color: #8080c0; font-size: 11px; background: transparent; border: none;")
        engine_layout.addWidget(engine_label)
        
        self.engine_group = QButtonGroup(self)
        
        self.rb_easyocr = QRadioButton("EasyOCR")
        self.rb_easyocr.setStyleSheet("""
            QRadioButton { color: #c0c0f0; font-size: 11px; font-weight: bold; background: transparent; spacing: 4px; }
            QRadioButton::indicator { width: 14px; height: 14px; border-radius: 8px; border: 2px solid #00f0ff55; background: rgba(0,240,255,0.05); }
            QRadioButton::indicator:checked { background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #00f0ff, stop:1 #00ff88); border: 2px solid #00f0ff; }
        """)
        self.rb_easyocr.setChecked(True)
        
        self.rb_mimo = QRadioButton("MiMo")
        self.rb_mimo.setStyleSheet("""
            QRadioButton { color: #c0c0f0; font-size: 11px; font-weight: bold; background: transparent; spacing: 4px; }
            QRadioButton::indicator { width: 14px; height: 14px; border-radius: 8px; border: 2px solid #ff00ff55; background: rgba(255,0,255,0.05); }
            QRadioButton::indicator:checked { background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 #ff00ff, stop:1 #ff8800); border: 2px solid #ff00ff; }
        """)
        
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
        left.setStyleSheet("QFrame { background: rgba(0,240,255,0.03); border-radius: 12px; border: 1px solid #00f0ff22; }")
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(16, 16, 16, 16)
        
        stats_group = QGroupBox("📊 统计面板")
        stats_layout = QVBoxLayout(stats_group)
        
        self.lbl_total = QLabel("待处理文件: 0")
        self.lbl_total.setStyleSheet("font-size: 28px; font-weight: bold; color: #00f0ff;")
        stats_layout.addWidget(self.lbl_total)
        
        self.lbl_extracted = QLabel("已提取记录: 0")
        self.lbl_extracted.setStyleSheet("font-size: 16px; color: #00ff88;")
        stats_layout.addWidget(self.lbl_extracted)
        
        self.lbl_coords = QLabel("含坐标记录: 0")
        self.lbl_coords.setStyleSheet("font-size: 16px; color: #ff00ff;")
        stats_layout.addWidget(self.lbl_coords)
        
        self.lbl_matched = QLabel("模板匹配: 0")
        self.lbl_matched.setStyleSheet("font-size: 16px; color: #ff8800;")
        stats_layout.addWidget(self.lbl_matched)
        
        stats_layout.addStretch()
        left_layout.addWidget(stats_group)
        
        # Recent files list
        files_group = QGroupBox("📂 文件列表")
        files_layout = QVBoxLayout(files_group)
        self.lbl_files = QLabel("拖拽或导入截图文件夹")
        self.lbl_files.setWordWrap(True)
        self.lbl_files.setStyleSheet("color: #6060a0; font-size: 13px;")
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
        """Apply neon glow to buttons."""
        for btn_name in ["btnExtract", "btnExport"]:
            btn = self.findChild(QPushButton, btn_name)
            if btn:
                shadow = QGraphicsDropShadowEffect()
                shadow.setBlurRadius(20)
                shadow.setColor(QColor("#00f0ff88" if "Extract" in btn_name else "#ff00ff88"))
                shadow.setOffset(0, 0)
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
        self.status_label.setStyleSheet("font-size: 13px; color: #ff8800;")
    
    def _import_coord(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择坐标文件", "D:/地下水", "Excel (*.xlsx *.xls)")
        if not path:
            return
        self.coord_path = path
        self.lbl_coords.setText("坐标文件: {}".format(os.path.basename(path)))
        self.lbl_coords.setStyleSheet("font-size: 16px; color: #ff8800;")
        self.status_label.setText("坐标已加载 // {}".format(os.path.basename(path)))
        self.status_label.setStyleSheet("font-size: 13px; color: #ff8800;")
    
    def _on_engine_changed(self, btn):
        """Handle engine radio button selection."""
        if btn == self.rb_easyocr:
            self.engine = "easyocr"
            self.status_label.setText("引擎: EasyOCR // 本地离线")
        else:
            self.engine = "mimo"
            self.status_label.setText("引擎: MiMo // 云端视觉")
        self.status_label.setStyleSheet("font-size: 13px; color: #00ff88;")
    
    def _extract(self):
        if not self.files:
            return
        self.btn_extract.setEnabled(False)
        self.btn_import.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("提取中... // PROCESSING")
        self.status_label.setStyleSheet("font-size: 13px; color: #00f0ff;")
        
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
        self.status_label.setStyleSheet("font-size: 13px; color: #00ff88;")
    
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
            if st == "已匹配": si.setForeground(QColor("#00ff88"))
            elif st == "新增": si.setForeground(QColor("#00f0ff"))
            elif st == "无编号": si.setForeground(QColor("#ff4488"))
            self.table.setItem(i, 0, si)
            # Row
            mr = str(r.get("_match_row", "") or "")
            self.table.setItem(i, 1, QTableWidgetItem(mr))
            # Source
            src = str(r.get("_source", "") or "")[:25]
            sri = QTableWidgetItem(src)
            sri.setForeground(QColor("#505080"))
            self.table.setItem(i, 2, sri)
            # Fields
            for j, (fn, _) in enumerate(FIELDS):
                val = str(r.get(fn, "") or "")
                item = QTableWidgetItem(val)
                if not val:
                    item.setForeground(QColor("#303050"))
                if fn == "野外编号" and val:
                    item.setForeground(QColor("#00f0ff"))
                if fn in ("经度", "纬度") and val:
                    item.setForeground(QColor("#ff00ff"))
                self.table.setItem(i, j + 3, item)
            if i % 10 == 0:
                QApplication.processEvents()  # keep UI responsive every 10 rows
        self.table.setUpdatesEnabled(True)  # single repaint for all changes
    
    def _export(self):
        if not self.records:
            QMessageBox.warning(self, "提示", "没有数据可导出")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出填充表",
            f"地下水统测表_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "Excel (*.xlsx)")
        if not path: return
        
        try:
            col_map = {f[0]: f[1] for f in FIELDS}
            wb = openpyxl.load_workbook(TEMPLATE_PATH)
            ws = wb["Sheet1"]
            tf = Font(name='仿宋', size=10)
            ta = Alignment(horizontal='center', vertical='center', wrap_text=True)
            tb = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
            
            updated = added = with_coords = 0
            seen_ids = set()
            
            for record in self.records:
                fid = str(record.get("野外编号", "")).strip()
                if not fid or fid in seen_ids or fid == 'None': continue
                seen_ids.add(fid)
                if record.get("_has_coords"): with_coords += 1
                
                target = record.get("_match_row")
                if target:
                    updated += 1
                else:
                    target = ws.max_row + 1
                    added += 1
                    ws.row_dimensions[target].height = 36
                
                for key, col in col_map.items():
                    val = record.get(key)
                    
                    # 水位标高 is always a formula
                    if key == "水位标高":
                        cell = ws[f"{col}{target}"]
                        cell.font = tf; cell.alignment = ta; cell.border = tb
                        cell.value = f"=L{target}-V{target}"
                        continue
                    
                    # 测点高程 = 地面高程
                    if key == "测点高程":
                        elev = record.get("地面高程")
                        if elev is not None and elev != '':
                            cell = ws[f"{col}{target}"]
                            cell.font = tf; cell.alignment = ta; cell.border = tb
                            try: cell.value = float(elev)
                            except: cell.value = str(elev)
                        continue
                    
                    if val is None or val == '': continue
                    
                    cell = ws[f"{col}{target}"]
                    cell.font = tf; cell.alignment = ta; cell.border = tb
                    if key in ("井台高度","井深","测点距地面高度","地下水位埋深","测点距水面距离","地面高程","X","Y"):
                        try: cell.value = float(val)
                        except: cell.value = str(val)
                    elif key == "调查日期":
                        cell.value = str(val).replace('-','')[:8]
                    else:
                        cell.value = str(val)
            
            wb.save(path)
            QMessageBox.information(self, "导出成功",
                f"已保存到:\n{path}\n\n更新 {updated} 行，新增 {added} 行\n其中 {with_coords} 条已自动填入坐标")
            self.status_label.setText(f"已导出 // EXPORTED: {updated}+{added} rows")
            self.status_label.setStyleSheet("font-size: 13px; color: #ff00ff;")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))
    
    def _clear(self):
        self.records = []
        self.files = []
        self.table.setRowCount(0)
        self.btn_export.setVisible(False)
        self.lbl_total.setText("待处理文件: 0")
        self.lbl_extracted.setText("已提取记录: 0")
        self.lbl_coords.setText("含坐标记录: 0")
        self.lbl_matched.setText("模板匹配: 0")
        self.lbl_files.setText("拖拽或导入截图文件夹")
        self.status_label.setText("就绪 // READY")
        self.status_label.setStyleSheet("font-size: 13px; color: #00ff88;")


# ── License Dialog ───────────────────────────────────────────────────
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
        
        title = QLabel("🔐 软件授权验证")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #00f0ff; font-size: 16px; font-weight: bold;")
        layout.addWidget(title)
        
        # Machine code row
        layout.addWidget(QLabel("机器码（发送给管理员获取密钥）："))
        
        mc_row = QHBoxLayout()
        self.mc_box = QLineEdit(self.machine_code)
        self.mc_box.setReadOnly(True)
        self.mc_box.setAlignment(Qt.AlignCenter)
        self.mc_box.setStyleSheet("""
            QLineEdit {
                background: #0d0d2b; color: #ffaa00; border: 2px solid #ffaa0044;
                border-radius: 8px; padding: 10px 14px; font-size: 15px;
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
        self._mc_feedback.setStyleSheet("color: #00ff88; font-size: 11px;")
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
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #00ff8844, stop:1 #00f0ff44);
                color: #fff; border: 1px solid #00f0ff88;
                border-radius: 10px; padding: 10px 28px;
                font-size: 15px; font-weight: bold;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #00ff8866, stop:1 #00f0ff66);
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
