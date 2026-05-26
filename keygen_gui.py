"""
Key Generator GUI - WPF/Cyberpunk style
"""
import hashlib, sys
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QMessageBox)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QPalette, QColor

SALT = "HYDRA-WELL-2026-X8K3M"

STYLE = """
QWidget {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #0a0a1a, stop:1 #0d0d2b);
    color: #e0e0ff;
    font-family: "Microsoft YaHei";
}
QLineEdit {
    background: #0d0d2b;
    color: #00ff88;
    border: 2px solid #00f0ff55;
    border-radius: 8px;
    padding: 10px 16px;
    font-size: 18px;
    font-weight: bold;
    letter-spacing: 2px;
}
QLineEdit:focus {
    border: 2px solid #00f0ff;
}
QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #ff00ff44, stop:1 #00f0ff44);
    color: #fff;
    border: 1px solid #ff00ff88;
    border-radius: 10px;
    padding: 12px 32px;
    font-size: 16px;
    font-weight: bold;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #ff00ff66, stop:1 #00f0ff66);
}
QPushButton#btnGen {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00ff8844, stop:1 #00f0ff44);
    color: #fff;
    font-size: 18px;
    padding: 14px 40px;
}
QPushButton#btnGen:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #00ff8866, stop:1 #00f0ff66);
}
QLabel {
    color: #c0c0f0;
}
QLabel#title {
    color: #00f0ff;
    font-size: 20px;
    font-weight: bold;
}
QLabel#result {
    color: #00ff88;
    font-size: 22px;
    font-weight: bold;
}
"""

class KeyGen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HYDRA KEYGEN // 授权密钥生成器")
        self.setFixedSize(520, 380)
        self._ui()
    
    def _ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(40, 30, 40, 30)
        
        title = QLabel("HYDRA KEYGEN")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        layout.addSpacing(10)
        
        label1 = QLabel("输入机器码：")
        label1.setAlignment(Qt.AlignCenter)
        layout.addWidget(label1)
        
        self.input_mc = QLineEdit()
        self.input_mc.setPlaceholderText("XXXX-XXXX-XXXX-XXXX")
        self.input_mc.setAlignment(Qt.AlignCenter)
        self.input_mc.textChanged.connect(self._on_input)
        layout.addWidget(self.input_mc)
        
        layout.addSpacing(10)
        
        self.btn_gen = QPushButton("⚡ 生成密钥")
        self.btn_gen.setObjectName("btnGen")
        self.btn_gen.clicked.connect(self._generate)
        layout.addWidget(self.btn_gen, alignment=Qt.AlignCenter)
        
        layout.addSpacing(10)
        
        # Result display — read-only QLineEdit so text is selectable
        self.result_box = QLineEdit()
        self.result_box.setReadOnly(True)
        self.result_box.setAlignment(Qt.AlignCenter)
        self.result_box.setPlaceholderText("密钥将显示在这里")
        self.result_box.setStyleSheet("""
            QLineEdit {
                background: #0d0d2b; color: #00ff88; border: 1px solid #00f0ff33;
                border-radius: 6px; padding: 8px 12px; font-size: 16px;
                font-weight: bold; letter-spacing: 2px;
            }
        """)
        layout.addWidget(self.result_box)
        
        # Copy button row
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        
        self.btn_copy = QPushButton("📋 复制密钥")
        self.btn_copy.setVisible(False)
        self.btn_copy.clicked.connect(self._copy_key)
        btn_row.addWidget(self.btn_copy)
        
        self.copy_feedback = QLabel("")
        self.copy_feedback.setStyleSheet("color: #00ff88; font-size: 12px;")
        btn_row.addWidget(self.copy_feedback)
        btn_row.addStretch()
        layout.addLayout(btn_row)
        
        layout.addStretch()
    
    def _on_input(self):
        self.result_box.setText("")
        self.btn_copy.setVisible(False)
        self.copy_feedback.setText("")
    
    def _generate(self):
        mc = self.input_mc.text().strip()
        if not mc:
            QMessageBox.warning(self, "提示", "请先输入机器码")
            return
        clean = mc.replace('-', '').strip().upper()
        if len(clean) < 8:
            QMessageBox.warning(self, "提示", "机器码格式不正确")
            return
        h = hashlib.sha256((clean + SALT).encode()).hexdigest()[:16].upper()
        key = '-'.join([h[i:i+4] for i in range(0, 16, 4)])
        self.result_box.setText(key)
        self.btn_copy.setVisible(True)
        self.copy_feedback.setText("")
        # Auto-copy to clipboard
        QApplication.clipboard().setText(key)
        self.copy_feedback.setText("✅ 已复制到剪贴板!")
    
    def _copy_key(self):
        key = self.result_box.text().strip()
        if key:
            QApplication.clipboard().setText(key)
            self.copy_feedback.setText("✅ 已复制到剪贴板!")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setStyleSheet(STYLE)
    font = QFont("Microsoft YaHei", 10)
    app.setFont(font)
    w = KeyGen()
    w.show()
    sys.exit(app.exec())
