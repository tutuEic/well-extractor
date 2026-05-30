# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all

# Collect everything from rapidocr_onnxruntime (Python + data + models)
rapidocr_datas, rapidocr_binaries, rapidocr_hiddenimports = collect_all('rapidocr_onnxruntime')

a = Analysis(
    ['desktop_app_v2.py'],
    pathex=[],
    binaries=rapidocr_binaries,
    datas=rapidocr_datas,
    hiddenimports=rapidocr_hiddenimports + [
        'onnxruntime',
        'cv2',
        'numpy',
        'PIL',
        'openpyxl',
        'pyclipper',
        'shapely',
        'pywt',
        'openai',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'easyocr', 'torch', 'torchvision', 'torchaudio',
        'matplotlib', 'IPython', 'jedi', 'parso', 'pygments',
        'tensorboard', 'sympy', 'pandas', 'scipy', 'skimage',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='\u5730\u4e0b\u6c34\u7edf\u6d4b\u4fe1\u606f\u63d0\u53d6\u5668',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)