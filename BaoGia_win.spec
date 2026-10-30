# -*- mode: python ; coding: utf-8 -*-
# Spec file PyInstaller danh cho Windows (.exe).
#
# Cach dung (tren Windows, trong thu muc BaoGiaApp):
#   python -m venv venv
#   venv\Scripts\activate
#   pip install -r requirements.txt
#   pyinstaller BaoGia_win.spec
#
# Ket qua: dist/Phan mem Bao gia - ITD Viet Nam/Phan mem Bao gia - ITD Viet Nam.exe
# (dang thu muc "onedir" - copy ca thu muc nay khi chuyen may khac, khong chi rieng file .exe)

block_cipher = None

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("ui.html", "."),
    ],
    # Khai bao ro cac module pywebview can tren Windows - day chinh la nguyen nhan loi tung gap
    # "No module named 'webview'" khi dong goi truoc day: PyInstaller khong tu do quet ra duoc het
    # cac import dong/gian tiep cua pywebview (tuy backend duoc chon luc chay: edge chromium/mshtml/cef).
    hiddenimports=[
        "webview",
        "webview.platforms.winforms",
        "webview.platforms.edgechromium",
        "webview.platforms.mshtml",
        "webview.platforms.cef",
        "clr_loader",
        "pythonnet",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Phan mem Bao gia - ITD Viet Nam",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name="Phan mem Bao gia - ITD Viet Nam",
)
