# -*- mode: python ; coding: utf-8 -*-
# Spec file PyInstaller DANH RIENG CHO macOS - phuong an THAY THE neu khong muon dung py2app
# (setup_mac.py / build_mac.sh). Chi chay duoc tren may Mac thuc su.
#
# Cach dung (tren macOS, trong thu muc BaoGiaApp):
#   python3 -m venv venv && source venv/bin/activate
#   pip install -r requirements.txt
#   pyinstaller BaoGia_mac.spec
#
# Ket qua: dist/Phan mem Bao gia - ITD Viet Nam.app

block_cipher = None

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("ui.html", "."),
    ],
    # Khai bao ro cac module pywebview can tren macOS (backend "cocoa") - day chinh la nguyen
    # nhan loi "No module named 'webview'" tung gap khi dong goi PyInstaller tren Windows truoc
    # day: PyInstaller khong tu do quet ra duoc cac import dong/gian tiep cua pywebview.
    hiddenimports=[
        "webview",
        "webview.platforms.cocoa",
        "objc",
        "Foundation",
        "AppKit",
        "WebKit",
        "Quartz",
        "PyObjCTools",
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
    name="BaoGiaApp",
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
    name="BaoGiaApp",
)

app = BUNDLE(
    coll,
    name="Phan mem Bao gia - ITD Viet Nam.app",
    icon=None,  # neu co file .icns, dien duong dan vao day
    bundle_identifier="vn.itd.baogia",
    info_plist={
        "CFBundleName": "Phan mem Bao gia - ITD Viet Nam",
        "CFBundleShortVersionString": "1.0.0",
        "NSHighResolutionCapable": True,
    },
)
