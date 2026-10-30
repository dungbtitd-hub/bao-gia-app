# -*- coding: utf-8 -*-
"""
Script dong goi "Phan mem Bao gia - ITD Viet Nam" thanh ung dung macOS (.app)
bang py2app - day la cach dong goi CHINH THUC ma pywebview khuyen dung tren
macOS (on dinh hon PyInstaller tren ne tang nay).

QUAN TRONG: Script nay CHI chay duoc tren may Mac thuc su (co he dieu hanh
macOS) - khong the "bien dich cheo" (cross-compile) tu Windows/Linux sang
macOS. Phai copy toan bo thu muc BaoGiaApp nay sang mot may Mac, roi chay
lenh dong goi ngay tren may Mac do.

Cach dung (chay tren macOS, trong thu muc chua file nay):
    python3 -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt
    python3 setup_mac.py py2app

Ket qua: ung dung "Phan mem Bao gia - ITD Viet Nam.app" nam trong thu muc dist/,
co the copy vao Applications va chay nhu mot app macOS binh thuong (khong can
cai Python tren may nguoi dung cuoi).

Neu muon build nhanh de tu kiem tra (khong toi uu, chay nhanh hon nhung mo
cham hon va phu thuoc vao Python/thu vien cai san tren may), dung them "-A":
    python3 setup_mac.py py2app -A
"""

import os
from setuptools import setup

APP = ["main.py"]

# Cac file du lieu can dong goi theo cung ung dung (giao dien HTML, file mau Excel vi du).
# ui.html nhung chua logo/anh deu da nhung duoi dang base64 (data URI) ngay trong file HTML
# nay nen khong can them anh rieng.
DATA_FILES = [
    "ui.html",
    "export_logic.py",
]
if os.path.exists("Vi_du_File_mau_cong_viec_chuyen_gia_co_me_con.xlsx"):
    DATA_FILES.append("Vi_du_File_mau_cong_viec_chuyen_gia_co_me_con.xlsx")

OPTIONS = {
    "argv_emulation": False,
    "strip": True,
    # Neu da co file icon .icns, bo dau # dong duoi va doi ten file cho dung:
    # "iconfile": "app_icon.icns",
    # Khai bao ro cac goi pywebview/PyObjC can tren macOS. Dung "packages" (chep nguyen ca goi)
    # thay vi "includes" (chi do theo import tinh) cho cac goi PyObjC - vi cac goi nay nap module
    # con kieu "dong" (qua objc.loadBundle...) nen py2app khong tu do quet ra het duoc, de thieu
    # va gay loi "Launch error" khi mo app (xem Contents/MacOS/<ten app> chay tu Terminal de thay
    # traceback that neu van loi).
    "includes": ["webview.platforms.cocoa"],
    "packages": [
        "webview", "objc", "Foundation", "AppKit", "WebKit", "Quartz", "CoreFoundation",
        "openpyxl", "docx", "PIL",
    ],
    # requirements.txt co cai "pyinstaller" (dung cho ban Windows .exe / phuong an BaoGia_mac.spec),
    # nhung ban than goi PyInstaller lai co cau truc thu muc "hooks" voi ten file kieu
    # "hook-PySide6.QtQuick.py" khong tuong thich voi cach py2app/modulegraph tu do quet TOAN BO
    # cac goi dang cai trong venv - gay crash "ImportError: No module named
    # PyInstaller.hooks.hook-PySide6.QtQuick" ngay ca khi app khong he dung toi PyInstaller luc chay.
    # Loai han PyInstaller (va goi hook di kem) ra khoi pham vi quet cua py2app de tranh loi nay.
    "excludes": ["PyInstaller", "pyinstaller_hooks_contrib"],
    "plist": {
        "CFBundleName": "Phan mem Bao gia - ITD Viet Nam",
        "CFBundleDisplayName": "Phan mem Bao gia - ITD Viet Nam",
        "CFBundleIdentifier": "vn.itd.baogia",
        "CFBundleShortVersionString": "1.0.0",
        "NSHighResolutionCapable": True,
    },
}

setup(
    app=APP,
    name="Phan mem Bao gia - ITD Viet Nam",
    data_files=DATA_FILES,
    options={"py2app": OPTIONS},
    setup_requires=["py2app"],
)
