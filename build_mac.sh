#!/bin/bash
# Script dong goi "Phan mem Bao gia - ITD Viet Nam" thanh ung dung macOS (.app).
# CHI chay duoc TREN MAY MAC THAT (co macOS) - khong the chay tu Windows/Linux.
#
# Cach dung: mo Terminal, cd vao thu muc BaoGiaApp nay, roi chay:
#   chmod +x build_mac.sh
#   ./build_mac.sh
#
# Sau khi chay xong, ung dung se nam trong: dist/Phan mem Bao gia - ITD Viet Nam.app

set -e
cd "$(dirname "$0")"

echo "==> Kiem tra Python3..."
if ! command -v python3 &> /dev/null; then
    echo "LOI: Khong tim thay python3. Hay cai Python3 truoc (vd: tu python.org hoac 'brew install python3')."
    exit 1
fi
python3 --version

echo "==> Tao moi trường ao (venv)..."
rm -rf venv
python3 -m venv venv
source venv/bin/activate

echo "==> Cai dat thu vien tu requirements.txt..."
pip install --upgrade pip
pip install -r requirements.txt

echo "==> Xoa ban build cu (neu co)..."
rm -rf build dist

echo "==> Dong goi bang py2app..."
python3 setup_mac.py py2app

echo ""
echo "==================================================================="
echo "XONG! Ung dung nam tai:"
echo "  $(pwd)/dist/Phan mem Bao gia - ITD Viet Nam.app"
echo ""
echo "Co the copy file .app nay vao thu muc Applications va chay binh thuong."
echo "==================================================================="
