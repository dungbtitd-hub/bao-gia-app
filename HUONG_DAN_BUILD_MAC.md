# Hướng dẫn tạo bản chạy trên macOS

## Điều quan trọng cần biết trước

Mình đang chạy trên môi trường Linux (cloud), nên **không thể tự biên dịch ra file .app của macOS tại đây** — đây là hạn chế chung của PyInstaller/py2app: phải build ngay trên một máy Mac thật (không "cross-compile" được từ Windows/Linux sang macOS, giống hệt việc trước đây bạn phải tự chạy PyInstaller trên máy Windows để ra file .exe).

Vì vậy mình đã chuẩn bị sẵn toàn bộ mã nguồn + script đóng gói — bạn chỉ cần copy cả thư mục `BaoGiaApp` này sang một máy Mac, rồi chạy 1 lệnh duy nhất là ra app.

Mã nguồn (`main.py`, `export_logic.py`, `ui.html`) đã là code thuần Python/HTML, không có gì đặc thù Windows, nên **chạy được trên macOS ngay cả khi không đóng gói** (xem Cách 3 ở dưới nếu chỉ cần thử nhanh).

## Cách 1 (khuyên dùng): đóng gói bằng py2app

Đây là cách mà chính thư viện pywebview khuyên dùng cho macOS (ổn định hơn PyInstaller trên nền tảng này).

Trên máy Mac, mở Terminal, vào thư mục `BaoGiaApp`, rồi chạy:

```bash
chmod +x build_mac.sh
./build_mac.sh
```

Script sẽ tự tạo virtual environment, cài thư viện, và đóng gói. Xong sẽ ra:

```
dist/Phan mem Bao gia - ITD Viet Nam.app
```

Có thể copy file `.app` này vào thư mục Applications và dùng như một app macOS bình thường (máy người dùng cuối không cần cài Python).

Nếu muốn tự làm từng bước (không dùng script):

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 setup_mac.py py2app
```

## Cách 2 (thay thế): đóng gói bằng PyInstaller

Nếu bạn muốn dùng PyInstaller cho quen (giống lúc đóng gói .exe trên Windows), mình cũng đã chuẩn bị file `BaoGia_mac.spec` riêng cho macOS (đã khai báo đầy đủ các "hidden imports" của pywebview trên macOS như `webview.platforms.cocoa`, `objc`, `WebKit`, `Foundation`... để tránh đúng lỗi "No module named 'webview'" đã từng gặp trên Windows):

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
pyinstaller BaoGia_mac.spec
```

Kết quả cũng nằm ở `dist/Phan mem Bao gia - ITD Viet Nam.app`.

## Cách 3: chạy thử nhanh, không đóng gói

Nếu chỉ cần kiểm tra app chạy đúng trên Mac trước khi đóng gói:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 main.py
```

## Vài lưu ý khi chạy file .app trên Mac

- **Cảnh báo "không thể xác minh nhà phát triển" (Gatekeeper)**: vì app chưa được Apple ký số (code sign) / notarize, lần đầu mở có thể bị macOS chặn. Cách mở: bấm chuột phải (hoặc giữ Control) vào file `.app` → chọn "Open" → xác nhận "Open" lần nữa. Lần sau mở lại bình thường bằng double-click.
- **Chip Intel hay Apple Silicon (M1/M2/M3...)**: file `.app` build ra sẽ khớp với kiến trúc CPU của máy Mac dùng để build. Nếu chỉ dùng trên đúng máy đó thì không vấn đề gì; nếu cần dùng trên cả 2 loại máy khác nhau, nói mình biết để chuẩn bị thêm hướng dẫn build kiểu "universal2".
- **File xuất Excel/Word/PDF**: không có gì thay đổi so với bản Windows — `openpyxl`/`python-docx` hoạt động giống nhau trên mọi hệ điều hành.

## Xử lý sự cố: mở app báo "Launch error / See the py2app website for debugging launch issues"

Đây là thông báo CHUNG CHUNG của py2app khi app gặp lỗi Python lúc khởi động (thường do thiếu module) — bản thân thông báo này không nói rõ lỗi gì. Cách xem lỗi THẬT:

1. Mở Terminal, chạy thẳng file thực thi bên trong app (thay đúng đường dẫn tới `dist/` của bạn):
   ```bash
   "dist/Phan mem Bao gia - ITD Viet Nam.app/Contents/MacOS/Phan mem Bao gia - ITD Viet Nam"
   ```
2. Terminal sẽ in ra traceback Python đầy đủ — gửi lại đoạn đó cho mình để xác định chính xác đang thiếu gì.

**Lỗi 1 đã xác định và sửa**: `ImportError: No module named PyInstaller.hooks.hook-PySide6.QtQuick`. Nguyên nhân: `requirements.txt` có cài `pyinstaller` (dùng cho bản Windows .exe / phương án build thay thế bằng `BaoGia_mac.spec`), nhưng bản thân cấu trúc thư mục "hooks" của gói PyInstaller lại không tương thích với cách py2app tự động quét toàn bộ thư viện cài trong venv — gây crash dù app không hề dùng tới PyInstaller lúc chạy. Đã sửa bằng cách loại trừ hẳn gói này khỏi phạm vi quét của py2app trong `setup_mac.py`.

**Lỗi 2 đã xác định và sửa**: sau khi sửa lỗi 1, build tiến xa hơn nhiều (chạy hết toàn bộ các "recipe" và bắt đầu đóng gói) nhưng dừng ở bước cuối với `ImportError: No module named 'PyObjCTools'`. Nguyên nhân: mình liệt kê thừa `PyObjCTools` vào danh sách "packages" bắt buộc trong khi gói này không tồn tại độc lập trong môi trường (không ảnh hưởng gì đến việc app có chạy đúng hay không) — đã bỏ tên này ra khỏi `setup_mac.py`.

Bạn build lại (không cần tạo lại venv từ đầu, chỉ cần xoá bản build cũ):

```bash
source venv/bin/activate
rm -rf build dist
python3 setup_mac.py py2app
```

Nếu vẫn còn lỗi, gửi mình đoạn traceback lấy được ở bước 1 để mình sửa đúng chỗ.

## Danh sách file liên quan đã chuẩn bị

- `requirements.txt` — đã bổ sung các gói PyObjC + `py2app` cần riêng cho macOS (chỉ cài khi chạy trên macOS, không ảnh hưởng bản Windows).
- `setup_mac.py` — script đóng gói py2app.
- `build_mac.sh` — script tự động hoá Cách 1 (chạy 1 lệnh).
- `BaoGia_mac.spec` — spec file PyInstaller cho Cách 2.
- `main.py` — đã cập nhật hàm `resource_path()` để nhận diện đúng đường dẫn tài nguyên khi chạy từ bản đóng gói py2app (biến môi trường `RESOURCEPATH`), bên cạnh cách nhận diện PyInstaller cũ — không ảnh hưởng gì đến bản Windows hiện tại.
