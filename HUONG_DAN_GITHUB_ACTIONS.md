# Đóng gói tự động trên GitHub Actions (Windows + macOS)

## Vì sao làm theo cách này

Mình không có quyền tạo repository hay đẩy code lên GitHub thay bạn trong phiên làm việc này (chưa kết nối tài khoản GitHub nào ở đây), nên không thể tự mở sẵn link repo cho bạn được. Thay vào đó mình đã chuẩn bị sẵn file cấu hình `.github/workflows/build.yml` — bạn chỉ cần tạo 1 repository trống và tải (upload) toàn bộ thư mục `BaoGiaApp` này lên, GitHub sẽ **tự động build cả bản Windows (.exe) và macOS (.app) trên máy chủ của họ**, bạn chỉ việc tải file kết quả về. Cách này còn tốt hơn build tay trên máy Mac/Windows của bạn — vì máy chủ GitHub luôn là môi trường sạch, không dính các vấn đề xung đột thư viện như mình với bạn đã gặp phải (PyInstaller, PyObjCTools...).

## Bước 1: Tạo repository trên GitHub

1. Vào https://github.com/new (cần đăng nhập tài khoản GitHub; nếu chưa có thì tạo miễn phí tại https://github.com/signup).
2. Đặt tên repository, ví dụ `bao-gia-app`.
3. Chọn **Private** nếu không muốn người khác thấy mã nguồn (GitHub Actions vẫn chạy miễn phí trên repo private với một số giờ build miễn phí hàng tháng — đủ dùng cho việc build thỉnh thoảng như thế này).
4. **Không** tích chọn "Add a README file" (để trống cho dễ upload).
5. Bấm **Create repository**.

## Bước 2: Tải (upload) toàn bộ thư mục BaoGiaApp lên

Cách đơn giản nhất — không cần cài git:

1. Ở trang repository vừa tạo, tìm dòng chữ **"uploading an existing file"** (hiện ngay trên trang hướng dẫn mặc định) và bấm vào.
2. Kéo thả **toàn bộ nội dung bên trong** thư mục `BaoGiaApp` (bao gồm cả thư mục ẩn `.github` — xem lưu ý bên dưới) vào khung upload.
3. Bấm **Commit changes**.

**Lưu ý quan trọng về thư mục `.github`**: thư mục này bắt đầu bằng dấu chấm nên có thể bị hệ điều hành ẩn đi khi bạn mở Finder/File Explorer để kéo thả. Cách chắc chắn nhất: dùng đúng tính năng "uploading an existing file" của GitHub và kéo thả **cả thư mục `BaoGiaApp` gốc** (không chỉ các file bên trong) vào khung — GitHub sẽ tự giữ đúng cấu trúc thư mục con kể cả thư mục ẩn, miễn là bạn kéo thả bằng trình duyệt (không phải chọn từng file qua hộp thoại "Open File" vốn thường lọc bỏ file/thư mục ẩn).

Nếu bạn quen dùng git, cách dòng lệnh (đáng tin cậy hơn, không lo sót file ẩn):
```bash
cd BaoGiaApp
git init
git add -A
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<tên-tài-khoản>/bao-gia-app.git
git push -u origin main
```

## Bước 3: Xem kết quả build

1. Vào tab **Actions** trên trang repository (thanh menu ngang phía trên).
2. Sau khi push/upload xong, một lượt chạy tên **"Build BaoGiaApp (Windows + macOS)"** sẽ tự khởi động (mất khoảng 3-8 phút cho mỗi bản).
3. Bấm vào lượt chạy đó để xem tiến trình. Nếu có lỗi, phần log chi tiết sẽ hiện ra — gửi lại cho mình đoạn lỗi đó, mình sẽ sửa tiếp.
4. Khi chạy xong (dấu tích xanh ✓), kéo xuống mục **Artifacts** ở cuối trang của lượt chạy đó — sẽ thấy 2 file:
   - `BaoGiaApp-Windows` — giải nén ra sẽ có file `.exe` để chạy trên Windows.
   - `BaoGiaApp-macOS` — chứa file `BaoGiaApp-macOS.zip`, giải nén ra sẽ có file `.app` để chạy trên Mac (lưu ý: vẫn cần làm theo bước "bấm chuột phải → Open" lần đầu do chưa ký số, xem phần "Vài lưu ý" trong `HUONG_DAN_BUILD_MAC.md`).

## Build lại sau khi sửa code

Mỗi khi mình gửi bạn phiên bản `export_logic.py`/`ui.html`/... mới, bạn chỉ cần thay các file đó vào repository (vẫn qua "Add file → Upload files" trên GitHub, hoặc `git add -A && git commit -m "update" && git push` nếu dùng dòng lệnh) — GitHub Actions sẽ tự động build lại cả 2 bản, không cần làm gì thêm. Cũng có thể tự kích hoạt build thủ công bất cứ lúc nào (không cần sửa code) bằng nút **"Run workflow"** ở tab Actions.

## Giới hạn cần biết

- Tài khoản GitHub miễn phí có khoảng 2.000 phút build/tháng cho repo private (repo public thì không giới hạn) — mỗi lượt build 2 bản này tốn khoảng 10-15 phút tổng cộng, nên dư sức dùng thường xuyên.
- File kết quả (artifacts) GitHub giữ mặc định 90 ngày rồi tự xoá — nếu cần giữ lâu hơn, tải về máy sau mỗi lần build là chắc ăn nhất.
