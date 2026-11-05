# -*- coding: utf-8 -*-
"""
Phần mềm Báo giá - ITD Việt Nam
Ứng dụng desktop offline, dùng pywebview để hiển thị giao diện HTML
và Python (openpyxl / python-docx) để xuất file Excel / Word.
"""

import os
import sys
import json
import webview
from export_logic import (
    export_excel_file,
    export_word_file,
    build_expert_template_workbook,
    parse_expert_import_excel,
    build_other_costs_template,
    parse_other_costs_excel,
    build_salary_template,
    parse_salary_excel,
    export_registry_report,
)


def app_data_dir():
    """Thu muc luu du lieu dung chung cua phan mem (cau hinh mac dinh + so quan ly bao gia)."""
    if sys.platform == "darwin":
        base = os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    else:
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    d = os.path.join(base, "BaoGiaITD")
    os.makedirs(d, exist_ok=True)
    return d


def _read_json(name, default):
    path = os.path.join(app_data_dir(), name)
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _write_json(name, data):
    path = os.path.join(app_data_dir(), name)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)



def resource_path(relative_path):
    """Lấy đường dẫn đúng khi chạy từ file .py, hoặc từ bản đã đóng gói:
    - Windows/PyInstaller (.exe): dữ liệu nằm trong sys._MEIPASS.
    - macOS/py2app (.app): dữ liệu nằm trong thư mục Resources của bundle,
      py2app set san bien moi truong RESOURCEPATH tro toi dung thu muc nay.
    """
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    elif os.environ.get("RESOURCEPATH"):
        base_path = os.environ["RESOURCEPATH"]
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)



def _send_to_trash(path):
    """Dua 1 file vao Thung rac cua he dieu hanh (Recycle Bin / Trash)."""
    import subprocess
    path = os.path.abspath(path)
    if sys.platform.startswith("win"):
        import ctypes
        from ctypes import wintypes

        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [("hwnd", wintypes.HWND), ("wFunc", wintypes.UINT), ("pFrom", wintypes.LPCWSTR),
                        ("pTo", wintypes.LPCWSTR), ("fFlags", ctypes.c_uint16), ("fAnyOperationsAborted", wintypes.BOOL),
                        ("hNameMappings", ctypes.c_void_p), ("lpszProgressTitle", wintypes.LPCWSTR)]
        FO_DELETE, FOF_SILENT, FOF_NOCONFIRMATION, FOF_ALLOWUNDO, FOF_NOERRORUI = 3, 0x4, 0x10, 0x40, 0x400
        op = SHFILEOPSTRUCTW(None, FO_DELETE, path + "\0", None,
                             FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT | FOF_NOERRORUI, False, None, None)
        rc = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
        if rc != 0 or op.fAnyOperationsAborted:
            raise OSError("Không chuyển được vào Thùng rác (mã lỗi %s)" % rc)
    elif sys.platform == "darwin":
        try:
            from Foundation import NSURL, NSFileManager
            ok, _url, err = NSFileManager.defaultManager().trashItemAtURL_resultingItemURL_error_(
                NSURL.fileURLWithPath_(path), None, None)
            if not ok:
                raise OSError(str(err))
        except ImportError:
            subprocess.run(["osascript", "-e", 'tell application "Finder" to delete POSIX file "%s"' % path.replace('"', '\\"')], check=True)
    else:
        subprocess.run(["gio", "trash", path], check=True)

class Api:
    """Các hàm Python được gọi từ JavaScript (window.pywebview.api.<tên_hàm>)."""

    def export_excel(self, state):
        try:
            default_name = f"BaoGia_{self._safe_name(state.get('quoteCode', 'BaoGia'))}.xlsx"
            save_path = self._ask_save_path(default_name, "Excel files (*.xlsx)")
            if not save_path:
                return {"cancelled": True}
            export_excel_file(state, save_path)
            return {"ok": True, "path": save_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def export_word(self, state):
        try:
            default_name = f"BaoGia_{self._safe_name(state.get('quoteCode', 'BaoGia'))}.docx"
            save_path = self._ask_save_path(default_name, "Word files (*.docx)")
            if not save_path:
                return {"cancelled": True}
            export_word_file(state, save_path)
            return {"ok": True, "path": save_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_quote_project(self, state, path=None):
        try:
            if path and os.path.isdir(os.path.dirname(path) or "."):
                save_path = path
            else:
                default_name = f"BaoGia_{self._safe_name(state.get('quoteCode', 'BaoGia'))}.json"
                save_path = self._ask_save_path(default_name, "Quote project files (*.json)")
            if not save_path:
                return {"cancelled": True}
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
            return {"ok": True, "path": save_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_quote_project(self):
        try:
            result = webview.windows[0].create_file_dialog(
                webview.OPEN_DIALOG,
                directory=self._default_dir(),
                file_types=("Quote project files (*.json)", "All files (*.*)"),
            )
            if not result:
                return {"cancelled": True}
            open_path = result[0] if isinstance(result, (list, tuple)) else result
            with open(open_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {"ok": True, "data": data, "path": open_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def pick_quote_files(self):
        """Chon NHIEU file bao gia (.json) de cap nhat vao so Quan ly bao gia - chi doc du lieu, khong mo len soan thao."""
        try:
            result = webview.windows[0].create_file_dialog(
                webview.OPEN_DIALOG,
                directory=self._default_dir(),
                allow_multiple=True,
                file_types=("Quote project files (*.json)", "All files (*.*)"),
            )
            if not result:
                return {"cancelled": True}
            paths = list(result) if isinstance(result, (list, tuple)) else [result]
            files = []
            for pth in paths:
                try:
                    with open(pth, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    if isinstance(data, dict):
                        data.pop("logoDataUri", None)  # khong can cho so quan ly, giam dung luong truyen
                    try:
                        mtime = os.path.getmtime(pth) * 1000
                    except Exception:
                        mtime = 0
                    files.append({"path": pth, "data": data, "mtime": mtime})
                except Exception as e:
                    files.append({"path": pth, "error": str(e)})
            return {"ok": True, "files": files}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def paths_exist(self, paths):
        """Kiem tra cac file trong so Quan ly bao gia con ton tai khong (de nhan biet file da bi di chuyen)."""
        out = {}
        for pth in paths or []:
            try:
                out[pth] = bool(pth) and os.path.isfile(pth)
            except Exception:
                out[pth] = False
        return out

    def move_to_trash(self, paths):
        """Chuyen cac file bao gia trung noi dung vao Thung rac (khoi phuc lai duoc), khong xoa vinh vien."""
        done, failed = [], []
        for pth in paths or []:
            try:
                if not pth or not os.path.isfile(pth):
                    raise FileNotFoundError("Không tìm thấy file")
                _send_to_trash(pth)
                done.append(pth)
            except Exception as e:
                failed.append({"path": pth, "error": str(e)})
        return {"ok": True, "done": done, "failed": failed}

    # ---- Mat khau dang nhap phan mem (luu dang bam PBKDF2 trong thu muc du lieu, file auth.json) ----

    @staticmethod
    def _hash(secret, salt):
        import hashlib
        return hashlib.pbkdf2_hmac("sha256", secret.encode("utf-8"), bytes.fromhex(salt), 200000).hex()

    @staticmethod
    def _norm_code(code):
        return "".join(ch for ch in str(code or "").upper() if ch.isalnum())

    def _new_auth(self, password):
        import secrets
        alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
        raw = "".join(secrets.choice(alphabet) for _ in range(12))
        code = raw[0:4] + "-" + raw[4:8] + "-" + raw[8:12]
        salt = secrets.token_hex(16)
        rsalt = secrets.token_hex(16)
        _write_json("auth.json", {"v": 1, "salt": salt, "hash": self._hash(password, salt),
                                  "rsalt": rsalt, "rhash": self._hash(self._norm_code(code), rsalt)})
        return code

    _authed = False

    def auth_status(self):
        a = _read_json("auth.json", {})
        return {"hasPassword": bool(a.get("hash")), "authed": bool(Api._authed and a.get("hash"))}

    def auth_setup(self, password):
        try:
            if _read_json("auth.json", {}).get("hash"):
                return {"ok": False, "error": "Mật khẩu đã được đặt trước đó."}
            if not password or len(password) < 4:
                return {"ok": False, "error": "Mật khẩu cần ít nhất 4 ký tự."}
            code = self._new_auth(password)
            Api._authed = True
            return {"ok": True, "recovery": code}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def auth_check(self, password):
        try:
            import hmac
            a = _read_json("auth.json", {})
            if not a.get("hash"):
                return False
            ok = hmac.compare_digest(a["hash"], self._hash(password or "", a["salt"]))
            if ok:
                Api._authed = True
            return ok
        except Exception:
            return False

    def auth_recover(self, code, new_password):
        try:
            import hmac
            a = _read_json("auth.json", {})
            if not a.get("rhash"):
                return {"ok": False, "error": "Chưa có mã khôi phục."}
            if not hmac.compare_digest(a["rhash"], self._hash(self._norm_code(code), a["rsalt"])):
                return {"ok": False, "error": "Mã khôi phục không đúng."}
            if not new_password or len(new_password) < 4:
                return {"ok": False, "error": "Mật khẩu cần ít nhất 4 ký tự."}
            return {"ok": True, "recovery": self._new_auth(new_password)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def auth_change(self, old_password, new_password):
        if not self.auth_check(old_password):
            return {"ok": False, "error": "Mật khẩu hiện tại không đúng."}
        if not new_password or len(new_password) < 4:
            return {"ok": False, "error": "Mật khẩu cần ít nhất 4 ký tự."}
        try:
            return {"ok": True, "recovery": self._new_auth(new_password)}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- Dong cua so: hoi luu bao gia neu co thay doi (xem on_closing o duoi) ----
    _allow_close = False

    def close_app(self):
        import threading
        Api._allow_close = True
        def _go():
            try:
                webview.windows[0].destroy()
            except Exception:
                pass
        threading.Thread(target=_go, daemon=True).start()
        return True

    def reveal_in_folder(self, path):
        """Mo thu muc chua file bao gia (va chon san file do) trong Explorer/Finder."""
        try:
            if not path:
                return {"ok": False, "error": "Báo giá này chưa có đường dẫn file."}
            folder = path if os.path.isdir(path) else os.path.dirname(path)
            if not os.path.isdir(folder):
                return {"ok": False, "error": "Không tìm thấy thư mục (có thể đã bị di chuyển hoặc xóa): " + str(folder)}
            import subprocess
            if sys.platform.startswith("win"):
                if os.path.isfile(path):
                    subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
                else:
                    os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", path] if os.path.isfile(path) else ["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def open_quote_path(self, path):
        """Mo lai 1 file bao gia theo duong dan da luu trong so Quan ly bao gia."""
        try:
            if not path or not os.path.exists(path):
                return {"ok": False, "error": "Không tìm thấy file (có thể đã bị di chuyển hoặc xóa): " + str(path)}
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {"ok": True, "data": data, "path": path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- Cau hinh mac dinh cho bao gia moi + so quan ly bao gia ----

    def load_app_data(self):
        try:
            return {"ok": True, "defaults": _read_json("defaults.json", {}), "registry": _read_json("registry.json", [])}
        except Exception as e:
            return {"ok": False, "error": str(e), "defaults": {}, "registry": []}

    def save_defaults(self, partial):
        try:
            cur = _read_json("defaults.json", {})
            cur.update(partial or {})
            _write_json("defaults.json", cur)
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_registry(self, records):
        try:
            _write_json("registry.json", records or [])
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def export_registry_excel(self, payload):
        try:
            save_path = self._ask_save_path(f"Bao_cao_bao_gia_{self._safe_name(payload.get('slug', ''))}.xlsx", "Excel files (*.xlsx)")
            if not save_path:
                return {"cancelled": True}
            export_registry_report(payload, save_path)
            return {"ok": True, "path": save_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- Import / Export file mẫu Excel: bảng phân tích chuyên gia ----

    def export_expert_template(self, salary_table):
        try:
            save_path = self._ask_save_path("Mau_cong_viec_chuyen_gia.xlsx", "Excel files (*.xlsx)")
            if not save_path:
                return {"cancelled": True}
            wb = build_expert_template_workbook(salary_table or {})
            wb.save(save_path)
            return {"ok": True, "path": save_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def import_expert_excel(self, salary_table):
        try:
            open_path = self._ask_open_path("Excel files (*.xlsx)")
            if not open_path:
                return {"cancelled": True}
            rows, warnings = parse_expert_import_excel(open_path, salary_table or {})
            return {"ok": True, "rows": rows, "warnings": warnings, "path": open_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- Import / Export file mẫu Excel: bảng chi phí khác ----

    def export_other_costs_template(self):
        try:
            save_path = self._ask_save_path("Mau_chi_phi_khac.xlsx", "Excel files (*.xlsx)")
            if not save_path:
                return {"cancelled": True}
            wb = build_other_costs_template()
            wb.save(save_path)
            return {"ok": True, "path": save_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def import_other_costs_excel(self):
        try:
            open_path = self._ask_open_path("Excel files (*.xlsx)")
            if not open_path:
                return {"cancelled": True}
            rows, warnings = parse_other_costs_excel(open_path)
            return {"ok": True, "rows": rows, "warnings": warnings, "path": open_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ---- Import / Export file mẫu Excel: bảng đơn giá tiền lương chuyên gia ----

    def export_salary_template(self, salary_table):
        try:
            save_path = self._ask_save_path("Mau_don_gia_tien_luong_chuyen_gia.xlsx", "Excel files (*.xlsx)")
            if not save_path:
                return {"cancelled": True}
            wb = build_salary_template(salary_table or {})
            wb.save(save_path)
            return {"ok": True, "path": save_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def import_salary_excel(self):
        try:
            open_path = self._ask_open_path("Excel files (*.xlsx)")
            if not open_path:
                return {"cancelled": True}
            data, warnings = parse_salary_excel(open_path)
            return {"ok": True, "data": data, "warnings": warnings, "path": open_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    @staticmethod
    def _safe_name(name):
        return "".join(c if (c.isalnum() or c in "._-") else "_" for c in name)

    def _default_dir(self):
        """Thu muc mac dinh cho hop chon noi luu/mo file (cai trong Tham so & thiet lap)."""
        try:
            d = (_read_json("defaults.json", {}) or {}).get("saveDir") or ""
            return d if d and os.path.isdir(d) else ""
        except Exception:
            return ""

    def pick_folder(self):
        """Chon thu muc mac dinh de luu bao gia."""
        try:
            result = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG, directory=self._default_dir())
            if not result:
                return {"cancelled": True}
            folder = result[0] if isinstance(result, (list, tuple)) else result
            return {"ok": True, "path": folder}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def folder_exists(self, path):
        try:
            return bool(path) and os.path.isdir(path)
        except Exception:
            return False

    # ---- Dong bo du lieu giua cac may (so Quan ly bao gia + tham so/cau hinh mac dinh) ----

    def export_sync(self, payload, default_name):
        try:
            save_path = self._ask_save_path(default_name or "DongBo_BaoGiaITD.json", "Sync files (*.json)")
            if not save_path:
                return {"cancelled": True}
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            return {"ok": True, "path": save_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def pick_sync_file(self):
        try:
            open_path = self._ask_open_path("Sync files (*.json)")
            if not open_path:
                return {"cancelled": True}
            with open(open_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {"ok": True, "data": data, "path": open_path}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def _ask_save_path(self, default_name, file_filter):
        result = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            directory=self._default_dir(),
            save_filename=default_name,
            file_types=(file_filter, "All files (*.*)"),
        )
        if not result:
            return None
        # create_file_dialog trả về str hoặc tuple tùy phiên bản pywebview
        if isinstance(result, (list, tuple)):
            return result[0]
        return result

    def _ask_open_path(self, file_filter):
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            directory=self._default_dir(),
            file_types=(file_filter, "All files (*.*)"),
        )
        if not result:
            return None
        if isinstance(result, (list, tuple)):
            return result[0]
        return result


if __name__ == "__main__":
    api = Api()
    window = webview.create_window(
        "Phần mềm Báo giá - ITD Việt Nam",
        resource_path("ui.html"),
        js_api=api,
        width=1400,
        height=900,
        min_size=(1000, 700),
    )

    def on_closing(*_args):
        """Bam X dong cua so: neu bao gia co thay doi chua luu thi de giao dien hoi (Khong luu / Luu de / Luu moi)."""
        import threading
        if Api._allow_close:
            return True

        def ask():
            try:
                r = window.evaluate_js("(typeof appCloseRequested === 'function') ? (setTimeout(function(){ appCloseRequested(); }, 0), 'ok') : 'none'")
            except Exception:
                r = None
            if r != "ok":
                Api._allow_close = True
                try:
                    window.destroy()
                except Exception:
                    pass
        threading.Thread(target=ask, daemon=True).start()
        return False

    window.events.closing += on_closing
    webview.start()
