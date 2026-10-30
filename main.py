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
)


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

    def save_quote_project(self, state):
        try:
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
                directory="",
                file_types=("Quote project files (*.json)", "All files (*.*)"),
            )
            if not result:
                return {"cancelled": True}
            open_path = result[0] if isinstance(result, (list, tuple)) else result
            with open(open_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {"ok": True, "data": data}
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

    def _ask_save_path(self, default_name, file_filter):
        result = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            directory="",
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
            directory="",
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
    webview.start()
