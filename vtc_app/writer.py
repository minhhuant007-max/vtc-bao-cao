import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from copy import copy
import os
import shutil
import tempfile
import re
from datetime import datetime, date, timedelta
from typing import Dict, Any, List, Optional

from .models import TrackedCustomer

def normalize_dealer_name(s: str) -> str:
    """Chuẩn hóa tên đại lý để so khớp chính xác: loại bỏ khoảng trắng kép, chuẩn hóa dấu gạch ngang."""
    if not s:
        return ""
    s = re.sub(r'\s+', ' ', str(s)).strip().upper()
    s = s.replace(' - ', '-').replace('- ', '-').replace(' -', '-')
    return s

def match_column(sp: str, col_map: Dict[str, int]) -> Optional[int]:
    """
    Tìm vị trí cột phù hợp nhất trong col_map:
    1. So khớp chính xác không phân biệt hoa thường.
    2. So khớp sau khi loại bỏ khoảng trắng và xuống dòng thừa.
    3. So khớp theo chuỗi dài nhất chứa nhau.
    """
    sp_norm = sp.strip().lower()
    for sp_name, col_idx in col_map.items():
        if sp_norm == sp_name.strip().lower():
            return col_idx
            
    sp_clean = re.sub(r'\s+', '', sp_norm)
    for sp_name, col_idx in col_map.items():
        col_clean = re.sub(r'\s+', '', sp_name.strip().lower())
        if sp_clean == col_clean:
            return col_idx
            
    sorted_cols = sorted(col_map.items(), key=lambda x: len(x[0]), reverse=True)
    for sp_name, col_idx in sorted_cols:
        col_norm = sp_name.strip().lower()
        if sp_norm == col_norm or sp_norm in col_norm or col_norm in sp_norm:
            return col_idx
            
    return None

class ExcelWriter:
    @staticmethod
    def create_daily_backup(file_path: str, backup_dir: str, current_date: Optional[date] = None) -> str:
        """Tạo gói sao lưu mang ngày hôm trước nếu hôm nay chưa có sao lưu."""
        os.makedirs(backup_dir, exist_ok=True)
        if not current_date:
            current_date = date.today()
        prev_date = current_date - timedelta(days=1)
        prev_date_str = prev_date.strftime("%Y-%m-%d")
        
        backup_folder = os.path.join(backup_dir, f"backup_{prev_date_str}")
        os.makedirs(backup_folder, exist_ok=True)
        
        fname = os.path.basename(file_path)
        dest_path = os.path.join(backup_folder, fname)
        
        if not os.path.exists(dest_path) and os.path.exists(file_path):
            shutil.copy2(file_path, dest_path)
            
        return dest_path

    @staticmethod
    def copy_row_format(ws, src_row: int, dst_row: int, max_col: int = 26):
        for c in range(1, max_col + 1):
            sc = ws.cell(src_row, c)
            dc = ws.cell(dst_row, c)
            if sc.has_style:
                dc.font = copy(sc.font)
                dc.border = copy(sc.border)
                dc.fill = copy(sc.fill)
                dc.number_format = sc.number_format
                dc.alignment = copy(sc.alignment)

    @staticmethod
    def insert_dealer_row(ws, row: int):
        """Chèn 1 dòng an toàn và cập nhật vùng merged_cells."""
        moved = [copy(area) for area in ws.merged_cells.ranges if area.min_row >= row]
        for area in moved:
            ws.unmerge_cells(str(area))
        ws.insert_rows(row)
        for area in moved:
            area.shift(row_shift=1)
            ws.merge_cells(str(area))

    @staticmethod
    def find_day_column(ws, row: int, start: int, day: date) -> Optional[int]:
        """Tìm cột của ngày cụ thể trong dòng header."""
        short = day.strftime("%d/%m")
        for col in range(start, ws.max_column + 1):
            value = ws.cell(row, col).value
            if isinstance(value, datetime):
                if value.date() == day:
                    return col
            elif isinstance(value, date):
                if value == day:
                    return col
            elif value and re.match(rf"^{short}(?:/({day.year}))?(?:\s|$)", str(value).strip()):
                return col
            elif value and short in str(value).strip():
                return col
        return None

    @staticmethod
    def get_target_sheet(wb, report_type: str, target_date: Optional[date] = None):
        """
        Xác định sheet đích:
        - Nếu có target_date và sheet theo ngày (dd.mm.yyyy) tồn tại, trả về sheet đó.
        - Nếu có target_date và sheet theo tháng tồn tại, trả về sheet đó.
        - Mặc định tìm '09.11.2026' hoặc wb.active.
        """
        if target_date:
            sheet_day_str = target_date.strftime("%d.%m.%Y")
            if sheet_day_str in wb.sheetnames:
                return wb[sheet_day_str]
            month_code = target_date.strftime("T%m.%Y")
            for s in wb.sheetnames:
                if month_code in s:
                    return wb[s]
            month_alt = target_date.strftime("thang %m.%Y")
            for s in wb.sheetnames:
                if month_alt in s or month_alt.replace(' ', '') in s.replace(' ', ''):
                    return wb[s]
        for s in wb.sheetnames:
            if "09.11.2026" in s or "BC NGAY" in s.upper():
                return wb[s]
        return wb.active

    @staticmethod
    def validate_template(path: str, report_type: str, target_date: Optional[date] = None, dealers: Any = ()) -> str:
        """
        Kiểm tra tính hợp lệ của tệp mẫu Excel tổng:
        - Chặn chọn nhầm tệp đơn hàng nguồn (Đồng Nai / GRANITE / Bảng kê).
        - Kiểm tra cấu trúc cột CỬA HÀNG ở Mục 01 và Mục 02.
        - Hỗ trợ cả tệp Excel tổng hợp nhất 3 khối và mẫu riêng rẽ từng nhóm.
        - Trả về tên sheet đích hợp lệ.
        """
        if not os.path.exists(path):
            raise FileNotFoundError(f"Không tìm thấy tệp: {path}")

        base_name = os.path.basename(path).upper()
        if any(k in base_name for k in ["BẢNG KÊ", "BANG KE", "BẢNG KÊ ĐƠN HÀNG"]):
            raise ValueError(f"Tệp '{os.path.basename(path)}' là tệp bảng kê đơn hàng nguồn, không phải tệp mẫu Excel tổng. Vui lòng chọn lại tệp mẫu.")

        with open(path, "rb") as source:
            wb = openpyxl.load_workbook(source, read_only=False)

        try:
            ws = ExcelWriter.get_target_sheet(wb, report_type, target_date)

            first_rows_text = " ".join(
                str(ws.cell(r, c).value or "").strip().upper()
                for r in range(1, min(ws.max_row + 1, 6))
                for c in range(1, min(ws.max_column + 1, 10))
            )
            if "BẢNG KÊ ĐƠN HÀNG" in first_rows_text or ("NGÀY CT" in first_rows_text and "SỐ CT" in first_rows_text):
                raise ValueError(f"Tệp '{os.path.basename(path)}' là tệp bảng kê đơn hàng nguồn, không phải tệp mẫu Excel tổng. Vui lòng chọn lại tệp mẫu.")

            markers = {}
            for r in range(1, min(ws.max_row + 1, 120)):
                vals = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, 6)]
                combined = " | ".join(vals)
                if "TỔNG NHÓM TRM + TPS" in combined or "TỔNG NHÓM TRM+TPS" in combined:
                    markers["TRM_TPS_END"] = r
                elif "TỔNG NHÓM TCR + TML" in combined or "TỔNG NHÓM TCR+TML" in combined:
                    markers["TCR_TML_END"] = r
                elif "TỔNG NHÓM TMR" in combined:
                    markers["TMR_END"] = r

            is_unified = bool(markers.get("TRM_TPS_END") and markers.get("TCR_TML_END"))

            if is_unified:
                has_shop_col = (
                    str(ws.cell(2, 4).value or "").strip().upper() == "CỬA HÀNG"
                    and str(ws.cell(27, 4).value or "").strip().upper() == "CỬA HÀNG"
                )
                if not has_shop_col:
                    raise ValueError(f"Tệp Excel '{os.path.basename(path)}' không đúng cấu trúc mẫu báo cáo tiêu thụ (thiếu cột 'CỬA HÀNG' ở Mục 01 hoặc Mục 02).")
            else:
                headers = {
                    "TCR_TML": (2, 4, 27, 8, "TCR + TML", {"TCR", "TML"}),
                    "TRM_TPS": (3, 3, 24, 7, "TRM + TPS", {"TRM", "TPS"}),
                    "TMR": (3, 3, 9, 7, "TMR", {"TMR"}),
                }
                header, shop_col, month_header, first_day, label, brands = headers[report_type]
                if report_type == "TMR" and ws.cell(3, 1).value != "Mã khách hàng":
                    shop_col, first_day = 2, 5

                text = lambda r, c: str(ws.cell(r, c).value or "").strip().upper()
                title = " ".join(text(r, c) for r in (1, 2) for c in range(1, 26))
                foreign = {
                    "TCR_TML": ("TOROMA", "TOPISA", "TOMIRA"),
                    "TRM_TPS": ("TOCERA", "TOMILAN", "TOMIRA"),
                    "TMR": ("TOCERA", "TOMILAN", "TOROMA", "TOPISA"),
                }
                valid_structure = (
                    text(header, shop_col) == "CỬA HÀNG"
                    and text(month_header, shop_col) == "CỬA HÀNG"
                    and not any(name in title for name in foreign[report_type])
                )
                if not valid_structure:
                    raise ValueError(f"Tệp Excel không khớp nhóm {label}. Vui lòng chọn đúng tệp mẫu của nhóm này.")

            return ws.title
        finally:
            wb.close()

    @staticmethod
    def apply_confirmed_sync(ws, report_type: str, items: List[Dict[str, Any]], is_unified: bool = False):
        """Áp dụng các thay đổi đại lý đã xác nhận trực tiếp vào worksheet trước khi lưu."""
        if not items:
            return
        if is_unified or report_type == "TCR_TML":
            group_col, brand_col = 5, 6
        else:
            group_col, brand_col = 4, 5
        for item in items:
            row = item.get("row_num")
            if not row:
                continue
            for key, fallback, col in [
                ("final_code", "app_code", 1),
                ("final_group", "app_group", group_col),
                ("th", "th", brand_col),
            ]:
                val = item.get(key, item.get(fallback))
                if val:
                    ws.cell(row, col).value = val

    @staticmethod
    def apply_confirmed_sync_to_sheet(
        template_path: str,
        report_type: str,
        confirmed_items: List[Dict[str, Any]],
        target_date: Optional[date] = None
    ):
        """
        Áp dụng các mã khách hàng và phân nhóm đã được người dùng xác nhận vào Mục 02 của Excel tổng.
        Bảo toàn 100% công thức, định dạng và số liệu các ngày cũ.
        """
        if not os.path.exists(template_path) or not confirmed_items:
            return

        wb = openpyxl.load_workbook(template_path, data_only=False)
        ws = ExcelWriter.get_target_sheet(wb, report_type, target_date)
        
        # Check if unified
        markers = {}
        for r in range(1, min(ws.max_row + 1, 120)):
            vals = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, 6)]
            comb = " | ".join(vals)
            if "TỔNG NHÓM TRM + TPS" in comb or "TỔNG NHÓM TRM+TPS" in comb:
                markers["TRM_TPS_END"] = r
            elif "TỔNG NHÓM TCR + TML" in comb or "TỔNG NHÓM TCR+TML" in comb:
                markers["TCR_TML_END"] = r
        is_unified = bool(markers.get("TRM_TPS_END") and markers.get("TCR_TML_END"))

        ExcelWriter.apply_confirmed_sync(ws, report_type, confirmed_items, is_unified=is_unified)

        file_dir = os.path.dirname(os.path.abspath(template_path))
        fd, temp_file = tempfile.mkstemp(prefix=".vtc-sync-", suffix=".xlsx", dir=file_dir)
        os.close(fd)
        wb.save(temp_file)
        wb.close()

        try:
            os.replace(temp_file, template_path)
        except PermissionError:
            if os.path.exists(temp_file):
                os.unlink(temp_file)
            raise PermissionError(f"Tệp '{template_path}' đang được mở bởi ứng dụng khác (ví dụ Excel). Vui lòng đóng tệp rồi thử lại!")

    @staticmethod
    def analyze_master_excel_dealers(
        template_path: str,
        report_type: str,
        rule_manager: Any,
        target_date: Optional[date] = None
    ) -> List[Dict[str, Any]]:
        """
        Phân tích đối chiếu Mục 02 của tệp Excel tổng với danh sách đại lý đã xác minh trong AppData.
        Hỗ trợ cả tệp Excel tổng hợp nhất 3 khối và các tệp cũ riêng rẽ.
        """
        if not os.path.exists(template_path):
            return []

        wb = openpyxl.load_workbook(template_path, data_only=True)
        ws = ExcelWriter.get_target_sheet(wb, report_type, target_date)
        items = []

        markers = {}
        for r in range(1, min(ws.max_row + 1, 120)):
            vals = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, 6)]
            combined = " | ".join(vals)
            if "TỔNG NHÓM TRM + TPS" in combined or "TỔNG NHÓM TRM+TPS" in combined:
                markers["TRM_TPS_END"] = r
            elif "TỔNG NHÓM TCR + TML" in combined or "TỔNG NHÓM TCR+TML" in combined:
                markers["TCR_TML_END"] = r
            elif "TỔNG NHÓM TMR" in combined:
                markers["TMR_END"] = r

        is_unified = bool(markers.get("TRM_TPS_END") and markers.get("TCR_TML_END"))

        if is_unified:
            if report_type == "TRM_TPS":
                b_start = 29
                b_end = markers["TRM_TPS_END"] - 1
                default_th = "TRM"
            elif report_type == "TCR_TML":
                b_start = markers["TRM_TPS_END"] + 1
                b_end = markers["TCR_TML_END"] - 1
                default_th = "TCR"
            else:  # TMR
                b_start = markers["TCR_TML_END"] + 1
                b_end = markers.get("TMR_END", ws.max_row) - 1
                default_th = "TMR"

            app_dealers = rule_manager.get_dealers(report_type) if rule_manager else []
            app_by_name = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", default_th)): d for d in app_dealers}
            app_by_code = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", default_th)): d for d in app_dealers if d.get("ma_khach")}

            for r in range(b_start, b_end + 1):
                ex_code = str(ws.cell(r, 1).value or "").strip()
                if ex_code.upper() in ["NONE", "MÃ KHÁCH HÀNG", "MÃ KH"]:
                    ex_code = ""
                ex_name = str(ws.cell(r, 4).value or "").strip()
                ex_grp = str(ws.cell(r, 5).value or "").strip()
                th = str(ws.cell(r, 6).value or default_th).strip().upper()

                if not ex_name and not ex_code:
                    continue

                matched = None
                if ex_code:
                    matched = app_by_code.get((ex_code.upper(), th))
                if not matched and ex_name:
                    matched = app_by_name.get((normalize_dealer_name(ex_name), th))
                if not matched and ex_name:
                    name_norm = normalize_dealer_name(ex_name)
                    cand = [d for d in app_dealers if normalize_dealer_name(d.get("ten_khach", "")) == name_norm]
                    if len(cand) == 1:
                        matched = cand[0]

                app_code = str(matched.get("ma_khach") or "").strip() if matched else ""
                app_grp = str(matched.get("nhom") or "").strip() if matched else ""

                has_code_change = bool(app_code and app_code.upper() != ex_code.upper())
                has_group_change = bool(app_grp and app_grp.upper() != ex_grp.upper())
                is_conflict = bool(ex_code and app_code and ex_code.upper() != app_code.upper())
                needs_verify = bool("BẢO VIỆT" in ex_name.upper() or app_code == "CH9999")

                note = ""
                if is_conflict:
                    note = f"⚠️ Xung đột mã: Excel='{ex_code}' vs App='{app_code}'"
                elif needs_verify:
                    note = "⚠️ Cần xác minh mã Bảo Việt trước khi ghi (mã tạm CH9999)"
                elif has_code_change and has_group_change:
                    note = f"Cập nhật mã '{app_code}' & nhóm '{app_grp}'"
                elif has_code_change:
                    note = f"Cập nhật mã '{app_code}'"
                elif has_group_change:
                    note = f"Cập nhật nhóm '{app_grp}'"
                else:
                    note = "Đã khớp chuẩn"

                items.append({
                    "row_num": r,
                    "th": th,
                    "name": ex_name,
                    "excel_code": ex_code,
                    "app_code": app_code if app_code else ex_code,
                    "excel_group": ex_grp,
                    "app_group": app_grp if app_grp else ex_grp,
                    "has_code_change": has_code_change,
                    "has_group_change": has_group_change,
                    "has_change": has_code_change or has_group_change,
                    "is_conflict": is_conflict,
                    "needs_verify": needs_verify,
                    "note": note,
                    "is_checked": not is_conflict
                })

            wb.close()
            return items

        # Fallback cho mẫu cũ riêng rẽ
        if report_type == "TCR_TML":
            sub_tcr_row = None
            sub_tml_row = None
            for r in range(28, ws.max_row + 1):
                for c in range(1, 5):
                    v = str(ws.cell(r, c).value or "").upper()
                    if "TỔNG THƯƠNG HIỆU TOCERA" in v and not sub_tcr_row:
                        sub_tcr_row = r
                    elif "TỔNG THƯƠNG HIỆU TOMILAN" in v and not sub_tml_row:
                        sub_tml_row = r
            if not sub_tcr_row: sub_tcr_row = 42
            if not sub_tml_row: sub_tml_row = 52

            app_dealers = rule_manager.get_dealers("TCR_TML") if rule_manager else []
            app_by_name = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "TCR")): d for d in app_dealers}
            app_by_code = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", "TCR")): d for d in app_dealers if d.get("ma_khach")}

            for r in range(29, sub_tml_row):
                if r == sub_tcr_row: continue
                th = "TCR" if r < sub_tcr_row else "TML"
                ex_code = str(ws.cell(r, 1).value or "").strip()
                if ex_code.upper() in ["NONE", "MÃ KHÁCH HÀNG", "MÃ KH"]: ex_code = ""
                ex_name = str(ws.cell(r, 4).value or "").strip()
                ex_grp = str(ws.cell(r, 5).value or "").strip()
                if not ex_name and not ex_code: continue

                matched = app_by_code.get((ex_code.upper(), th)) if ex_code else None
                if not matched and ex_name: matched = app_by_name.get((normalize_dealer_name(ex_name), th))
                app_code = str(matched.get("ma_khach") or "").strip() if matched else ""
                app_grp = str(matched.get("nhom") or "").strip() if matched else ""

                items.append({
                    "row_num": r, "th": th, "name": ex_name,
                    "excel_code": ex_code, "app_code": app_code if app_code else ex_code,
                    "excel_group": ex_grp, "app_group": app_grp if app_grp else ex_grp,
                    "has_code_change": bool(app_code and app_code.upper() != ex_code.upper()),
                    "has_group_change": bool(app_grp and app_grp.upper() != ex_grp.upper()),
                    "has_change": bool(app_code and app_code.upper() != ex_code.upper()) or bool(app_grp and app_grp.upper() != ex_grp.upper()),
                    "is_conflict": bool(ex_code and app_code and ex_code.upper() != app_code.upper()),
                    "needs_verify": bool("BẢO VIỆT" in ex_name.upper() or app_code == "CH9999"),
                    "note": "Chuẩn hóa", "is_checked": not bool(ex_code and app_code and ex_code.upper() != app_code.upper())
                })

        elif report_type == "TRM_TPS":
            sub_trm_row = None
            sub_tps_row = None
            for r in range(24, ws.max_row + 1):
                for c in range(1, 4):
                    v = str(ws.cell(r, c).value or "").upper()
                    if "TỔNG TOROMA" in v and not sub_trm_row:
                        sub_trm_row = r
                    elif "TỔNG TOPISA" in v and not sub_tps_row:
                        sub_tps_row = r
            if not sub_trm_row: sub_trm_row = 34
            if not sub_tps_row: sub_tps_row = 43

            app_dealers = rule_manager.get_dealers("TRM_TPS") if rule_manager else []
            app_by_name = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "TRM")): d for d in app_dealers}
            app_by_code = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", "TRM")): d for d in app_dealers if d.get("ma_khach")}

            for r in range(25, sub_tps_row):
                if r in [sub_trm_row, sub_trm_row + 1]:
                    continue
                th = "TRM" if r < sub_trm_row else "TPS"
                ex_code = str(ws.cell(r, 1).value or "").strip()
                if ex_code.upper() in ["NONE", "MÃ KHÁCH HÀNG", "MÃ KH"]: ex_code = ""
                ex_name = str(ws.cell(r, 3).value or "").strip()
                ex_grp = str(ws.cell(r, 4).value or "").strip()
                if not ex_name or "CỬA HÀNG" in ex_name.upper(): continue

                matched = app_by_code.get((ex_code.upper(), th)) if ex_code else None
                if not matched and ex_name: matched = app_by_name.get((normalize_dealer_name(ex_name), th))
                app_code = str(matched.get("ma_khach") or "").strip() if matched else ""
                app_grp = str(matched.get("nhom") or "").strip() if matched else ""

                items.append({
                    "row_num": r, "th": th, "name": ex_name,
                    "excel_code": ex_code, "app_code": app_code if app_code else ex_code,
                    "excel_group": ex_grp, "app_group": app_grp if app_grp else ex_grp,
                    "has_code_change": bool(app_code and app_code.upper() != ex_code.upper()),
                    "has_group_change": bool(app_grp and app_grp.upper() != ex_grp.upper()),
                    "has_change": bool(app_code and app_code.upper() != ex_code.upper()) or bool(app_grp and app_grp.upper() != ex_grp.upper()),
                    "is_conflict": bool(ex_code and app_code and ex_code.upper() != app_code.upper()),
                    "needs_verify": bool("BẢO VIỆT" in ex_name.upper() or app_code == "CH9999"),
                    "note": "Chuẩn hóa", "is_checked": not bool(ex_code and app_code and ex_code.upper() != app_code.upper())
                })

        elif report_type == "TMR":
            sub_tmr_row = None
            for r in range(10, ws.max_row + 1):
                for c in range(1, 4):
                    v = str(ws.cell(r, c).value or "").upper()
                    if "TỔNG (M²)" in v or "TỔNG" in v:
                        sub_tmr_row = r
                        break
                if sub_tmr_row: break
            if not sub_tmr_row: sub_tmr_row = 13

            app_dealers = rule_manager.get_dealers("TMR", brand="TMR") if rule_manager else []
            app_by_name = {normalize_dealer_name(d.get("ten_khach", "")): d for d in app_dealers}
            app_by_code = {str(d.get("ma_khach") or "").strip().upper(): d for d in app_dealers if d.get("ma_khach")}

            has_code_col = (ws.cell(9, 1).value == "Mã khách hàng" or ws.cell(3, 1).value == "Mã khách hàng")

            for r in range(10, sub_tmr_row):
                if has_code_col:
                    ex_code = str(ws.cell(r, 1).value or "").strip()
                    if ex_code.upper() in ["NONE", "MÃ KHÁCH HÀNG", "MÃ KH"]: ex_code = ""
                    ex_name = str(ws.cell(r, 3).value or "").strip()
                    ex_grp = str(ws.cell(r, 4).value or "").strip()
                else:
                    ex_code = ""
                    ex_name = str(ws.cell(r, 2).value or "").strip()
                    ex_grp = str(ws.cell(r, 3).value or "").strip()

                if not ex_name or "CỬA HÀNG" in ex_name.upper(): continue

                matched = app_by_code.get(ex_code.upper()) if ex_code else None
                if not matched and ex_name: matched = app_by_name.get(normalize_dealer_name(ex_name))
                app_code = str(matched.get("ma_khach") or "").strip() if matched else ""
                app_grp = str(matched.get("nhom") or "").strip() if matched else ""

                items.append({
                    "row_num": r, "th": "TMR", "name": ex_name,
                    "excel_code": ex_code, "app_code": app_code if app_code else ex_code,
                    "excel_group": ex_grp, "app_group": app_grp if app_grp else ex_grp,
                    "has_code_change": bool(app_code and app_code.upper() != ex_code.upper()),
                    "has_group_change": bool(app_grp and app_grp.upper() != ex_grp.upper()),
                    "has_change": bool(app_code and app_code.upper() != ex_code.upper()) or bool(app_grp and app_grp.upper() != ex_grp.upper()),
                    "is_conflict": bool(ex_code and app_code and ex_code.upper() != app_code.upper()),
                    "needs_verify": bool("BẢO VIỆT" in ex_name.upper() or app_code == "CH9999"),
                    "note": "Chuẩn hóa", "is_checked": not bool(ex_code and app_code and ex_code.upper() != app_code.upper())
                })

        wb.close()
        return items

    @staticmethod
    def update_summary_excel(
        template_path: str,
        report_type: str,
        app_result: Dict[str, Any],
        target_date: date,
        backup_dir: str,
        rule_manager: Optional[Any] = None,
        confirmed_items: Optional[List[Dict[str, Any]]] = None
    ) -> str:
        """
        Cập nhật kết quả vào tệp Excel tổng, giữ nguyên định dạng, công thức và bố cục.
        - Mục 01: Có mã khách và CHỈ hiện đại lý có phát sinh hợp lệ.
        - Mục 02: Đồng bộ danh sách theo dõi, tự động chèn dòng đại lý mới vào Mục 02, tránh sinh dòng trùng khi chạy lại.
        - Ghi an toàn bằng tệp tạm và kiểm tra đối chiếu trước khi thay thế tệp chính.
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Không tìm thấy tệp Excel tổng: {template_path}")

        dealers = rule_manager.get_dealers(report_type) if rule_manager else []
        ExcelWriter.validate_template(template_path, report_type, target_date, dealers)

        ExcelWriter.create_daily_backup(template_path, backup_dir, target_date)

        wb = openpyxl.load_workbook(template_path, data_only=False)

        day_str = target_date.strftime("%d/%m/%Y")
        sheet_day_str = target_date.strftime("%d.%m.%Y")
        day_header_short = target_date.strftime("%d/%m")

        muc01_rows = app_result.get("muc01_rows", [])

        # Kiểm tra bố cục 3 khối chuẩn trong workbook
        first_ws = wb.active
        markers = {}
        for r in range(1, min(first_ws.max_row + 1, 120)):
            vals = [str(first_ws.cell(r, c).value or "").strip().upper() for c in range(1, 6)]
            combined = " | ".join(vals)
            if "TỔNG NHÓM TRM + TPS" in combined or "TỔNG NHÓM TRM+TPS" in combined:
                markers["TRM_TPS_END"] = r
            elif "TỔNG NHÓM TCR + TML" in combined or "TỔNG NHÓM TCR+TML" in combined:
                markers["TCR_TML_END"] = r
            elif "TỔNG NHÓM TMR" in combined:
                markers["TMR_END"] = r
            elif "TỔNG CHUNG" in combined:
                markers["GRAND_TOTAL"] = r

        is_unified = bool(markers.get("TRM_TPS_END") and markers.get("TCR_TML_END"))

        if is_unified:
            # === CẬP NHẬT TỆP EXCEL TỔNG HỢP NHẤT (CHUẨN 3 KHỐI) ===
            is_new_sheet = False
            if sheet_day_str in wb.sheetnames:
                ws = wb[sheet_day_str]
            else:
                base_ws = wb["09.11.2026"] if "09.11.2026" in wb.sheetnames else wb.active
                ws = wb.copy_worksheet(base_ws)
                ws.title = sheet_day_str
                is_new_sheet = True

            ws.cell(1, 2).value = f"BÁO CÁO BÁN HÀNG NGÀY {day_str} NHÓM VIP - D1 - GPP"

            ws.cell(2, 1).value = "Mã khách hàng"
            ws.cell(2, 2).value = "STT"
            ws.cell(2, 4).value = "CỬA HÀNG"
            ws.cell(2, 5).value = "NHÓM"
            ws.cell(2, 6).value = "TH"
            ws.cell(2, 7).value = "∑SL MÉT"
            ws.cell(2, 8).value = "∑SL THÙNG"

            col_map = {}
            for c in range(9, 26):
                val = ws.cell(2, c).value
                if val:
                    col_map[str(val).replace('\n', ' ').strip()] = c

            existing_other_rows = []
            if not is_new_sheet:
                for r in range(3, 25):
                    name = ws.cell(r, 4).value
                    th = ws.cell(r, 6).value
                    if name and str(name).strip() and "TỔNG" not in str(name).upper():
                        th_clean = str(th or "").strip().upper()
                        row_grp = None
                        if th_clean in ["TRM", "TPS"]: row_grp = "TRM_TPS"
                        elif th_clean in ["TCR", "TML"]: row_grp = "TCR_TML"
                        elif th_clean == "TMR": row_grp = "TMR"

                        if row_grp and row_grp != report_type:
                            row_specs = {}
                            for sp_name, col_idx in col_map.items():
                                val = ws.cell(r, col_idx).value
                                if val:
                                    try: row_specs[sp_name] = float(val)
                                    except Exception: pass
                            existing_other_rows.append({
                                "ma_khach": str(ws.cell(r, 1).value or "").strip(),
                                "ten_khach": str(name).strip(),
                                "nhom": str(ws.cell(r, 5).value or "").strip(),
                                "thuong_hieu": th_clean,
                                "specs": row_specs
                            })

            combined_rows = []
            for item in existing_other_rows:
                if item["thuong_hieu"] in ["TRM", "TPS"]: combined_rows.append(item)
            if report_type == "TRM_TPS": combined_rows.extend(muc01_rows)

            for item in existing_other_rows:
                if item["thuong_hieu"] in ["TCR", "TML"]: combined_rows.append(item)
            if report_type == "TCR_TML": combined_rows.extend(muc01_rows)

            for item in existing_other_rows:
                if item["thuong_hieu"] == "TMR": combined_rows.append(item)
            if report_type == "TMR": combined_rows.extend(muc01_rows)

            start_r = 3
            end_r = 24
            num_active = len(combined_rows)
            for i in range(end_r - start_r + 1):
                curr_r = start_r + i
                if i < num_active:
                    cust_data = combined_rows[i]
                    ws.cell(curr_r, 1).value = cust_data.get("ma_khach", "")
                    ws.cell(curr_r, 2).value = i + 1
                    ws.cell(curr_r, 4).value = cust_data.get("ten_khach", "")
                    ws.cell(curr_r, 5).value = cust_data.get("nhom", "")
                    ws.cell(curr_r, 6).value = cust_data.get("thuong_hieu", "")

                    ws.cell(curr_r, 7).value = f"=I{curr_r}+J{curr_r}*1.5+K{curr_r}+L{curr_r}*1.62+M{curr_r}*1.8+N{curr_r}*1.44+O{curr_r}*1.44+P{curr_r}*1.92+Q{curr_r}*0.99+R{curr_r}*1.44+S{curr_r}*0.96+T{curr_r}*1.92+U{curr_r}*1.92+V{curr_r}*1.5+W{curr_r}*1.8+X{curr_r}*1.8+Y{curr_r}*1.92"
                    ws.cell(curr_r, 8).value = f"=SUM(I{curr_r}:Y{curr_r})"

                    for c in range(9, 26):
                        ws.cell(curr_r, c).value = None

                    for sp, qty in cust_data["specs"].items():
                        c_idx = match_column(sp, col_map)
                        if c_idx:
                            ws.cell(curr_r, c_idx).value = qty

                    ws.row_dimensions[curr_r].hidden = False
                else:
                    ws.cell(curr_r, 1).value = None
                    ws.cell(curr_r, 2).value = None
                    ws.cell(curr_r, 4).value = None
                    ws.cell(curr_r, 5).value = None
                    ws.cell(curr_r, 6).value = None
                    ws.cell(curr_r, 7).value = f"=I{curr_r}+J{curr_r}*1.5+K{curr_r}+L{curr_r}*1.62+M{curr_r}*1.8+N{curr_r}*1.44+O{curr_r}*1.44+P{curr_r}*1.92+Q{curr_r}*0.99+R{curr_r}*1.44+S{curr_r}*0.96+T{curr_r}*1.92+U{curr_r}*1.92+V{curr_r}*1.5+W{curr_r}*1.8+X{curr_r}*1.8+Y{curr_r}*1.92"
                    ws.cell(curr_r, 8).value = f"=SUM(I{curr_r}:Y{curr_r})"
                    for c in range(9, 26):
                        ws.cell(curr_r, c).value = None
                    ws.row_dimensions[curr_r].hidden = True

            ws.row_dimensions[25].hidden = False
            ws.cell(25, 2).value = "TỔNG"
            ws.cell(25, 7).value = "=SUM(G3:G24)"
            ws.cell(25, 8).value = "=SUM(H3:H24)"
            for c in range(9, 26):
                cl = get_column_letter(c)
                ws.cell(25, c).value = f"=SUM({cl}3:{cl}24)"

            ws.cell(27, 1).value = "Mã khách hàng"
            ws.cell(27, 2).value = "STT"
            ws.cell(27, 4).value = "CỬA HÀNG"
            ws.cell(27, 5).value = "NHÓM"
            ws.cell(27, 6).value = "TH"
            ws.cell(27, 7).value = "∑SL MÉT"

            target_col = ExcelWriter.find_day_column(ws, 27, 8, target_date)
            if not target_col:
                target_col = ws.max_column + 1
                ws.cell(27, target_col).value = target_date
                ws.cell(28, target_col).value = day_header_short

            if confirmed_items:
                ExcelWriter.apply_confirmed_sync(ws, report_type, confirmed_items, is_unified=True)

            sub_trm_row = None
            sub_tcr_row = None
            sub_tmr_row = None
            grand_total_row = None

            for r in range(28, ws.max_row + 1):
                vals = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, 6)]
                comb = " | ".join(vals)
                if "TỔNG NHÓM TRM + TPS" in comb or "TỔNG NHÓM TRM+TPS" in comb:
                    sub_trm_row = r
                elif "TỔNG NHÓM TCR + TML" in comb or "TỔNG NHÓM TCR+TML" in comb:
                    sub_tcr_row = r
                elif "TỔNG NHÓM TMR" in comb:
                    sub_tmr_row = r
                elif "TỔNG CHUNG" in comb:
                    grand_total_row = r

            if not sub_trm_row: sub_trm_row = 45
            if not sub_tcr_row: sub_tcr_row = 68
            if not sub_tmr_row: sub_tmr_row = 72
            if not grand_total_row: grand_total_row = 73

            if report_type == "TRM_TPS":
                b_start = 29
                curr_sub_row = sub_trm_row
            elif report_type == "TCR_TML":
                b_start = sub_trm_row + 1
                curr_sub_row = sub_tcr_row
            else:
                b_start = sub_tcr_row + 1
                curr_sub_row = sub_tmr_row

            existing_dealers = {}
            for r in range(b_start, curr_sub_row):
                code = str(ws.cell(r, 1).value or "").strip().upper()
                name = str(ws.cell(r, 4).value or "").strip()
                if code or name:
                    existing_dealers[(code, normalize_dealer_name(name))] = r

            dealers = rule_manager.get_dealers(report_type) if rule_manager else []
            for d in dealers:
                d_code = str(d.get("ma_khach") or "").strip().upper()
                d_name = str(d.get("ten_khach") or "").strip()
                d_name_norm = normalize_dealer_name(d_name)
                found = any((d_code and d_code == k[0]) or (d_name_norm and d_name_norm == k[1]) for k in existing_dealers)
                if not found:
                    new_r = curr_sub_row
                    ExcelWriter.insert_dealer_row(ws, new_r)
                    ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 26)
                    ws.cell(new_r, 1).value = d.get("ma_khach", "")
                    ws.cell(new_r, 2).value = new_r - b_start + 1
                    ws.cell(new_r, 4).value = d.get("ten_khach", "")
                    ws.cell(new_r, 5).value = d.get("nhom", "")
                    ws.cell(new_r, 6).value = d.get("thuong_hieu", "")
                    ws.cell(new_r, 7).value = f"=SUM(H{new_r}:{get_column_letter(max(26, ws.max_column))}{new_r})"
                    curr_sub_row += 1
                    if report_type == "TRM_TPS":
                        sub_trm_row += 1
                        sub_tcr_row += 1
                        sub_tmr_row += 1
                        grand_total_row += 1
                    elif report_type == "TCR_TML":
                        sub_tcr_row += 1
                        sub_tmr_row += 1
                        grand_total_row += 1
                    else:
                        sub_tmr_row += 1
                        grand_total_row += 1
                    existing_dealers[(d_code, d_name_norm)] = new_r

            b_end = curr_sub_row - 1

            app_by_code_th = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", "").strip().upper()): d for d in dealers if d.get("ma_khach")}
            app_by_name_th = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "").strip().upper()): d for d in dealers}
            app_by_name_only = {normalize_dealer_name(d.get("ten_khach", "")): d for d in dealers}

            muc01_by_code_th = {(row["ma_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row for row in muc01_rows if row.get("ma_khach")}
            muc01_by_name_th = {(normalize_dealer_name(row["ten_khach"]), row["thuong_hieu"].strip().upper()): row for row in muc01_rows}

            for r in range(b_start, b_end + 1):
                name = str(ws.cell(r, 4).value or "").strip()
                raw_code = str(ws.cell(r, 1).value or "").strip()
                th = str(ws.cell(r, 6).value or "").strip().upper()
                name_norm = normalize_dealer_name(name)

                rule_dealer = None
                if raw_code:
                    rule_dealer = app_by_code_th.get((raw_code.upper(), th))
                if not rule_dealer and name_norm:
                    rule_dealer = app_by_name_th.get((name_norm, th))
                if not rule_dealer and name_norm:
                    rule_dealer = app_by_name_only.get(name_norm)

                if confirmed_items is None and not raw_code and rule_dealer and rule_dealer.get("ma_khach"):
                    if not ("BẢO VIỆT" in name.upper() and rule_dealer["ma_khach"] == "CH9999"):
                        raw_code = rule_dealer["ma_khach"]
                        ws.cell(r, 1).value = raw_code

                if rule_dealer:
                    if confirmed_items is None and not ws.cell(r, 5).value and rule_dealer.get("nhom"):
                        ws.cell(r, 5).value = rule_dealer["nhom"]
                    if not ws.cell(r, 6).value and rule_dealer.get("thuong_hieu"):
                        ws.cell(r, 6).value = rule_dealer["thuong_hieu"]

                matched_data = None
                if raw_code:
                    matched_data = muc01_by_code_th.get((raw_code.upper(), th))
                if not matched_data and name_norm:
                    matched_data = muc01_by_name_th.get((name_norm, th))

                if matched_data:
                    ws.cell(r, target_col).value = matched_data["tong_m2"]
                else:
                    ws.cell(r, target_col).value = None

            cl_t = get_column_letter(target_col)
            ws.cell(sub_trm_row, target_col).value = f"=SUM({cl_t}29:{cl_t}{sub_trm_row - 1})"
            ws.cell(sub_tcr_row, target_col).value = f"=SUM({cl_t}{sub_trm_row + 1}:{cl_t}{sub_tcr_row - 1})"
            ws.cell(sub_tmr_row, target_col).value = f"=SUM({cl_t}{sub_tcr_row + 1}:{cl_t}{sub_tmr_row - 1})"
            ws.cell(grand_total_row, target_col).value = f"=SUM({cl_t}{sub_trm_row},{cl_t}{sub_tcr_row},{cl_t}{sub_tmr_row})"

            ws.cell(sub_trm_row, 7).value = f"=SUM(G29:G{sub_trm_row - 1})"
            ws.cell(sub_tcr_row, 7).value = f"=SUM(G{sub_trm_row + 1}:G{sub_tcr_row - 1})"
            ws.cell(sub_tmr_row, 7).value = f"=SUM(G{sub_tcr_row + 1}:G{sub_tmr_row - 1})"
            ws.cell(grand_total_row, 7).value = f"=SUM(G{sub_trm_row},G{sub_tcr_row},G{sub_tmr_row})"

        else:
            # === CẬP NHẬT TỆP EXCEL RIÊNG RẼ CHO TỪNG NHÓM (MẪU ĐƠN LẺ) ===
            if report_type == "TCR_TML":
                if sheet_day_str in wb.sheetnames:
                    ws = wb[sheet_day_str]
                else:
                    base_ws = wb["09.11.2026"] if "09.11.2026" in wb.sheetnames else wb.active
                    ws = wb.copy_worksheet(base_ws)
                    ws.title = sheet_day_str

                ws.cell(1, 2).value = f"BÁO CÁO BÁN HÀNG NGÀY {day_str} NHÓM VIP - D1 - GPP TOCERA VÀ TOMILAN"

                ws.cell(2, 1).value = "Mã khách hàng"
                ws.cell(2, 2).value = "STT"
                ws.cell(2, 4).value = "CỬA HÀNG"
                ws.cell(2, 5).value = "NHÓM"
                ws.cell(2, 6).value = "TH"
                ws.cell(2, 7).value = "∑SL MÉT"
                ws.cell(2, 8).value = "∑SL THÙNG"

                col_map = {}
                for c in range(9, ws.max_column + 1):
                    val = ws.cell(2, c).value
                    if val:
                        col_map[str(val).replace('\n', ' ').strip()] = c

                start_r = 3
                end_r = 24
                num_active = len(muc01_rows)
                for i in range(end_r - start_r + 1):
                    curr_r = start_r + i
                    if i < num_active:
                        cust_data = muc01_rows[i]
                        ws.cell(curr_r, 1).value = cust_data.get("ma_khach", "")
                        ws.cell(curr_r, 2).value = i + 1
                        ws.cell(curr_r, 4).value = cust_data.get("ten_khach", "")
                        ws.cell(curr_r, 5).value = cust_data.get("nhom", "")
                        ws.cell(curr_r, 6).value = cust_data.get("thuong_hieu", "")

                        ws.cell(curr_r, 7).value = f"=I{curr_r}+J{curr_r}*1.5+K{curr_r}+L{curr_r}*1.62+M{curr_r}*1.8+N{curr_r}*1.44+O{curr_r}*1.44+P{curr_r}*1.92+Q{curr_r}*0.99+R{curr_r}*1.44+S{curr_r}*0.96+T{curr_r}*1.92+U{curr_r}*1.92+V{curr_r}*1.5+W{curr_r}*1.8+X{curr_r}*1.8+Y{curr_r}*1.92"
                        ws.cell(curr_r, 8).value = f"=SUM(I{curr_r}:Y{curr_r})"

                        for c in range(9, ws.max_column + 1):
                            ws.cell(curr_r, c).value = None

                        for sp, qty in cust_data["specs"].items():
                            c_idx = match_column(sp, col_map)
                            if c_idx:
                                ws.cell(curr_r, c_idx).value = qty

                        ws.row_dimensions[curr_r].hidden = False
                    else:
                        ws.cell(curr_r, 1).value = None
                        ws.cell(curr_r, 2).value = None
                        ws.cell(curr_r, 4).value = None
                        ws.cell(curr_r, 5).value = None
                        ws.cell(curr_r, 6).value = None
                        ws.cell(curr_r, 7).value = f"=I{curr_r}+J{curr_r}*1.5+K{curr_r}+L{curr_r}*1.62+M{curr_r}*1.8+N{curr_r}*1.44+O{curr_r}*1.44+P{curr_r}*1.92+Q{curr_r}*0.99+R{curr_r}*1.44+S{curr_r}*0.96+T{curr_r}*1.92+U{curr_r}*1.92+V{curr_r}*1.5+W{curr_r}*1.8+X{curr_r}*1.8+Y{curr_r}*1.92"
                        ws.cell(curr_r, 8).value = f"=SUM(I{curr_r}:Y{curr_r})"
                        for c in range(9, ws.max_column + 1):
                            ws.cell(curr_r, c).value = None
                        ws.row_dimensions[curr_r].hidden = True

                ws.row_dimensions[25].hidden = False
                ws.cell(25, 2).value = "TỔNG"
                ws.cell(25, 7).value = "=SUM(G3:G24)"
                ws.cell(25, 8).value = "=SUM(H3:H24)"
                for c in range(9, 26):
                    cl = get_column_letter(c)
                    ws.cell(25, c).value = f"=SUM({cl}3:{cl}24)"

                ws.cell(27, 1).value = "Mã khách hàng"
                ws.cell(27, 2).value = "STT"
                ws.cell(27, 4).value = "CỬA HÀNG"
                ws.cell(27, 5).value = "NHÓM"
                ws.cell(27, 6).value = "TH"
                ws.cell(27, 7).value = "∑SL MÉT"

                sub_tcr_row = None
                sub_tml_row = None
                for r in range(28, ws.max_row + 1):
                    for c in range(1, 5):
                        v = str(ws.cell(r, c).value or "").upper()
                        if "TỔNG THƯƠNG HIỆU TOCERA" in v and not sub_tcr_row:
                            sub_tcr_row = r
                        elif "TỔNG THƯƠNG HIỆU TOMILAN" in v and not sub_tml_row:
                            sub_tml_row = r
                if not sub_tcr_row: sub_tcr_row = 42
                if not sub_tml_row: sub_tml_row = 52

                if confirmed_items:
                    ExcelWriter.apply_confirmed_sync(ws, report_type, confirmed_items, is_unified=False)

                all_dealers = rule_manager.get_dealers("TCR_TML") if rule_manager else []
                tcr_dealers = [d for d in all_dealers if d.get("thuong_hieu") == "TCR"]
                tml_dealers = [d for d in all_dealers if d.get("thuong_hieu") == "TML"]

                # 1. Sync TCR dealers
                existing_tcr = {}
                for r in range(29, sub_tcr_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 4).value or "").strip()
                    if code or name:
                        existing_tcr[(code, normalize_dealer_name(name))] = r

                for d in tcr_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip()
                    d_name_norm = normalize_dealer_name(d_name)
                    found = any((d_code and d_code == k[0]) or (d_name_norm and d_name_norm == k[1]) for k in existing_tcr)
                    if not found:
                        new_r = sub_tcr_row
                        ExcelWriter.insert_dealer_row(ws, new_r)
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 26)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_tcr_row - 28
                        ws.cell(new_r, 4).value = d.get("ten_khach", "")
                        ws.cell(new_r, 5).value = d.get("nhom", "")
                        ws.cell(new_r, 6).value = "TCR"
                        sub_tcr_row += 1
                        sub_tml_row += 1
                        existing_tcr[(d_code, d_name_norm)] = new_r

                # 2. Sync TML dealers
                existing_tml = {}
                tml_start_row = sub_tcr_row + 1
                for r in range(tml_start_row, sub_tml_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 4).value or "").strip()
                    if code or name:
                        existing_tml[(code, normalize_dealer_name(name))] = r

                for d in tml_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip()
                    d_name_norm = normalize_dealer_name(d_name)
                    found = any((d_code and d_code == k[0]) or (d_name_norm and d_name_norm == k[1]) for k in existing_tml)
                    if not found:
                        new_r = sub_tml_row
                        ExcelWriter.insert_dealer_row(ws, new_r)
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 26)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_tml_row - tml_start_row + 1
                        ws.cell(new_r, 4).value = d.get("ten_khach", "")
                        ws.cell(new_r, 5).value = d.get("nhom", "")
                        ws.cell(new_r, 6).value = "TML"
                        sub_tml_row += 1
                        existing_tml[(d_code, d_name_norm)] = new_r

                target_col = ExcelWriter.find_day_column(ws, 27, 8, target_date)
                if not target_col:
                    target_col = ws.max_column + 1
                    ws.cell(27, target_col).value = target_date
                    ws.cell(28, target_col).value = day_header_short

                app_dealers = rule_manager.get_dealers("TCR_TML") if rule_manager else []
                app_by_name_th = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "TCR")): d for d in app_dealers}
                app_by_code_th = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", "TCR")): d for d in app_dealers if d.get("ma_khach")}
                app_by_name_only = {normalize_dealer_name(d.get("ten_khach", "")): d for d in app_dealers}

                muc01_by_code_th = {(row["ma_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row for row in muc01_rows if row.get("ma_khach")}
                muc01_by_name_th = {(normalize_dealer_name(row["ten_khach"]), row["thuong_hieu"].strip().upper()): row for row in muc01_rows}

                for r in range(29, sub_tml_row):
                    if r == sub_tcr_row: continue
                    name = str(ws.cell(r, 4).value or "").strip()
                    raw_code = str(ws.cell(r, 1).value or "").strip()
                    th = str(ws.cell(r, 6).value or ("TCR" if r < sub_tcr_row else "TML")).strip().upper()
                    name_norm = normalize_dealer_name(name)

                    rule_dealer = None
                    if raw_code: rule_dealer = app_by_code_th.get((raw_code.upper(), th))
                    if not rule_dealer and name_norm: rule_dealer = app_by_name_th.get((name_norm, th))
                    if not rule_dealer and name_norm: rule_dealer = app_by_name_only.get(name_norm)

                    if confirmed_items is None and not raw_code and rule_dealer and rule_dealer.get("ma_khach"):
                        raw_code = rule_dealer["ma_khach"]
                        ws.cell(r, 1).value = raw_code

                    if rule_dealer:
                        if confirmed_items is None and not ws.cell(r, 5).value and rule_dealer.get("nhom"):
                            ws.cell(r, 5).value = rule_dealer["nhom"]
                        ws.cell(r, 6).value = th

                    matched_data = None
                    if raw_code: matched_data = muc01_by_code_th.get((raw_code.upper(), th))
                    if not matched_data and name_norm: matched_data = muc01_by_name_th.get((name_norm, th))

                    if matched_data:
                        ws.cell(r, target_col).value = matched_data["tong_m2"]
                    else:
                        ws.cell(r, target_col).value = None

                cl_t = get_column_letter(target_col)
                ws.cell(sub_tcr_row, target_col).value = f"=SUM({cl_t}29:{cl_t}{sub_tcr_row - 1})"
                ws.cell(sub_tml_row, target_col).value = f"=SUM({cl_t}{sub_tcr_row + 1}:{cl_t}{sub_tml_row - 1})"
                row_tot_m2 = sub_tml_row + 1
                ws.cell(row_tot_m2, target_col).value = f"=SUM({cl_t}{sub_tcr_row},{cl_t}{sub_tml_row})"

            elif report_type == "TRM_TPS":
                ws = ExcelWriter.get_target_sheet(wb, report_type, target_date)

                if ws.cell(3, 5).value != "TH":
                    ws.insert_cols(5)
                    for r in range(1, min(ws.max_row + 1, 60)):
                        c4 = ws.cell(r, 4)
                        c5 = ws.cell(r, 5)
                        if c4.has_style:
                            c5.font = copy(c4.font)
                            c5.border = copy(c4.border)
                            c5.fill = copy(c4.fill)
                            c5.number_format = c4.number_format
                            c5.alignment = Alignment(horizontal="center", vertical="center")

                ws.cell(3, 1).value = "Mã khách hàng"
                ws.cell(3, 2).value = "Stt"
                ws.cell(3, 3).value = "CỬA HÀNG"
                ws.cell(3, 4).value = "NHÓM"
                ws.cell(3, 5).value = "TH"
                ws.cell(3, 6).value = "TỔNG (m²)"
                ws.cell(3, 7).value = "TỔNG (thùng)"

                try:
                    v2 = str(ws.cell(2, 2).value or "").strip()
                    if "NGÀY" in v2.upper():
                        prefix = v2.split("NGÀY")[0].strip()
                        ws.cell(2, 2).value = f"{prefix} NGÀY {day_str}"
                    else:
                        ws.cell(2, 2).value = f"BÁO CÁO BÁN HÀNG NHÓM ĐẠI LÝ TOROMA VÀ TOPISA NGÀY {day_str}"
                except Exception:
                    pass

                col_map = {}
                for c in range(8, ws.max_column + 1):
                    val = ws.cell(3, c).value
                    if val:
                        col_map[str(val).replace('\n', ' ').strip()] = c

                start_r = 4
                end_r = 20
                num_active = len(muc01_rows)
                for i in range(end_r - start_r + 1):
                    curr_r = start_r + i
                    if i < num_active:
                        cust_data = muc01_rows[i]
                        ws.cell(curr_r, 1).value = cust_data.get("ma_khach", "")
                        ws.cell(curr_r, 2).value = i + 1
                        ws.cell(curr_r, 3).value = cust_data.get("ten_khach", "")
                        ws.cell(curr_r, 4).value = cust_data.get("nhom", "")
                        ws.cell(curr_r, 5).value = cust_data.get("thuong_hieu", "")

                        ws.cell(curr_r, 6).value = f"=H{curr_r}*0.96+I{curr_r}*1.5+J{curr_r}+K{curr_r}*1.8+L{curr_r}*1.8+M{curr_r}*1.8+N{curr_r}*1.92+O{curr_r}*1.44+P{curr_r}*0.96+Q{curr_r}*1.92+R{curr_r}*1.8+S{curr_r}*1.5+T{curr_r}*1.8+U{curr_r}*1.8+V{curr_r}*1.92"
                        ws.cell(curr_r, 7).value = f"=SUM(H{curr_r}:V{curr_r})"

                        for c in range(8, ws.max_column + 1):
                            ws.cell(curr_r, c).value = None

                        for sp, qty in cust_data["specs"].items():
                            c_idx = match_column(sp, col_map)
                            if c_idx:
                                ws.cell(curr_r, c_idx).value = qty

                        ws.row_dimensions[curr_r].hidden = False
                    else:
                        ws.cell(curr_r, 1).value = None
                        ws.cell(curr_r, 2).value = None
                        ws.cell(curr_r, 3).value = None
                        ws.cell(curr_r, 4).value = None
                        ws.cell(curr_r, 5).value = None
                        ws.cell(curr_r, 6).value = f"=H{curr_r}*0.96+I{curr_r}*1.5+J{curr_r}+K{curr_r}*1.8+L{curr_r}*1.8+M{curr_r}*1.8+N{curr_r}*1.92+O{curr_r}*1.44+P{curr_r}*0.96+Q{curr_r}*1.92+R{curr_r}*1.8+S{curr_r}*1.5+T{curr_r}*1.8+U{curr_r}*1.8+V{curr_r}*1.92"
                        ws.cell(curr_r, 7).value = f"=SUM(H{curr_r}:V{curr_r})"
                        for c in range(8, ws.max_column + 1):
                            ws.cell(curr_r, c).value = None
                        ws.row_dimensions[curr_r].hidden = True

                ws.row_dimensions[21].hidden = False
                ws.cell(21, 2).value = "TỔNG"
                ws.cell(21, 6).value = "=SUM(F4:F20)"
                ws.cell(21, 7).value = "=SUM(G4:G20)"
                for c in range(8, 23):
                    cl = get_column_letter(c)
                    ws.cell(21, c).value = f"=SUM({cl}4:{cl}20)"

                ws.cell(24, 1).value = "Mã khách hàng"
                ws.cell(24, 2).value = "Stt"
                ws.cell(24, 3).value = "CỬA HÀNG"
                ws.cell(24, 4).value = "NHÓM"
                ws.cell(24, 5).value = "TH"
                ws.cell(24, 6).value = "TỔNG"

                sub_trm_row = None
                sub_tps_row = None
                for r in range(24, ws.max_row + 1):
                    for c in range(1, 5):
                        v = str(ws.cell(r, c).value or "").upper()
                        if "TỔNG TOROMA" in v and not sub_trm_row:
                            sub_trm_row = r
                        elif "TỔNG TOPISA" in v and not sub_tps_row:
                            sub_tps_row = r

                if not sub_trm_row: sub_trm_row = 34
                if not sub_tps_row: sub_tps_row = 43

                if confirmed_items:
                    ExcelWriter.apply_confirmed_sync(ws, report_type, confirmed_items, is_unified=False)

                all_dealers = rule_manager.get_dealers("TRM_TPS") if rule_manager else []
                trm_dealers = [d for d in all_dealers if d.get("thuong_hieu") == "TRM"]
                tps_dealers = [d for d in all_dealers if d.get("thuong_hieu") == "TPS"]

                # 1. Sync TRM dealers
                existing_trm = {}
                for r in range(25, sub_trm_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 3).value or "").strip()
                    if code or name:
                        existing_trm[(code, normalize_dealer_name(name))] = r

                for d in trm_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip()
                    d_name_norm = normalize_dealer_name(d_name)
                    found = any((d_code and d_code == k[0]) or (d_name_norm and d_name_norm == k[1]) for k in existing_trm)
                    if not found:
                        new_r = sub_trm_row
                        ExcelWriter.insert_dealer_row(ws, new_r)
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 22)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_trm_row - 24
                        ws.cell(new_r, 3).value = d.get("ten_khach", "")
                        ws.cell(new_r, 4).value = d.get("nhom", "")
                        ws.cell(new_r, 5).value = "TRM"
                        ws.cell(new_r, 6).value = f"=SUM(G{new_r}:U{new_r})"
                        sub_trm_row += 1
                        sub_tps_row += 1
                        existing_trm[(d_code, d_name_norm)] = new_r

                # 2. Sync TPS dealers
                existing_tps = {}
                tps_start_row = sub_trm_row + 2
                for r in range(tps_start_row, sub_tps_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 3).value or "").strip()
                    if code or name:
                        existing_tps[(code, normalize_dealer_name(name))] = r

                for d in tps_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip()
                    d_name_norm = normalize_dealer_name(d_name)
                    found = any((d_code and d_code == k[0]) or (d_name_norm and d_name_norm == k[1]) for k in existing_tps)
                    if not found:
                        new_r = sub_tps_row
                        ExcelWriter.insert_dealer_row(ws, new_r)
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 22)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_tps_row - tps_start_row + 1
                        ws.cell(new_r, 3).value = d.get("ten_khach", "")
                        ws.cell(new_r, 4).value = d.get("nhom", "")
                        ws.cell(new_r, 5).value = "TPS"
                        ws.cell(new_r, 6).value = f"=SUM(G{new_r}:U{new_r})"
                        sub_tps_row += 1
                        existing_tps[(d_code, d_name_norm)] = new_r

                for r in range(25, sub_trm_row):
                    ws.cell(r, 5).value = "TRM"
                    ws.cell(r, 6).value = f"=SUM(G{r}:U{r})"

                ws.cell(sub_trm_row, 6).value = f"=SUM(F25:F{sub_trm_row - 1})"
                for c in range(7, 22):
                    cl = get_column_letter(c)
                    ws.cell(sub_trm_row, c).value = f"=SUM({cl}25:{cl}{sub_trm_row - 1})"

                tps_start_row = sub_trm_row + 2
                for r in range(tps_start_row, sub_tps_row):
                    ws.cell(r, 5).value = "TPS"
                    ws.cell(r, 6).value = f"=SUM(G{r}:U{r})"

                ws.cell(sub_tps_row, 6).value = f"=SUM(F{tps_start_row}:F{sub_tps_row - 1})"
                for c in range(7, 22):
                    cl = get_column_letter(c)
                    ws.cell(sub_tps_row, c).value = f"=SUM({cl}{tps_start_row}:{cl}{sub_tps_row - 1})"

                row_day_total = sub_tps_row + 2
                ws.cell(row_day_total, 6).value = f"=F{sub_trm_row}+F{sub_tps_row}"
                for c in range(7, 22):
                    cl = get_column_letter(c)
                    ws.cell(row_day_total, c).value = f"={cl}{sub_trm_row}+{cl}{sub_tps_row}"

                target_col = ExcelWriter.find_day_column(ws, 24, 7, target_date)
                if not target_col:
                    target_col = ws.max_column + 1
                    ws.cell(24, target_col).value = day_header_short

                app_dealers = rule_manager.get_dealers("TRM_TPS") if rule_manager else []
                app_by_name_th = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "TRM")): d for d in app_dealers}
                app_by_code_th = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", "TRM")): d for d in app_dealers if d.get("ma_khach")}
                app_by_name_only = {normalize_dealer_name(d.get("ten_khach", "")): d for d in app_dealers}

                muc01_by_code_th = {(row["ma_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row for row in muc01_rows if row.get("ma_khach")}
                muc01_by_name_th = {(normalize_dealer_name(row["ten_khach"]), row["thuong_hieu"].strip().upper()): row for row in muc01_rows}

                for r in range(25, sub_tps_row):
                    if r in [sub_trm_row, sub_trm_row + 1]: continue
                    name = str(ws.cell(r, 3).value or "").strip()
                    raw_code = str(ws.cell(r, 1).value or "").strip()
                    th = str(ws.cell(r, 5).value or ("TRM" if r < sub_trm_row else "TPS")).strip().upper()
                    name_norm = normalize_dealer_name(name)

                    rule_dealer = None
                    if raw_code: rule_dealer = app_by_code_th.get((raw_code.upper(), th))
                    if not rule_dealer and name_norm: rule_dealer = app_by_name_th.get((name_norm, th))
                    if not rule_dealer and name_norm: rule_dealer = app_by_name_only.get(name_norm)

                    if confirmed_items is None and not raw_code and rule_dealer and rule_dealer.get("ma_khach"):
                        raw_code = rule_dealer["ma_khach"]
                        ws.cell(r, 1).value = raw_code

                    if rule_dealer:
                        if confirmed_items is None and not ws.cell(r, 4).value and rule_dealer.get("nhom"):
                            ws.cell(r, 4).value = rule_dealer["nhom"]
                        ws.cell(r, 5).value = th

                    matched_data = None
                    if raw_code: matched_data = muc01_by_code_th.get((raw_code.upper(), th))
                    if not matched_data and name_norm: matched_data = muc01_by_name_th.get((name_norm, th))

                    if matched_data:
                        ws.cell(r, target_col).value = matched_data["tong_m2"]
                    else:
                        ws.cell(r, target_col).value = None

            elif report_type == "TMR":
                ws = ExcelWriter.get_target_sheet(wb, report_type, target_date)

                if ws.cell(3, 1).value != "Mã khách hàng":
                    ws.insert_cols(1)
                    ws.cell(3, 1).value = "Mã khách hàng"
                    ws.cell(9, 1).value = "Mã khách hàng"

                if ws.cell(3, 5).value != "TH":
                    ws.insert_cols(5)
                    ws.cell(3, 5).value = "TH"
                    ws.cell(9, 5).value = "TH"

                try:
                    ws.cell(2, 1).value = "BÁO CÁO BÁN HÀNG NGÀY"
                    ws.cell(2, 7).value = f"{day_str} NHÓM D1 - GPP TOMIRA"
                except Exception:
                    pass

                ws.cell(3, 1).value = "Mã khách hàng"
                ws.cell(3, 2).value = "STT"
                ws.cell(3, 3).value = "CỬA HÀNG"
                ws.cell(3, 4).value = "NHÓM"
                ws.cell(3, 5).value = "TH"
                ws.cell(3, 6).value = "TỔNG (mét)"
                ws.cell(3, 7).value = "Tổng (thùng)"

                col_map = {}
                for c in range(8, ws.max_column + 1):
                    val = ws.cell(3, c).value
                    if val:
                        col_map[str(val).replace('\n', ' ').strip()] = c

                start_r = 4
                end_r = 6
                num_active = len(muc01_rows)
                for i in range(end_r - start_r + 1):
                    curr_r = start_r + i
                    if i < num_active:
                        cust_data = muc01_rows[i]
                        ws.cell(curr_r, 1).value = cust_data.get("ma_khach", "")
                        ws.cell(curr_r, 2).value = i + 1
                        ws.cell(curr_r, 3).value = cust_data.get("ten_khach", "")
                        ws.cell(curr_r, 4).value = cust_data.get("nhom", "")
                        ws.cell(curr_r, 5).value = "TMR"

                        ws.cell(curr_r, 6).value = f"=H{curr_r}+(I{curr_r}*1.5)+J{curr_r}+(K{curr_r}*1.62)+(L{curr_r}*1.8)+(M{curr_r}*1.44)+(N{curr_r}*1.8)+(O{curr_r}*1.44)+(P{curr_r}*1.8)+(Q{curr_r}*1.92)+(R{curr_r}*1.44)+(S{curr_r}*1.44)+(T{curr_r}*1.92)+(U{curr_r}*1.5)+(V{curr_r}*1.8)+(W{curr_r}*1.8)+(X{curr_r}*1.92)"
                        ws.cell(curr_r, 7).value = f"=SUM(H{curr_r}:X{curr_r})"

                        for c in range(8, ws.max_column + 1):
                            ws.cell(curr_r, c).value = None

                        for sp, qty in cust_data["specs"].items():
                            c_idx = match_column(sp, col_map)
                            if c_idx:
                                ws.cell(curr_r, c_idx).value = qty

                        ws.row_dimensions[curr_r].hidden = False
                    else:
                        ws.cell(curr_r, 1).value = None
                        ws.cell(curr_r, 2).value = None
                        ws.cell(curr_r, 3).value = None
                        ws.cell(curr_r, 4).value = None
                        ws.cell(curr_r, 5).value = None
                        ws.cell(curr_r, 6).value = f"=H{curr_r}+(I{curr_r}*1.5)+J{curr_r}+(K{curr_r}*1.62)+(L{curr_r}*1.8)+(M{curr_r}*1.44)+(N{curr_r}*1.8)+(O{curr_r}*1.44)+(P{curr_r}*1.8)+(Q{curr_r}*1.92)+(R{curr_r}*1.44)+(S{curr_r}*1.44)+(T{curr_r}*1.92)+(U{curr_r}*1.5)+(V{curr_r}*1.8)+(W{curr_r}*1.8)+(X{curr_r}*1.92)"
                        ws.cell(curr_r, 7).value = f"=SUM(H{curr_r}:X{curr_r})"
                        for c in range(8, ws.max_column + 1):
                            ws.cell(curr_r, c).value = None
                        ws.row_dimensions[curr_r].hidden = True

                ws.row_dimensions[7].hidden = False
                ws.cell(7, 1).value = "TỔNG"
                ws.cell(7, 6).value = "=SUM(F4:F6)"
                ws.cell(7, 7).value = "=SUM(G4:G6)"
                for c in range(8, 25):
                    cl = get_column_letter(c)
                    ws.cell(7, c).value = f"=SUM({cl}4:{cl}6)"

                ws.cell(9, 1).value = "Mã khách hàng"
                ws.cell(9, 2).value = "STT"
                ws.cell(9, 3).value = "CỬA HÀNG"
                ws.cell(9, 4).value = "NHÓM"
                ws.cell(9, 5).value = "TH"
                ws.cell(9, 6).value = "TỔNG\n(m²)"

                sub_tmr_row = None
                for r in range(10, ws.max_row + 1):
                    for c in range(1, 4):
                        v = str(ws.cell(r, c).value or "").upper()
                        if "TỔNG (M²)" in v or "TỔNG" in v:
                            sub_tmr_row = r
                            break
                    if sub_tmr_row: break
                if not sub_tmr_row: sub_tmr_row = 13

                if confirmed_items:
                    ExcelWriter.apply_confirmed_sync(ws, report_type, confirmed_items, is_unified=False)

                tmr_dealers = rule_manager.get_dealers("TMR", brand="TMR") if rule_manager else []
                existing_tmr = {}
                for r in range(10, sub_tmr_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 3).value or "").strip()
                    if code or name:
                        existing_tmr[(code, normalize_dealer_name(name))] = r

                for d in tmr_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip()
                    d_name_norm = normalize_dealer_name(d_name)
                    found = any((d_code and d_code == k[0]) or (d_name_norm and d_name_norm == k[1]) for k in existing_tmr)
                    if not found:
                        new_r = sub_tmr_row
                        ExcelWriter.insert_dealer_row(ws, new_r)
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 25)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_tmr_row - 9
                        ws.cell(new_r, 3).value = d.get("ten_khach", "")
                        ws.cell(new_r, 4).value = d.get("nhom", "")
                        ws.cell(new_r, 5).value = "TMR"
                        ws.cell(new_r, 6).value = f"=SUM(G{new_r}:X{new_r})"
                        sub_tmr_row += 1
                        existing_tmr[(d_code, d_name_norm)] = new_r

                target_col = ExcelWriter.find_day_column(ws, 9, 7, target_date)
                if not target_col:
                    target_col = ws.max_column + 1
                    ws.cell(9, target_col).value = day_header_short

                app_dealers = rule_manager.get_dealers("TMR", brand="TMR") if rule_manager else []
                app_by_name = {normalize_dealer_name(d.get("ten_khach", "")): d for d in app_dealers}
                app_by_code = {str(d.get("ma_khach") or "").strip().upper(): d for d in app_dealers if d.get("ma_khach")}

                muc01_by_code_th = {(row["ma_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row for row in muc01_rows if row.get("ma_khach")}
                muc01_by_name_th = {(normalize_dealer_name(row["ten_khach"]), row["thuong_hieu"].strip().upper()): row for row in muc01_rows}

                for r in range(10, sub_tmr_row):
                    code = str(ws.cell(r, 1).value or "").strip()
                    name = str(ws.cell(r, 3).value or "").strip()
                    name_norm = normalize_dealer_name(name)

                    rule_dealer = None
                    if code: rule_dealer = app_by_code.get(code.upper())
                    if not rule_dealer and name_norm: rule_dealer = app_by_name.get(name_norm)

                    if confirmed_items is None and not code and rule_dealer and rule_dealer.get("ma_khach"):
                        code = rule_dealer["ma_khach"]
                        ws.cell(r, 1).value = code

                    if rule_dealer:
                        if confirmed_items is None and not ws.cell(r, 4).value and rule_dealer.get("nhom"):
                            ws.cell(r, 4).value = rule_dealer["nhom"]
                        ws.cell(r, 5).value = "TMR"

                    matched_data = None
                    if code: matched_data = muc01_by_code_th.get((code.upper(), "TMR"))
                    if not matched_data and name_norm: matched_data = muc01_by_name_th.get((name_norm, "TMR"))

                    if matched_data:
                        ws.cell(r, target_col).value = matched_data["tong_m2"]
                    else:
                        ws.cell(r, target_col).value = None

                ws.cell(sub_tmr_row, 1).value = "TỔNG (m²)"
                ws.cell(sub_tmr_row, 6).value = f"=SUM(F10:F{sub_tmr_row - 1})"
                cl_t = get_column_letter(target_col)
                ws.cell(sub_tmr_row, target_col).value = f"=SUM({cl_t}10:{cl_t}{sub_tmr_row - 1})"

        # Kích hoạt sheet vừa cập nhật để khi người dùng mở Excel sẽ vào đúng sheet này
        ws_name = ws.title
        wb.active = wb.index(ws)

        # Lưu an toàn vào file tạm trong cùng thư mục và xác thực trước khi replace
        file_dir = os.path.dirname(os.path.abspath(template_path))
        fd, temporary_path = tempfile.mkstemp(prefix=".vtc-", suffix=".xlsx", dir=file_dir)
        os.close(fd)

        try:
            wb.save(temporary_path)
            wb.close()

            # Đóng rồi mở lại file từ đĩa để xác thực dữ liệu đã lưu
            saved = openpyxl.load_workbook(temporary_path, read_only=True)
            try:
                if ws_name not in saved.sheetnames:
                    raise ValueError(f"Không tìm thấy sheet '{ws_name}' trong tệp đã lưu.")
                saved_ws = saved[ws_name]
                if saved.active.title != ws_name:
                    pass
            finally:
                saved.close()

            os.replace(temporary_path, template_path)
        except PermissionError as exc:
            raise PermissionError(f"Tệp '{template_path}' đang được mở bởi ứng dụng khác (ví dụ Excel). Vui lòng đóng tệp rồi thử lại!") from exc
        finally:
            if os.path.exists(temporary_path):
                try:
                    os.unlink(temporary_path)
                except Exception:
                    pass

        return template_path
