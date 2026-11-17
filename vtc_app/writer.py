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
    3. So khớp theo chuỗi dài nhất chứa nhau (ưu tiên tên đặc thù dài hơn, ví dụ '40x40 (12v/t)' trước '40x40').
    """
    sp_norm = sp.strip().lower()
    # 1. Exact match
    for sp_name, col_idx in col_map.items():
        if sp_norm == sp_name.strip().lower():
            return col_idx
            
    # 2. Normalized match without whitespaces
    sp_clean = re.sub(r'\s+', '', sp_norm)
    for sp_name, col_idx in col_map.items():
        col_clean = re.sub(r'\s+', '', sp_name.strip().lower())
        if sp_clean == col_clean:
            return col_idx
            
    # 3. Longest substring match (sorted by length descending)
    sorted_cols = sorted(col_map.items(), key=lambda x: len(x[0]), reverse=True)
    for sp_name, col_idx in sorted_cols:
        col_norm = sp_name.strip().lower()
        if sp_norm == col_norm or sp_norm in col_norm or col_norm in sp_norm:
            return col_idx
            
    return None

class ExcelWriter:
    @staticmethod
    def create_daily_backup(file_path: str, backup_dir: str, current_date: Optional[date] = None) -> str:
        """
        Tạo gói sao lưu mang ngày hôm trước nếu hôm nay chưa có sao lưu.
        """
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
    def copy_row_format(ws, src_row: int, dst_row: int, max_col: int = 25):
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
    def analyze_master_excel_dealers(
        template_path: str,
        report_type: str,
        rule_manager: Any
    ) -> List[Dict[str, Any]]:
        """
        Phân tích đối chiếu Mục 02 của tệp Excel tổng với danh sách đại lý đã xác minh trong AppData.
        Phát hiện:
        - Các dòng thiếu mã khách hàng (cần điền từ AppData).
        - Các dòng khác mã khách hàng (xung đột mã giữa Excel và AppData).
        - Các dòng nhóm khác nhau (cần chuẩn hóa từ số 1, 2 sang VIP, D1, Gpp).
        - Riêng mã CH9999 của Bảo Việt: gắn cờ xác minh trước khi ghi.
        """
        if not os.path.exists(template_path):
            return []

        wb = openpyxl.load_workbook(template_path, data_only=True)
        ws = wb.active
        items = []

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
            if not sub_tcr_row:
                sub_tcr_row = 42
            if not sub_tml_row:
                sub_tml_row = 52

            app_dealers = rule_manager.get_dealers("TCR_TML") if rule_manager else []
            app_by_name = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "TCR")): d for d in app_dealers}
            app_by_code = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", "TCR")): d for d in app_dealers if d.get("ma_khach")}

            for r in range(29, sub_tml_row):
                if r == sub_tcr_row:
                    continue
                th = "TCR" if r < sub_tcr_row else "TML"
                ex_code = str(ws.cell(r, 1).value or "").strip()
                if ex_code.upper() in ["NONE", "MÃ KHÁCH HÀNG", "MÃ KH"]:
                    ex_code = ""
                ex_name = str(ws.cell(r, 4).value or "").strip()
                ex_grp = str(ws.cell(r, 5).value or "").strip()
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
            if not sub_trm_row:
                sub_trm_row = 34
            if not sub_tps_row:
                sub_tps_row = 43

            app_dealers = rule_manager.get_dealers("TRM_TPS") if rule_manager else []
            app_by_name = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "TRM")): d for d in app_dealers}
            app_by_code = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", "TRM")): d for d in app_dealers if d.get("ma_khach")}

            for r in range(25, sub_tps_row):
                if r in [sub_trm_row, sub_trm_row + 1]:
                    continue
                th = "TRM" if r < sub_trm_row else "TPS"
                ex_code = str(ws.cell(r, 1).value or "").strip()
                if ex_code.upper() in ["NONE", "MÃ KHÁCH HÀNG", "MÃ KH"]:
                    ex_code = ""
                ex_name = str(ws.cell(r, 3).value or "").strip()
                ex_grp = str(ws.cell(r, 4).value or "").strip()
                if not ex_name or "CỬA HÀNG" in ex_name.upper():
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

        elif report_type == "TMR":
            sub_tmr_row = None
            for r in range(10, ws.max_row + 1):
                for c in range(1, 4):
                    v = str(ws.cell(r, c).value or "").upper()
                    if "TỔNG (M²)" in v or "TỔNG" in v:
                        sub_tmr_row = r
                        break
                if sub_tmr_row:
                    break
            if not sub_tmr_row:
                sub_tmr_row = 13

            app_dealers = rule_manager.get_dealers("TMR", brand="TMR") if rule_manager else []
            app_by_name = {normalize_dealer_name(d.get("ten_khach", "")): d for d in app_dealers}
            app_by_code = {str(d.get("ma_khach") or "").strip().upper(): d for d in app_dealers if d.get("ma_khach")}

            has_code_col = (ws.cell(9, 1).value == "Mã khách hàng" or ws.cell(3, 1).value == "Mã khách hàng")

            for r in range(10, sub_tmr_row):
                if has_code_col:
                    ex_code = str(ws.cell(r, 1).value or "").strip()
                    if ex_code.upper() in ["NONE", "MÃ KHÁCH HÀNG", "MÃ KH"]:
                        ex_code = ""
                    ex_name = str(ws.cell(r, 3).value or "").strip()
                    ex_grp = str(ws.cell(r, 4).value or "").strip()
                else:
                    ex_code = ""
                    ex_name = str(ws.cell(r, 2).value or "").strip()
                    ex_grp = str(ws.cell(r, 3).value or "").strip()

                if not ex_name or "CỬA HÀNG" in ex_name.upper():
                    continue

                matched = None
                if ex_code:
                    matched = app_by_code.get(ex_code.upper())
                if not matched and ex_name:
                    matched = app_by_name.get(normalize_dealer_name(ex_name))

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
                    "th": "TMR",
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

    @staticmethod
    def apply_confirmed_sync_to_sheet(
        template_path: str,
        report_type: str,
        confirmed_items: List[Dict[str, Any]]
    ):
        """
        Áp dụng các mã khách hàng và phân nhóm đã được người dùng xác nhận vào Mục 02 của Excel tổng.
        Bảo toàn 100% công thức, định dạng và số liệu các ngày cũ.
        """
        if not os.path.exists(template_path) or not confirmed_items:
            return

        wb = openpyxl.load_workbook(template_path, data_only=False)
        ws = wb.active

        if report_type == "TCR_TML":
            ws.cell(27, 1).value = "Mã khách hàng"
            for it in confirmed_items:
                r = it.get("row_num")
                if r:
                    code_to_write = it.get("final_code", it.get("app_code", ""))
                    grp_to_write = it.get("final_group", it.get("app_group", ""))
                    th_to_write = it.get("th", "TCR")
                    if code_to_write:
                        ws.cell(r, 1).value = code_to_write
                    if grp_to_write:
                        ws.cell(r, 5).value = grp_to_write
                    if th_to_write:
                        ws.cell(r, 6).value = th_to_write

        elif report_type == "TRM_TPS":
            ws.cell(24, 1).value = "Mã khách hàng"
            for it in confirmed_items:
                r = it.get("row_num")
                if r:
                    code_to_write = it.get("final_code", it.get("app_code", ""))
                    grp_to_write = it.get("final_group", it.get("app_group", ""))
                    th_to_write = it.get("th", "TRM")
                    if code_to_write:
                        ws.cell(r, 1).value = code_to_write
                    if grp_to_write:
                        ws.cell(r, 4).value = grp_to_write
                    if th_to_write:
                        ws.cell(r, 5).value = th_to_write

        elif report_type == "TMR":
            # Đảm bảo cột 1 là Mã khách hàng nếu chưa có
            if ws.cell(9, 1).value != "Mã khách hàng" and ws.cell(3, 1).value != "Mã khách hàng":
                ws.insert_cols(1)
                ws.cell(3, 1).value = "Mã khách hàng"
                ws.cell(9, 1).value = "Mã khách hàng"
            if ws.cell(3, 5).value != "TH" and ws.cell(9, 5).value != "TH":
                ws.insert_cols(5)
                ws.cell(3, 5).value = "TH"
                ws.cell(9, 5).value = "TH"

            for it in confirmed_items:
                r = it.get("row_num")
                if r:
                    code_to_write = it.get("final_code", it.get("app_code", ""))
                    grp_to_write = it.get("final_group", it.get("app_group", ""))
                    if code_to_write:
                        ws.cell(r, 1).value = code_to_write
                    if grp_to_write:
                        ws.cell(r, 4).value = grp_to_write
                    ws.cell(r, 5).value = "TMR"

        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, f"vtc_sync_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        wb.save(temp_file)
        wb.close()

        try:
            shutil.move(temp_file, template_path)
        except PermissionError:
            raise PermissionError(f"Tệp '{template_path}' đang được mở bởi ứng dụng khác (ví dụ Excel). Vui lòng đóng tệp rồi thử lại!")

    @staticmethod
    def update_summary_excel(
        template_path: str,
        report_type: str,
        app_result: Dict[str, Any],
        target_date: date,
        backup_dir: str,
        rule_manager: Optional[Any] = None
    ) -> str:
        """
        Cập nhật kết quả vào tệp Excel tổng, giữ nguyên định dạng, công thức và bố cục.
        - Mục 01: Có mã khách và CHỈ hiện đại lý có phát sinh hợp lệ.
        - Mục 02: Đồng bộ danh sách theo dõi, tự động chèn dòng trước dòng tổng khi có đại lý mới, mở rộng vùng công thức.
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Không tìm thấy tệp Excel tổng: {template_path}")

        # 1. Tạo bản sao lưu trước khi ghi
        ExcelWriter.create_daily_backup(template_path, backup_dir, target_date)

        # 2. Mở workbook với data_only=False để giữ nguyên công thức
        wb = openpyxl.load_workbook(template_path, data_only=False)

        day_str = target_date.strftime("%d/%m/%Y")
        sheet_day_str = target_date.strftime("%d.%m.%Y")
        month_str = target_date.strftime("thang %m.%Y")
        month_code = target_date.strftime("T%m.%Y")
        day_header_short = target_date.strftime("%d/%m")

        muc01_rows = app_result.get("muc01_rows", [])

        if report_type == "TCR_TML":
            # Sheet theo ngày
            if sheet_day_str in wb.sheetnames:
                ws = wb[sheet_day_str]
            else:
                source_sheet = wb.active
                ws = wb.copy_worksheet(source_sheet)
                ws.title = sheet_day_str

            # Tiêu đề ngày dòng 1
            ws.cell(1, 2).value = f"BÁO CÁO BÁN HÀNG NGÀY {day_str} NHÓM VIP - D1 - GPP TOCERA VÀ TOMILAN"

            # Headers Mục 01 ở dòng 2
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

            # Mục 01: Dòng 3 đến 24
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

                    # Công thức m2 và thùng dòng phát sinh
                    ws.cell(curr_r, 7).value = f"=I{curr_r}+J{curr_r}*1.5+K{curr_r}+L{curr_r}*1.62+M{curr_r}*1.8+N{curr_r}*1.44+O{curr_r}*1.44+P{curr_r}*1.92+Q{curr_r}*0.99+R{curr_r}*1.44+S{curr_r}*0.96+T{curr_r}*1.92+U{curr_r}*1.92+V{curr_r}*1.5+W{curr_r}*1.8+X{curr_r}*1.8+Y{curr_r}*1.92"
                    ws.cell(curr_r, 8).value = f"=SUM(I{curr_r}:Y{curr_r})"

                    # Xóa các cột specs cũ
                    for c in range(9, ws.max_column + 1):
                        ws.cell(curr_r, c).value = None

                    # Ghi specs mới
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

            # Dòng 25: TỔNG Mục 01
            ws.row_dimensions[25].hidden = False
            ws.cell(25, 2).value = "TỔNG"
            ws.cell(25, 7).value = "=SUM(G3:G24)"
            ws.cell(25, 8).value = "=SUM(H3:H24)"
            for c in range(9, 26):
                cl = get_column_letter(c)
                ws.cell(25, c).value = f"=SUM({cl}3:{cl}24)"

            # Headers Mục 02 ở dòng 27
            ws.cell(27, 1).value = "Mã khách hàng"
            ws.cell(27, 2).value = "STT"
            ws.cell(27, 4).value = "CỬA HÀNG"
            ws.cell(27, 5).value = "NHÓM"
            ws.cell(27, 6).value = "TH"
            ws.cell(27, 7).value = "∑SL MÉT"

            # Dynamic Dealer Sync for TCR_TML if rule_manager provided
            sub_tcr_row = None
            sub_tml_row = None
            for r in range(28, ws.max_row + 1):
                for c in range(1, 5):
                    v = str(ws.cell(r, c).value or "").upper()
                    if "TỔNG THƯƠNG HIỆU TOCERA" in v and not sub_tcr_row:
                        sub_tcr_row = r
                    elif "TỔNG THƯƠNG HIỆU TOMILAN" in v and not sub_tml_row:
                        sub_tml_row = r

            if not sub_tcr_row:
                sub_tcr_row = 42
            if not sub_tml_row:
                sub_tml_row = 52

            if rule_manager:
                all_dealers = rule_manager.get_dealers("TCR_TML")
                tcr_dealers = [d for d in all_dealers if d.get("thuong_hieu") == "TCR"]
                tml_dealers = [d for d in all_dealers if d.get("thuong_hieu") == "TML"]

                # 1. Sync TCR dealers
                existing_tcr = {}
                for r in range(29, sub_tcr_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 4).value or "").strip().upper()
                    if code or name:
                        existing_tcr[(code, name)] = r

                for d in tcr_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip().upper()
                    found = any((d_code and d_code == k[0]) or (d_name and d_name == k[1]) for k in existing_tcr)
                    if not found:
                        ws.insert_rows(sub_tcr_row, 1)
                        new_r = sub_tcr_row
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 26)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_tcr_row - 28
                        ws.cell(new_r, 4).value = d.get("ten_khach", "")
                        ws.cell(new_r, 5).value = d.get("nhom", "")
                        ws.cell(new_r, 6).value = "TCR"
                        sub_tcr_row += 1
                        sub_tml_row += 1
                        existing_tcr[(d_code, d_name)] = new_r

                # 2. Sync TML dealers
                existing_tml = {}
                tml_start_row = sub_tcr_row + 1
                for r in range(tml_start_row, sub_tml_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 4).value or "").strip().upper()
                    if code or name:
                        existing_tml[(code, name)] = r

                for d in tml_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip().upper()
                    found = any((d_code and d_code == k[0]) or (d_name and d_name == k[1]) for k in existing_tml)
                    if not found:
                        ws.insert_rows(sub_tml_row, 1)
                        new_r = sub_tml_row
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 26)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_tml_row - tml_start_row + 1
                        ws.cell(new_r, 4).value = d.get("ten_khach", "")
                        ws.cell(new_r, 5).value = d.get("nhom", "")
                        ws.cell(new_r, 6).value = "TML"
                        sub_tml_row += 1
                        existing_tml[(d_code, d_name)] = new_r

            # Tìm cột ngày tương ứng trong Mục 02 (từ Col 8 trở đi)
            target_col = None
            for c in range(8, ws.max_column + 1):
                cell_val = ws.cell(27, c).value
                if isinstance(cell_val, datetime) and cell_val.date() == target_date:
                    target_col = c
                    break
                elif isinstance(cell_val, str) and day_header_short in cell_val:
                    target_col = c
                    break

            if target_col:
                muc01_by_code_th = {
                    (row["ma_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row
                    for row in muc01_rows if row.get("ma_khach")
                }
                muc01_by_name_th = {
                    (row["ten_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row
                    for row in muc01_rows
                }

                # Cập nhật Mục 02 cho TCR và TML
                for r in range(29, sub_tml_row):
                    if r == sub_tcr_row:
                        continue
                    name = ws.cell(r, 4).value
                    raw_code = ws.cell(r, 1).value
                    th = ws.cell(r, 6).value or ("TCR" if r < sub_tcr_row else "TML")
                    th = str(th).strip().upper()

                    matched_data = None
                    if raw_code and str(raw_code).strip() != "None":
                        code_clean = str(raw_code).strip().upper()
                        matched_data = muc01_by_code_th.get((code_clean, th))

                    if not matched_data and name:
                        name_clean = str(name).strip().upper()
                        matched_data = muc01_by_name_th.get((name_clean, th))

                    if matched_data:
                        if not raw_code or str(raw_code).strip() == "None":
                            ws.cell(r, 1).value = matched_data.get("ma_khach", "")
                        ws.cell(r, target_col).value = matched_data["tong_m2"]
                    else:
                        ws.cell(r, target_col).value = None

                # Công thức tổng thương hiệu và tổng ngày ở Mục 02
                cl_t = get_column_letter(target_col)
                # TỔNG THƯƠNG HIỆU TOCERA
                ws.cell(sub_tcr_row, target_col).value = f"=SUM({cl_t}29:{cl_t}{sub_tcr_row - 1})"
                # TỔNG THƯƠNG HIỆU TOMILAN
                ws.cell(sub_tml_row, target_col).value = f"=SUM({cl_t}{sub_tcr_row + 1}:{cl_t}{sub_tml_row - 1})"
                # TỔNG M2 (Row sub_tml_row + 1)
                row_tot_m2 = sub_tml_row + 1
                ws.cell(row_tot_m2, target_col).value = f"=SUM({cl_t}{sub_tcr_row},{cl_t}{sub_tml_row})"

        elif report_type == "TMR":
            sheet_title = None
            for s in wb.sheetnames:
                if month_code.lower() in s.lower() or target_date.strftime("%m.%Y") in s:
                    sheet_title = s
                    break
            if not sheet_title:
                sheet_title = wb.sheetnames[-1]

            ws = wb[sheet_title]

            # Kiểm tra nếu chưa có cột Mã khách hàng tại cột 1
            if ws.cell(3, 1).value != "Mã khách hàng":
                ws.insert_cols(1)
                ws.cell(3, 1).value = "Mã khách hàng"
                ws.cell(9, 1).value = "Mã khách hàng"

            # Kiểm tra nếu chưa có cột TH tại cột 5 (ngay sau NHÓM)
            if ws.cell(3, 5).value != "TH":
                ws.insert_cols(5)
                ws.cell(3, 5).value = "TH"
                ws.cell(9, 5).value = "TH"

            # Dòng 2 tiêu đề
            try:
                ws.cell(2, 7).value = f"{day_str} NHÓM D1 - Gpp TOMIRA"
            except Exception:
                pass

            # Headers Mục 01 ở dòng 3
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

            # Mục 01: Dòng 4 đến 6
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

                    # Công thức m2 và thùng
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

            # Dòng 7: TỔNG Mục 01
            ws.row_dimensions[7].hidden = False
            ws.cell(7, 1).value = "TỔNG"
            ws.cell(7, 6).value = "=SUM(F4:F6)"
            ws.cell(7, 7).value = "=SUM(G4:G6)"
            for c in range(8, 25):
                cl = get_column_letter(c)
                ws.cell(7, c).value = f"=SUM({cl}4:{cl}6)"

            # Headers Mục 02 ở dòng 9
            ws.cell(9, 1).value = "Mã khách hàng"
            ws.cell(9, 2).value = "STT"
            ws.cell(9, 3).value = "CỬA HÀNG"
            ws.cell(9, 4).value = "NHÓM"
            ws.cell(9, 5).value = "TH"
            ws.cell(9, 6).value = "TỔNG\n(m²)"

            # Dynamic Dealer Sync for TMR
            sub_tmr_row = None
            for r in range(10, ws.max_row + 1):
                for c in range(1, 4):
                    v = str(ws.cell(r, c).value or "").upper()
                    if "TỔNG (M²)" in v or "TỔNG" in v:
                        sub_tmr_row = r
                        break
                if sub_tmr_row:
                    break
            if not sub_tmr_row:
                sub_tmr_row = 13

            if rule_manager:
                tmr_dealers = rule_manager.get_dealers("TMR", brand="TMR")
                existing_tmr = {}
                for r in range(10, sub_tmr_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 3).value or "").strip().upper()
                    if code or name:
                        existing_tmr[(code, name)] = r

                for d in tmr_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip().upper()
                    found = any((d_code and d_code == k[0]) or (d_name and d_name == k[1]) for k in existing_tmr)
                    if not found:
                        ws.insert_rows(sub_tmr_row, 1)
                        new_r = sub_tmr_row
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 25)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_tmr_row - 9
                        ws.cell(new_r, 3).value = d.get("ten_khach", "")
                        ws.cell(new_r, 4).value = d.get("nhom", "")
                        ws.cell(new_r, 5).value = "TMR"
                        ws.cell(new_r, 6).value = f"=SUM(G{new_r}:X{new_r})"
                        sub_tmr_row += 1
                        existing_tmr[(d_code, d_name)] = new_r

            # Tìm cột ngày tương ứng trong Mục 02 (từ Col 7 trở đi)
            target_col = None
            for c in range(7, ws.max_column + 1):
                val = ws.cell(9, c).value
                if val and day_header_short in str(val):
                    target_col = c
                    break

            if target_col:
                muc01_by_code_th = {
                    (row["ma_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row
                    for row in muc01_rows if row.get("ma_khach")
                }
                muc01_by_name_th = {
                    (row["ten_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row
                    for row in muc01_rows
                }

                for r in range(10, sub_tmr_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 3).value or "").strip().upper()

                    matched_data = None
                    if code:
                        matched_data = muc01_by_code_th.get((code, "TMR"))
                    if not matched_data and name:
                        matched_data = muc01_by_name_th.get((name, "TMR"))

                    if matched_data:
                        ws.cell(r, target_col).value = matched_data["tong_m2"]
                    else:
                        ws.cell(r, target_col).value = None

                # Subtotal row
                ws.cell(sub_tmr_row, 1).value = "TỔNG (m²)"
                ws.cell(sub_tmr_row, 6).value = f"=SUM(F10:F{sub_tmr_row - 1})"
                cl_t = get_column_letter(target_col)
                ws.cell(sub_tmr_row, target_col).value = f"=SUM({cl_t}10:{cl_t}{sub_tmr_row - 1})"

        elif report_type == "TRM_TPS":
            sheet_title = None
            for s in wb.sheetnames:
                if month_code in s or f"T{target_date.month}" in s:
                    sheet_title = s
                    break
            if not sheet_title:
                sheet_title = wb.sheetnames[0]

            ws = wb[sheet_title]

            # Kiểm tra nếu chưa có cột TH (tại cột 5)
            if ws.cell(3, 5).value != "TH":
                ws.insert_cols(5)
                for r in range(1, 55):
                    c4 = ws.cell(r, 4)
                    c5 = ws.cell(r, 5)
                    if c4.has_style:
                        c5.font = copy(c4.font)
                        c5.border = copy(c4.border)
                        c5.fill = copy(c4.fill)
                        c5.number_format = c4.number_format
                        c5.alignment = Alignment(horizontal="center", vertical="center")

            # 1. Headers Mục 01 (Dòng 3)
            ws.cell(3, 1).value = "Mã khách hàng"
            ws.cell(3, 2).value = "Stt"
            ws.cell(3, 3).value = "CỬA HÀNG"
            ws.cell(3, 4).value = "NHÓM"
            ws.cell(3, 5).value = "TH"
            ws.cell(3, 6).value = "TỔNG (m²)"
            ws.cell(3, 7).value = "TỔNG (thùng)"

            try:
                if not isinstance(ws.cell(2, 15), openpyxl.cell.cell.MergedCell):
                    ws.cell(2, 15).value = f" {day_str}"
                elif not isinstance(ws.cell(2, 16), openpyxl.cell.cell.MergedCell):
                    ws.cell(2, 16).value = f" {day_str}"
            except Exception:
                pass

            col_map = {}
            for c in range(8, ws.max_column + 1):
                val = ws.cell(3, c).value
                if val:
                    col_map[str(val).replace('\n', ' ').strip()] = c

            # 2. Mục 01: Dòng 4 đến 20
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

                    # Công thức m2 và thùng
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

            # 3. Dòng 21: TỔNG Mục 01
            ws.row_dimensions[21].hidden = False
            ws.cell(21, 2).value = "TỔNG"
            ws.cell(21, 6).value = "=SUM(F4:F20)"
            ws.cell(21, 7).value = "=SUM(G4:G20)"
            for c in range(8, 23):
                cl = get_column_letter(c)
                ws.cell(21, c).value = f"=SUM({cl}4:{cl}20)"

            # 4. Headers Mục 02 (Dòng 24)
            ws.cell(24, 1).value = "Mã khách hàng"
            ws.cell(24, 2).value = "Stt"
            ws.cell(24, 3).value = "CỬA HÀNG"
            ws.cell(24, 4).value = "NHÓM"
            ws.cell(24, 5).value = "TH"
            ws.cell(24, 6).value = "TỔNG"

            # Dynamic Dealer Sync for TRM_TPS
            sub_trm_row = None
            sub_tps_row = None
            for r in range(24, ws.max_row + 1):
                for c in range(1, 5):
                    v = str(ws.cell(r, c).value or "").upper()
                    if "TỔNG TOROMA" in v and not sub_trm_row:
                        sub_trm_row = r
                    elif "TỔNG TOPISA" in v and not sub_tps_row:
                        sub_tps_row = r

            if not sub_trm_row:
                sub_trm_row = 34
            if not sub_tps_row:
                sub_tps_row = 43

            if rule_manager:
                all_dealers = rule_manager.get_dealers("TRM_TPS")
                trm_dealers = [d for d in all_dealers if d.get("thuong_hieu") == "TRM"]
                tps_dealers = [d for d in all_dealers if d.get("thuong_hieu") == "TPS"]

                # 1. Sync TRM dealers
                existing_trm = {}
                for r in range(25, sub_trm_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 3).value or "").strip().upper()
                    if code or name:
                        existing_trm[(code, name)] = r

                for d in trm_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip().upper()
                    found = any((d_code and d_code == k[0]) or (d_name and d_name == k[1]) for k in existing_trm)
                    if not found:
                        ws.insert_rows(sub_trm_row, 1)
                        new_r = sub_trm_row
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 22)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_trm_row - 24
                        ws.cell(new_r, 3).value = d.get("ten_khach", "")
                        ws.cell(new_r, 4).value = d.get("nhom", "")
                        ws.cell(new_r, 5).value = "TRM"
                        ws.cell(new_r, 6).value = f"=SUM(G{new_r}:U{new_r})"
                        sub_trm_row += 1
                        sub_tps_row += 1
                        existing_trm[(d_code, d_name)] = new_r

                # 2. Sync TPS dealers
                existing_tps = {}
                tps_start_row = sub_trm_row + 2
                for r in range(tps_start_row, sub_tps_row):
                    code = str(ws.cell(r, 1).value or "").strip().upper()
                    name = str(ws.cell(r, 3).value or "").strip().upper()
                    if code or name:
                        existing_tps[(code, name)] = r

                for d in tps_dealers:
                    d_code = str(d.get("ma_khach") or "").strip().upper()
                    d_name = str(d.get("ten_khach") or "").strip().upper()
                    found = any((d_code and d_code == k[0]) or (d_name and d_name == k[1]) for k in existing_tps)
                    if not found:
                        ws.insert_rows(sub_tps_row, 1)
                        new_r = sub_tps_row
                        ExcelWriter.copy_row_format(ws, new_r - 1, new_r, 22)
                        ws.cell(new_r, 1).value = d.get("ma_khach", "")
                        ws.cell(new_r, 2).value = sub_tps_row - tps_start_row + 1
                        ws.cell(new_r, 3).value = d.get("ten_khach", "")
                        ws.cell(new_r, 4).value = d.get("nhom", "")
                        ws.cell(new_r, 5).value = "TPS"
                        ws.cell(new_r, 6).value = f"=SUM(G{new_r}:U{new_r})"
                        sub_tps_row += 1
                        existing_tps[(d_code, d_name)] = new_r

            # Cập nhật công thức dòng Mục 02
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

            # Tìm cột ngày tương ứng trong Mục 02 (từ Col 7 trở đi)
            target_col = None
            for c in range(7, ws.max_column + 1):
                val = ws.cell(24, c).value
                if val and day_header_short in str(val):
                    target_col = c
                    break

            if target_col:
                muc01_by_code_th = {
                    (row["ma_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row
                    for row in muc01_rows if row.get("ma_khach")
                }
                muc01_by_name_th = {
                    (row["ten_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row
                    for row in muc01_rows
                }

                for r in range(25, sub_tps_row):
                    if r in [sub_trm_row, sub_trm_row + 1]:
                        continue
                    name = ws.cell(r, 3).value
                    raw_code = ws.cell(r, 1).value
                    th = ws.cell(r, 5).value or ("TRM" if r < sub_trm_row else "TPS")
                    th = str(th).strip().upper()

                    matched_data = None
                    if raw_code and str(raw_code).strip() != "None":
                        code_clean = str(raw_code).strip().upper()
                        matched_data = muc01_by_code_th.get((code_clean, th))

                    if not matched_data and name:
                        name_clean = str(name).strip().upper()
                        matched_data = muc01_by_name_th.get((name_clean, th))

                    if matched_data:
                        if not raw_code or str(raw_code).strip() == "None":
                            ws.cell(r, 1).value = matched_data.get("ma_khach", "")
                        ws.cell(r, target_col).value = matched_data["tong_m2"]
                    else:
                        ws.cell(r, target_col).value = None

        # Ghi an toàn qua tệp tạm
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, f"vtc_temp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        wb.save(temp_file)
        wb.close()

        test_wb = openpyxl.load_workbook(temp_file, read_only=True)
        test_wb.close()

        try:
            shutil.move(temp_file, template_path)
        except PermissionError:
            raise PermissionError(f"Tệp '{template_path}' đang được mở bởi ứng dụng khác (ví dụ Excel). Vui lòng đóng tệp rồi thử lại!")

        return template_path
