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
        for s in wb.sheetnames:
            if "09.11.2026" in s or "BC NGAY" in s.upper():
                return wb[s]
        return wb.active

    @staticmethod
    def analyze_master_excel_dealers(
        template_path: str,
        report_type: str,
        rule_manager: Any,
        target_date: Optional[date] = None
    ) -> List[Dict[str, Any]]:
        """
        Phân tích đối chiếu Mục 02 của tệp Excel tổng với danh sách đại lý đã xác minh trong AppData.
        Hỗ trợ cả tệp Excel tổng hợp nhất 3 khối và các tệp cũ.
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
        wb.close()
        return items

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

        markers = {}
        for r in range(1, min(ws.max_row + 1, 120)):
            vals = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, 6)]
            combined = " | ".join(vals)
            if "TỔNG NHÓM TRM + TPS" in combined or "TỔNG NHÓM TRM+TPS" in combined:
                markers["TRM_TPS_END"] = r
            elif "TỔNG NHÓM TCR + TML" in combined or "TỔNG NHÓM TCR+TML" in combined:
                markers["TCR_TML_END"] = r

        is_unified = bool(markers.get("TRM_TPS_END") and markers.get("TCR_TML_END"))

        if is_unified:
            ws.cell(27, 1).value = "Mã khách hàng"
            for it in confirmed_items:
                r = it.get("row_num")
                if r:
                    code_to_write = it.get("final_code", it.get("app_code", ""))
                    grp_to_write = it.get("final_group", it.get("app_group", ""))
                    th_to_write = it.get("th", "")
                    if code_to_write: ws.cell(r, 1).value = code_to_write
                    if grp_to_write: ws.cell(r, 5).value = grp_to_write
                    if th_to_write: ws.cell(r, 6).value = th_to_write
        else:
            if report_type == "TCR_TML":
                ws.cell(27, 1).value = "Mã khách hàng"
                for it in confirmed_items:
                    r = it.get("row_num")
                    if r:
                        code_to_write = it.get("final_code", it.get("app_code", ""))
                        grp_to_write = it.get("final_group", it.get("app_group", ""))
                        th_to_write = it.get("th", "TCR")
                        if code_to_write: ws.cell(r, 1).value = code_to_write
                        if grp_to_write: ws.cell(r, 5).value = grp_to_write
                        if th_to_write: ws.cell(r, 6).value = th_to_write
            elif report_type == "TRM_TPS":
                ws.cell(24, 1).value = "Mã khách hàng"
                for it in confirmed_items:
                    r = it.get("row_num")
                    if r:
                        code_to_write = it.get("final_code", it.get("app_code", ""))
                        grp_to_write = it.get("final_group", it.get("app_group", ""))
                        th_to_write = it.get("th", "TRM")
                        if code_to_write: ws.cell(r, 1).value = code_to_write
                        if grp_to_write: ws.cell(r, 4).value = grp_to_write
                        if th_to_write: ws.cell(r, 5).value = th_to_write
            elif report_type == "TMR":
                for it in confirmed_items:
                    r = it.get("row_num")
                    if r:
                        code_to_write = it.get("final_code", it.get("app_code", ""))
                        grp_to_write = it.get("final_group", it.get("app_group", ""))
                        if code_to_write: ws.cell(r, 1).value = code_to_write
                        if grp_to_write: ws.cell(r, 4).value = grp_to_write
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
        - Mục 01: Có mã khách và CHỈ hiện đại lý có phát sinh hợp lệ. Khi cập nhật một nhóm, giữ nguyên số liệu của các nhóm khác trong ngày.
        - Mục 02: Đồng bộ danh sách theo dõi, chia 3 nhóm (TRM+TPS, TCR+TML, TMR) + Tổng chung.
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Không tìm thấy tệp Excel tổng: {template_path}")

        # 1. Tạo bản sao lưu trước khi ghi
        ExcelWriter.create_daily_backup(template_path, backup_dir, target_date)

        # 2. Mở workbook với data_only=False để giữ nguyên công thức
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

            # Tiêu đề dòng 1
            ws.cell(1, 2).value = f"BÁO CÁO BÁN HÀNG NGÀY {day_str} NHÓM VIP - D1 - GPP"

            # Headers Mục 01 ở dòng 2
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

            # Đọc các dòng Mục 01 đang có của các nhóm KHÁC (để giữ lại)
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
                            specs_dict = {}
                            for c in range(9, 26):
                                v = ws.cell(r, c).value
                                if v is not None and v != "":
                                    try:
                                        q = float(v)
                                        if q > 0:
                                            for sp_k, sp_c in col_map.items():
                                                if sp_c == c:
                                                    specs_dict[sp_k] = q
                                                    break
                                    except (ValueError, TypeError):
                                        pass
                            existing_other_rows.append({
                                "ma_khach": str(ws.cell(r, 1).value or "").strip(),
                                "ten_khach": str(name).strip(),
                                "nhom": str(ws.cell(r, 5).value or "").strip(),
                                "thuong_hieu": th_clean,
                                "specs": specs_dict
                            })

            # Ghép danh sách Mục 01: các nhóm khác giữ nguyên + nhóm hiện tại cập nhật mới
            combined_rows = existing_other_rows + [dict(r) for r in muc01_rows]

            # Sắp xếp đúng theo thứ tự 3 nhóm: 1. TRM+TPS, 2. TCR+TML, 3. TMR
            def sort_key(item):
                th = item.get("thuong_hieu", "").upper()
                order = {"TRM": 10, "TPS": 11, "TCR": 20, "TML": 21, "TMR": 30}.get(th, 99)
                return (order, item.get("ten_khach", ""))

            combined_rows.sort(key=sort_key)

            # Ghi Mục 01 (dòng 3 đến 24)
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

                    # Công thức m2 và thùng
                    ws.cell(curr_r, 7).value = f"=I{curr_r}+J{curr_r}*1.5+K{curr_r}+L{curr_r}*1.62+M{curr_r}*1.8+N{curr_r}*1.44+O{curr_r}*1.44+P{curr_r}*1.92+Q{curr_r}*0.99+R{curr_r}*1.44+S{curr_r}*0.96+T{curr_r}*1.92+U{curr_r}*1.92+V{curr_r}*1.5+W{curr_r}*1.8+X{curr_r}*1.8+Y{curr_r}*1.92"
                    ws.cell(curr_r, 8).value = f"=SUM(I{curr_r}:Y{curr_r})"

                    for c in range(9, 26):
                        ws.cell(curr_r, c).value = None

                    for sp, qty in cust_data.get("specs", {}).items():
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

            # Tìm cột ngày tương ứng trong Mục 02 (từ Col 8 trở đi)
            target_col = None
            for c in range(8, ws.max_column + 1):
                cell_val = ws.cell(27, c).value
                if isinstance(cell_val, datetime) and cell_val.date() == target_date:
                    target_col = c
                    break
                elif isinstance(cell_val, date) and cell_val == target_date:
                    target_col = c
                    break
                elif cell_val and day_header_short in str(cell_val):
                    target_col = c
                    break

            if not target_col:
                target_col = max(ws.max_column + 1, 8)
                ws.cell(27, target_col).value = datetime.combine(target_date, datetime.min.time())
                dow_map = {0: "THỨ 2", 1: "THỨ 3", 2: "THỨ 4", 3: "THỨ 5", 4: "THỨ 6", 5: "THỨ 7", 6: "CHỦ NHẬT"}
                ws.cell(28, target_col).value = dow_map.get(target_date.weekday(), "")

            # Re-scan markers on current ws
            sub_trm_row = 45
            sub_tcr_row = 68
            sub_tmr_row = 72
            grand_total_row = 73

            for r in range(28, min(ws.max_row + 1, 120)):
                v = str(ws.cell(r, 2).value or "").upper()
                if "TỔNG NHÓM TRM + TPS" in v or "TỔNG NHÓM TRM+TPS" in v:
                    sub_trm_row = r
                elif "TỔNG NHÓM TCR + TML" in v or "TỔNG NHÓM TCR+TML" in v:
                    sub_tcr_row = r
                elif "TỔNG NHÓM TMR" in v:
                    sub_tmr_row = r
                elif "TỔNG CHUNG" in v:
                    grand_total_row = r

            if report_type == "TRM_TPS":
                b_start = 29
                b_end = sub_trm_row - 1
                curr_sub_row = sub_trm_row
                default_th = "TRM"
            elif report_type == "TCR_TML":
                b_start = sub_trm_row + 1
                b_end = sub_tcr_row - 1
                curr_sub_row = sub_tcr_row
                default_th = "TCR"
            else:  # TMR
                b_start = sub_tcr_row + 1
                b_end = sub_tmr_row - 1
                curr_sub_row = sub_tmr_row
                default_th = "TMR"

            app_dealers = rule_manager.get_dealers(report_type) if rule_manager else []
            app_by_code_th = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", "").strip().upper()): d for d in app_dealers if d.get("ma_khach")}
            app_by_name_th = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "").strip().upper()): d for d in app_dealers}
            app_by_name_only = {normalize_dealer_name(d.get("ten_khach", "")): d for d in app_dealers}

            muc01_by_code_th = {(row["ma_khach"].strip().upper(), row["thuong_hieu"].strip().upper()): row for row in muc01_rows if row.get("ma_khach")}
            muc01_by_name_th = {(normalize_dealer_name(row["ten_khach"]), row["thuong_hieu"].strip().upper()): row for row in muc01_rows}
            muc01_by_name_only = {normalize_dealer_name(row["ten_khach"]): row for row in muc01_rows}

            for r in range(b_start, b_end + 1):
                raw_code = str(ws.cell(r, 1).value or "").strip()
                if raw_code.upper() in ["NONE", "MÃ KHÁCH HÀNG", "MÃ KH"]: raw_code = ""
                name = str(ws.cell(r, 4).value or "").strip()
                name_norm = normalize_dealer_name(name)
                th = str(ws.cell(r, 6).value or default_th).strip().upper()

                rule_dealer = None
                if raw_code:
                    rule_dealer = app_by_code_th.get((raw_code.upper(), th))
                if not rule_dealer and name_norm:
                    rule_dealer = app_by_name_th.get((name_norm, th))
                if not rule_dealer and name_norm:
                    rule_dealer = app_by_name_only.get(name_norm)

                if not raw_code and rule_dealer and rule_dealer.get("ma_khach"):
                    if not ("BẢO VIỆT" in name.upper() and rule_dealer["ma_khach"] == "CH9999"):
                        raw_code = rule_dealer["ma_khach"]
                        ws.cell(r, 1).value = raw_code

                if rule_dealer:
                    if not ws.cell(r, 5).value and rule_dealer.get("nhom"):
                        ws.cell(r, 5).value = rule_dealer["nhom"]
                    if not ws.cell(r, 6).value and rule_dealer.get("thuong_hieu"):
                        ws.cell(r, 6).value = rule_dealer["thuong_hieu"]
                        th = rule_dealer["thuong_hieu"].upper()

                matched = None
                if raw_code:
                    matched = muc01_by_code_th.get((raw_code.upper(), th))
                if not matched and name:
                    matched = muc01_by_name_th.get((name_norm, th))
                if not matched and name:
                    matched = muc01_by_name_only.get(name_norm)

                if matched:
                    ws.cell(r, target_col).value = matched["tong_m2"]
                else:
                    ws.cell(r, target_col).value = None

                # Công thức Col 7 dòng đại lý
                last_col_l = get_column_letter(max(27, ws.max_column))
                ws.cell(r, 7).value = f"=SUM(H{r}:{last_col_l}{r})"

            # Cập nhật công thức tổng nhóm và tổng chung tại target_col và Col 7
            cl_t = get_column_letter(target_col)
            # Tổng nhóm TRM+TPS
            ws.cell(sub_trm_row, target_col).value = f"=SUM({cl_t}29:{cl_t}{sub_trm_row - 1})"
            ws.cell(sub_trm_row, 7).value = f"=SUM(G29:G{sub_trm_row - 1})"
            # Tổng nhóm TCR+TML
            ws.cell(sub_tcr_row, target_col).value = f"=SUM({cl_t}{sub_trm_row + 1}:{cl_t}{sub_tcr_row - 1})"
            ws.cell(sub_tcr_row, 7).value = f"=SUM(G{sub_trm_row + 1}:G{sub_tcr_row - 1})"
            # Tổng nhóm TMR
            ws.cell(sub_tmr_row, target_col).value = f"=SUM({cl_t}{sub_tcr_row + 1}:{cl_t}{sub_tmr_row - 1})"
            ws.cell(sub_tmr_row, 7).value = f"=SUM(G{sub_tcr_row + 1}:G{sub_tmr_row - 1})"
            # Tổng chung
            ws.cell(grand_total_row, target_col).value = f"=SUM({cl_t}{sub_trm_row},{cl_t}{sub_tcr_row},{cl_t}{sub_tmr_row})"
            ws.cell(grand_total_row, 7).value = f"=SUM(G{sub_trm_row},G{sub_tcr_row},G{sub_tmr_row})"

        else:
            # === FALLBACK CHO MẪU CŨ RIÊNG RẼ ===
            if report_type == "TCR_TML":
                if sheet_day_str in wb.sheetnames: ws = wb[sheet_day_str]
                else:
                    ws = wb.copy_worksheet(wb.active)
                    ws.title = sheet_day_str

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
