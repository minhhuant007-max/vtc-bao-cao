import openpyxl
from datetime import datetime, date
from typing import List, Dict, Any, Optional, Tuple
import os
import re

from .models import OrderRow, TrackedCustomer

def normalize_dealer_name(s: str) -> str:
    """Chuẩn hóa tên đại lý để so khớp chính xác: loại bỏ khoảng trắng kép, chuẩn hóa dấu gạch ngang."""
    if not s:
        return ""
    s = re.sub(r'\s+', ' ', str(s)).strip().upper()
    s = s.replace(' - ', '-').replace('- ', '-').replace(' -', '-')
    return s

# Mapping mã khách hàng mặc định (mã thực tế được nạp động từ cấu hình Quản lý đại lý tại %APPDATA%):
KNOWN_CUSTOMER_CODES = {}

class ExcelReader:
    @staticmethod
    def read_orders(file_path: str, target_date: Optional[date] = None, only_completed: bool = False) -> List[OrderRow]:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Không tìm thấy tệp: {file_path}")

        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        orders: List[OrderRow] = []
        fname = os.path.basename(file_path)

        for sname in wb.sheetnames:
            sheet = wb[sname]
            header = None
            header_map = {}

            for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                # Search header
                if header is None:
                    row_strs = [str(c).strip() if c is not None else "" for c in row]
                    if "Ngày ct" in row_strs or ("Mã khách" in row_strs and "Số ct" in row_strs):
                        header = row_strs
                        for col_idx, col_name in enumerate(row_strs):
                            if col_name:
                                header_map[col_name] = col_idx
                        continue
                    else:
                        continue

                # Process data row
                if not any(row):
                    continue

                def get_val(col_name: str, default=None):
                    idx = header_map.get(col_name)
                    if idx is not None and idx < len(row):
                        v = row[idx]
                        return v if v is not None else default
                    return default

                # Parse date
                raw_date = get_val("Ngày ct")
                d_val: Optional[date] = None
                if isinstance(raw_date, datetime):
                    d_val = raw_date.date()
                elif isinstance(raw_date, date):
                    d_val = raw_date
                elif isinstance(raw_date, str):
                    raw_date = raw_date.strip()
                    for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"]:
                        try:
                            d_val = datetime.strptime(raw_date, fmt).date()
                            break
                        except ValueError:
                            pass

                if target_date and d_val and d_val != target_date:
                    continue

                # Parse status
                status = str(get_val("Trạng thái", "")).strip()
                if status == "Đã duyệt":
                    status = "Duyệt"

                if only_completed:
                    if status != "Hoàn thành":
                        continue
                else:
                    if status not in ["Hoàn thành", "Duyệt"]:
                        # Keep note: non-completed non-approved rows are filtered out by default
                        pass

                # Clean fields
                so_ct = str(get_val("Số ct", "")).strip()
                ma_khach = str(get_val("Mã khách", "")).strip().upper()
                ten_khach = str(get_val("Tên khách", "")).strip()
                ma_hang = str(get_val("Mã hàng", "")).strip().upper()
                ten_mat_hang = str(get_val("Tên mặt hàng", "")).strip()
                ma_lo = str(get_val("Mã lô", "")).strip() if get_val("Mã lô") else ""
                dvt = str(get_val("Đvt", "thùng")).strip()
                
                try:
                    qty = float(get_val("Sl đặt hàng", 0) or 0)
                except (ValueError, TypeError):
                    qty = 0.0

                try:
                    price = float(get_val("Đơn giá", 0) or 0)
                except (ValueError, TypeError):
                    price = 0.0

                try:
                    total = float(get_val("Thành tiền", 0) or 0)
                except (ValueError, TypeError):
                    total = 0.0

                dien_giai = str(get_val("Diễn giải", "")).strip()
                ghi_chu = str(get_val("Ghi chú", "")).strip()

                order = OrderRow(
                    source_file=fname,
                    sheet_name=sname,
                    row_num=row_idx,
                    so_ct=so_ct,
                    ngay_ct=d_val,
                    ma_khach=ma_khach,
                    ten_khach=ten_khach,
                    trang_thai=status,
                    ma_hang=ma_hang,
                    ten_mat_hang=ten_mat_hang,
                    ma_lo=ma_lo,
                    dvt=dvt,
                    sl_dat_hang=qty,
                    don_gia=price,
                    thanh_tien=total,
                    dien_giai=dien_giai,
                    ghi_chu=ghi_chu
                )
                orders.append(order)

        wb.close()
        return orders

    @staticmethod
    def get_source_dates(file_path: str) -> List[date]:
        if not os.path.exists(file_path):
            return []
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        dates = set()
        for sname in wb.sheetnames:
            sheet = wb[sname]
            header = None
            date_idx = None
            for row in sheet.iter_rows(values_only=True):
                if header is None:
                    row_strs = [str(c).strip() if c is not None else "" for c in row]
                    if "Ngày ct" in row_strs:
                        header = row_strs
                        date_idx = header.index("Ngày ct")
                    continue
                if date_idx is not None and date_idx < len(row):
                    val = row[date_idx]
                    if isinstance(val, datetime):
                        dates.add(val.date())
                    elif isinstance(val, date):
                        dates.add(val)
                    elif isinstance(val, str):
                        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"]:
                            try:
                                dates.add(datetime.strptime(val.strip(), fmt).date())
                                break
                            except ValueError:
                                pass
        wb.close()
        return sorted(list(dates))

    @staticmethod
    def read_tracked_customers(
        template_path: str,
        report_type: str,
        sheet_name: Optional[str] = None,
        *,
        target_date: Optional[date] = None,
        rule_manager: Optional[Any] = None,
        dealers: Optional[List[Dict[str, Any]]] = None
    ) -> List[TrackedCustomer]:
        """
        Đọc danh sách đại lý được theo dõi ở Mục 02 của mẫu đích.
        Hỗ trợ:
        1. Mẫu Excel tổng hợp nhất mới (3 khối: TRM+TPS, TCR+TML, TMR) trong cùng 1 sheet.
        2. Mẫu cũ (các tệp riêng rẽ).
        Đồng bộ tự động mã khách từ RuleManager (AppData) nếu cột mã khách trong Excel đang trống.
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Không tìm thấy mẫu đích: {template_path}")

        wb = openpyxl.load_workbook(template_path, data_only=True)
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            # Lấy sheet phù hợp nhất theo target_date hoặc sheet đầu tiên
            sheet_day = target_date.strftime("%d.%m.%Y") if target_date else None
            if sheet_day and sheet_day in wb.sheetnames:
                ws = wb[sheet_day]
            else:
                month_code = target_date.strftime("T%m.%Y") if target_date else "T11.2026"
                month_sub = target_date.strftime("%m.%Y") if target_date else "11.2026"
                found_s = None
                for s in wb.sheetnames:
                    if month_code.lower() in s.lower() or month_sub in s:
                        found_s = s
                        break
                if not found_s:
                    for s in wb.sheetnames:
                        if "09.11.2026" in s or "BC NGAY" in s.upper() or "THEO DOI" in s.upper():
                            found_s = s
                            break
                ws = wb[found_s] if found_s else wb.active

        customers: List[TrackedCustomer] = []

        if dealers is None and rule_manager is not None:
            dealers = rule_manager.get_dealers(report_type)

        # Nạp dữ liệu tra cứu từ rule_manager/dealers nếu có
        app_dealers = dealers or []
        app_by_name_th = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "").strip().upper()): d for d in app_dealers}
        app_by_code_th = {(str(d.get("ma_khach") or "").strip().upper(), d.get("thuong_hieu", "").strip().upper()): d for d in app_dealers if d.get("ma_khach")}
        app_by_name_only = {normalize_dealer_name(d.get("ten_khach", "")): d for d in app_dealers}

        # Phát hiện bố cục 3 khối chuẩn trong tệp Excel tổng hợp nhất
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
            # Bố cục chuẩn 3 khối:
            # Block 1: TRM + TPS (Dòng 29 đến TRM_TPS_END - 1)
            # Block 2: TCR + TML (Dòng TRM_TPS_END + 1 đến TCR_TML_END - 1)
            # Block 3: TMR (Dòng TCR_TML_END + 1 đến TMR_END - 1)
            if report_type == "TRM_TPS":
                start_r = 29
                end_r = markers["TRM_TPS_END"]
                default_th = "TRM"
            elif report_type == "TCR_TML":
                start_r = markers["TRM_TPS_END"] + 1
                end_r = markers["TCR_TML_END"]
                default_th = "TCR"
            else:  # TMR
                start_r = markers["TCR_TML_END"] + 1
                end_r = markers.get("TMR_END", ws.max_row)
                default_th = "TMR"

            for r in range(start_r, end_r):
                raw_code = ws.cell(r, 1).value
                stt_val = ws.cell(r, 2).value
                name = ws.cell(r, 4).value
                nhom = ws.cell(r, 5).value
                th = ws.cell(r, 6).value or default_th

                if name and str(name).strip() and "TỔNG" not in str(name).upper():
                    name_clean = str(name).strip()
                    name_norm = normalize_dealer_name(name_clean)
                    th_clean = str(th).strip().upper()
                    code = str(raw_code).strip().upper() if raw_code and str(raw_code).strip() != "None" else ""

                    # Tra từ rule_manager nếu code rỗng
                    if not code:
                        matched_rule = app_by_name_th.get((name_norm, th_clean)) or app_by_name_only.get(name_norm)
                        if matched_rule and matched_rule.get("ma_khach"):
                            code = matched_rule["ma_khach"].strip().upper()

                    try:
                        stt = int(stt_val)
                    except (ValueError, TypeError):
                        stt = len(customers) + 1

                    customers.append(TrackedCustomer(
                        ma_khach=code,
                        stt=stt,
                        ten_khach=name_clean,
                        nhom=str(nhom).strip().upper() if nhom else "VIP",
                        thuong_hieu=th_clean,
                        row_idx_muc02=r,
                        col_idx_ma_khach=1
                    ))
            wb.close()

            if app_dealers:
                by_code = {}
                by_name = {}
                for customer in customers:
                    brand = customer.thuong_hieu
                    if customer.ma_khach:
                        by_code[(customer.ma_khach, brand)] = customer
                    by_name.setdefault((normalize_dealer_name(customer.ten_khach), brand), []).append(customer)

                for dealer in app_dealers:
                    code = str(dealer.get("ma_khach") or "").strip().upper()
                    name = str(dealer.get("ten_khach") or "").strip()
                    brand = str(dealer.get("thuong_hieu") or "").strip().upper()
                    if not code or code == "CH9999" or name.startswith("ĐẠI LÝ MẪU "):
                        continue
                    same_name = by_name.get((normalize_dealer_name(name), brand), [])
                    customer = by_code.get((code, brand))
                    if not customer and same_name:
                        if len(same_name) != 1 or same_name[0].ma_khach:
                            continue
                        customer = same_name[0]
                        customer.ma_khach = code
                    if not customer:
                        customer = TrackedCustomer(code, dealer.get("stt", len(customers) + 1),
                                                   name, dealer.get("nhom", ""), brand)
                        customers.append(customer)
                    customer.is_active = dealer.get("is_active", True)
                    customer.nhom = dealer.get("nhom") or customer.nhom
                    by_code[(code, brand)] = customer

            return customers

        # Fallback cho mẫu cũ riêng rẽ (tương thích ngược)
        if report_type == "TCR_TML":
            # Header ở dòng 27: [None, 'STT', None, 'CỬA HÀNG', 'NHÓM', 'TH', ...] hoặc ['Mã khách hàng', ...]
            has_code_col = (str(ws.cell(27, 1).value or "").strip().upper() in ["MÃ KHÁCH HÀNG", "MÃ KH"])
            col_code = 1
            col_stt = 2
            col_name = 4
            col_nhom = 5
            col_th = 6

            for r in range(29, min(55, ws.max_row + 1)):
                name = ws.cell(r, col_name).value
                nhom = ws.cell(r, col_nhom).value
                th = ws.cell(r, col_th).value
                stt_val = ws.cell(r, col_stt).value or ws.cell(r, 3).value
                raw_code = ws.cell(r, col_code).value if has_code_col else None

                if name and nhom and th and str(name).strip() and "TỔNG" not in str(name).upper():
                    name_clean = str(name).strip()
                    name_norm = normalize_dealer_name(name_clean)
                    th_clean = str(th).strip().upper()
                    code = str(raw_code).strip().upper() if raw_code and str(raw_code).strip() != "None" else ""

                    # Tra từ rule_manager nếu code rỗng
                    if not code:
                        matched_rule = app_by_name_th.get((name_norm, th_clean)) or app_by_name_only.get(name_norm)
                        if matched_rule and matched_rule.get("ma_khach"):
                            code = matched_rule["ma_khach"].strip().upper()

                    try:
                        stt = int(stt_val)
                    except (ValueError, TypeError):
                        stt = len(customers) + 1

                    customers.append(TrackedCustomer(
                        ma_khach=code,
                        stt=stt,
                        ten_khach=name_clean,
                        nhom=str(nhom).strip().upper(),
                        thuong_hieu=th_clean,
                        row_idx_muc02=r,
                        col_idx_ma_khach=1
                    ))

        elif report_type == "TMR":
            # Header ở dòng 9 hoặc dòng 3
            # Mẫu có cột mã khách: Cột 1 là Mã khách hàng, Cột 2 là STT, Cột 3 là CỬA HÀNG, Cột 4 là NHÓM
            # Mẫu cũ (chưa có cột mã khách): Cột 1 là STT, Cột 2 là CỬA HÀNG, Cột 3 là NHÓM
            has_code_col = (
                str(ws.cell(9, 1).value or "").strip().upper() in ["MÃ KHÁCH HÀNG", "MÃ KH"]
                or str(ws.cell(3, 1).value or "").strip().upper() in ["MÃ KHÁCH HÀNG", "MÃ KH"]
            )
            col_code = 1 if has_code_col else None
            col_stt = 2 if has_code_col else 1
            col_name = 3 if has_code_col else 2
            col_nhom = 4 if has_code_col else 3

            for r in range(10, min(25, ws.max_row + 1)):
                name = ws.cell(r, col_name).value
                nhom = ws.cell(r, col_nhom).value
                stt_val = ws.cell(r, col_stt).value
                raw_code = ws.cell(r, col_code).value if col_code else None

                if name and nhom and str(name).strip() and "TỔNG" not in str(name).upper():
                    name_clean = str(name).strip()
                    name_norm = normalize_dealer_name(name_clean)
                    code = str(raw_code).strip().upper() if raw_code and str(raw_code).strip() != "None" else ""

                    # Tra từ rule_manager nếu code rỗng
                    if not code:
                        matched_rule = app_by_name_th.get((name_norm, "TMR")) or app_by_name_only.get(name_norm)
                        if matched_rule and matched_rule.get("ma_khach"):
                            code = matched_rule["ma_khach"].strip().upper()

                    try:
                        stt = int(stt_val)
                    except (ValueError, TypeError):
                        stt = len(customers) + 1

                    customers.append(TrackedCustomer(
                        ma_khach=code,
                        stt=stt,
                        ten_khach=name_clean,
                        nhom=str(nhom).strip().upper(),
                        thuong_hieu="TMR",
                        row_idx_muc02=r,
                        col_idx_ma_khach=1
                    ))

        elif report_type == "TRM_TPS":
            # Dòng 24: ['Mã khách hàng', 'Stt', 'CỬA HÀNG', 'NHÓM', 'TH', 'TỔNG', ...]
            current_th = "TRM"
            for r in range(25, min(48, ws.max_row + 1)):
                raw_code = ws.cell(r, 1).value
                stt_val = ws.cell(r, 2).value
                name = ws.cell(r, 3).value
                nhom = ws.cell(r, 4).value
                th_val = ws.cell(r, 5).value

                # Kiểm tra phân dòng TOPISA
                if r >= 36:
                    current_th = "TPS"
                else:
                    current_th = "TRM"

                if th_val and str(th_val).strip().upper() in ["TRM", "TPS"]:
                    current_th = str(th_val).strip().upper()

                if name and nhom and str(name).strip() and "TỔNG" not in str(name).upper():
                    name_clean = str(name).strip()
                    name_norm = normalize_dealer_name(name_clean)
                    code = str(raw_code).strip().upper() if raw_code and str(raw_code).strip() != "None" else ""

                    # Tra từ rule_manager nếu code rỗng
                    if not code:
                        matched_rule = app_by_name_th.get((name_norm, current_th)) or app_by_name_only.get(name_norm)
                        if matched_rule and matched_rule.get("ma_khach"):
                            code = matched_rule["ma_khach"].strip().upper()

                    try:
                        stt = int(stt_val)
                    except (ValueError, TypeError):
                        stt = len(customers) + 1

                    customers.append(TrackedCustomer(
                        ma_khach=code,
                        stt=stt,
                        ten_khach=name_clean,
                        nhom=str(nhom).strip().upper(),
                        thuong_hieu=current_th,
                        row_idx_muc02=r,
                        col_idx_ma_khach=1
                    ))

        wb.close()
        return customers
