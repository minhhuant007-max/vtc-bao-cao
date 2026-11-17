import openpyxl
from datetime import datetime, date
from typing import List, Dict, Any, Optional, Tuple
import os

from .models import OrderRow, TrackedCustomer

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
    def read_tracked_customers(template_path: str, report_type: str, sheet_name: Optional[str] = None) -> List[TrackedCustomer]:
        """
        Đọc danh sách đại lý được theo dõi ở Mục 02 của mẫu đích.
        Tự động nhận diện cột Mã khách hàng, STT, Cửa hàng, Nhóm, Thương hiệu.
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Không tìm thấy mẫu đích: {template_path}")

        wb = openpyxl.load_workbook(template_path, data_only=True)
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            # Lấy sheet phù hợp nhất
            if report_type == "TCR_TML":
                ws = wb.active
            elif report_type == "TMR":
                # Thường là thang 11.2026 hoặc sheet cuối
                ws = wb["thang 11.2026"] if "thang 11.2026" in wb.sheetnames else wb.active
            elif report_type == "TRM_TPS":
                ws = wb["BC NGAY T11.2026"] if "BC NGAY T11.2026" in wb.sheetnames else wb.active
            else:
                ws = wb.active

        customers: List[TrackedCustomer] = []

        if report_type == "TCR_TML":
            # Mục 02 từ dòng 29 đến dòng 51
            # Header ở dòng 27: [None, 'STT', None, 'CỬA HÀNG', 'NHÓM', 'TH', ...]
            for r in range(29, min(52, ws.max_row + 1)):
                name = ws.cell(r, 4).value
                nhom = ws.cell(r, 5).value
                th = ws.cell(r, 6).value
                stt_val = ws.cell(r, 2).value or ws.cell(r, 3).value
                raw_code = ws.cell(r, 1).value
                if name and nhom and th and str(name).strip():
                    name_clean = str(name).strip()
                    code = str(raw_code).strip().upper() if raw_code else KNOWN_CUSTOMER_CODES.get(name_clean, "")
                    try:
                        stt = int(stt_val)
                    except (ValueError, TypeError):
                        stt = len(customers) + 1
                    customers.append(TrackedCustomer(
                        ma_khach=code,
                        stt=stt,
                        ten_khach=name_clean,
                        nhom=str(nhom).strip().upper(),
                        thuong_hieu=str(th).strip().upper(),
                        row_idx_muc02=r,
                        col_idx_ma_khach=1
                    ))

        elif report_type == "TMR":
            # Mục 02 từ dòng 10 đến dòng 12
            # Dòng 9: ['STT', 'CỬA HÀNG', 'NHÓM', 'TỔNG', ...]
            for r in range(10, min(20, ws.max_row + 1)):
                name = ws.cell(r, 2).value
                nhom = ws.cell(r, 3).value
                stt_val = ws.cell(r, 1).value
                if name and nhom and str(name).strip():
                    name_clean = str(name).strip()
                    code = KNOWN_CUSTOMER_CODES.get(name_clean, "")
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
            # Mục 02 từ dòng 25 đến dòng 42
            # Dòng 24: ['Mã khách hàng', 'Stt', 'CỬA HÀNG', 'NHÓM', 'TH', 'TỔNG', ...] hoặc ['Mã khách hàng', 'Stt', 'CỬA HÀNG', 'NHÓM', 'TỔNG', ...]
            # Dòng 25-33: TOROMA (TRM)
            # Dòng 36-42: TOPISA (TPS)
            current_th = "TRM"
            for r in range(25, min(45, ws.max_row + 1)):
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

                # Nếu cột 5 là TH đã điền TRM / TPS
                if th_val and str(th_val).strip().upper() in ["TRM", "TPS"]:
                    current_th = str(th_val).strip().upper()

                if name and nhom and str(name).strip():
                    name_clean = str(name).strip()
                    code = str(raw_code).strip().upper() if raw_code and str(raw_code).strip() != "None" else KNOWN_CUSTOMER_CODES.get(name_clean, "")
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
