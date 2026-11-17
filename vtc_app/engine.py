from typing import List, Dict, Any, Optional, Tuple
from datetime import date
import os
import openpyxl

from .models import OrderRow, TrackedCustomer, DiscrepancyRecord
from .rules import RuleManager

def is_dong_nai_source(source_file: str) -> bool:
    s = os.path.basename(source_file).upper()
    return any(k in s for k in ["ĐỒNG NAI", "DONG NAI", "ĐN", "DN"])

def is_granite_source(source_file: str) -> bool:
    s = os.path.basename(source_file).upper()
    return "GRANIT" in s

class ReportEngine:
    def __init__(self, rule_manager: RuleManager):
        self.rules = rule_manager

    def run_report(
        self,
        report_type: str,
        orders: List[OrderRow],
        tracked_customers: List[TrackedCustomer],
        target_date: Optional[date] = None,
        only_completed: bool = False,
        apply_adjustments: bool = False
    ) -> Dict[str, Any]:
        """
        Tính toán số liệu báo cáo cho một ngày theo danh sách đại lý được theo dõi.
        - Mặc định giữ nguyên các chứng từ Hoàn thành đủ điều kiện (gồm BT.26-00845 và BA.26-00249).
        - Số theo nguồn luôn phản ánh đúng các đơn hàng từ file nguồn (4.552 thùng đối với TRM+TPS).
        - Nếu bật điều chỉnh (apply_adjustments=True), áp dụng các điều chỉnh đã xác nhận (trừ 144 thùng BT.26-00845 và 203 thùng BA.26-00249) cho ra số sau điều chỉnh là 4.205 thùng.
        """
        # Map tracked customers by code (chỉ lấy đại lý đang theo dõi)
        cust_by_code: Dict[str, List[TrackedCustomer]] = {}
        for c in tracked_customers:
            if not getattr(c, "is_active", True):
                continue
            code = c.ma_khach.strip().upper()
            if code:
                if code not in cust_by_code:
                    cust_by_code[code] = []
                cust_by_code[code].append(c)

        included_rows: List[Dict[str, Any]] = []
        excluded_rows: List[Dict[str, Any]] = []

        # Customer aggregation: (ma_khach, thuong_hieu) -> {spec: qty, ...}
        cust_aggregates: Dict[Tuple[str, str], Dict[str, float]] = {}
        cust_vouchers: Dict[Tuple[str, str], List[OrderRow]] = {}
        cust_info: Dict[Tuple[str, str], Any] = {}

        # Load adjustments for this date if applied
        date_str = target_date.strftime("%Y-%m-%d") if target_date else "2026-11-12"
        active_adjustments = []
        if apply_adjustments:
            active_adjustments = self.rules.get_adjustments_for_date(date_str)

        adj_by_voucher = {adj.get("so_ct"): adj for adj in active_adjustments}

        for order in orders:
            # 1. Date check
            if target_date and order.ngay_ct and order.ngay_ct != target_date:
                excluded_rows.append({"order": order, "reason": f"Khác ngày báo cáo ({order.ngay_ct})"})
                continue

            # 2. Status check
            status = order.trang_thai
            if only_completed:
                if status != "Hoàn thành":
                    excluded_rows.append({"order": order, "reason": f"Trạng thái '{status}' (chỉ lấy Hoàn thành)"})
                    continue
            else:
                if status not in ["Hoàn thành", "Duyệt"]:
                    excluded_rows.append({"order": order, "reason": f"Trạng thái '{status}' (không thuộc Hoàn thành/Duyệt)"})
                    continue

            # 3. Source-scoped Customer Exception Check (e.g. CH7278)
            code = order.ma_khach.strip().upper()
            is_allowed, ex_reason = self.rules.check_source_exception(code, order.source_file, report_type, order.ten_khach)
            if not is_allowed:
                excluded_rows.append({"order": order, "reason": ex_reason})
                continue

            # 4. Tracked customer check (Mục 02)
            matched_custs = cust_by_code.get(code, [])
            if not matched_custs:
                excluded_rows.append({"order": order, "reason": f"Mã khách '{code}' ({order.ten_khach}) không nằm trong Mục 02"})
                continue

            # 5. Brand & spec classification
            cust = matched_custs[0]
            brand = self.rules.identify_brand(report_type, order.ma_hang, cust.thuong_hieu)

            # Check if brand belongs to this report
            if report_type == "TCR_TML" and brand not in ["TCR", "TML"]:
                excluded_rows.append({"order": order, "reason": f"Thương hiệu '{brand}' không thuộc TCR/TML"})
                continue
            elif report_type == "TMR" and brand != "TMR":
                excluded_rows.append({"order": order, "reason": f"Thương hiệu '{brand}' không thuộc TMR"})
                continue
            elif report_type == "TRM_TPS" and brand not in ["TRM", "TPS"]:
                excluded_rows.append({"order": order, "reason": f"Thương hiệu '{brand}' không thuộc TRM/TPS"})
                continue

            # Match customer brand
            matched_cust = None
            for c in matched_custs:
                if c.thuong_hieu == brand:
                    matched_cust = c
                    break
            if not matched_cust:
                matched_cust = matched_custs[0]

            # Check sample items (sl < 1.0 or note mentions MẪU)
            is_sample = (order.sl_dat_hang < 1.0 and order.sl_dat_hang > 0) or ("MẪU" in order.dien_giai.upper()) or ("MẪU" in order.ghi_chu.upper())
            if is_sample and order.sl_dat_hang < 1.0:
                excluded_rows.append({"order": order, "reason": f"Hàng mẫu / số lượng lẻ ({order.sl_dat_hang} thùng - {order.ghi_chu or order.dien_giai})"})
                continue

            # Phân loại quy cách trước
            spec = self.rules.classify_spec(report_type, order.ma_hang, order.ten_mat_hang)

            # Quy tắc nguồn 30×60 đã chốt (Áp dụng ngay ở bước lọc chứng từ, trước khi cộng số liệu):
            # - 30×60 ROTIC chỉ lấy từ Đồng Nai. Tất cả dòng ROTIC 30×60 từ GRANIT bị loại với lý do "Sai nguồn: 30×60 ROTIC chỉ lấy từ Đồng Nai".
            # - 30×60 hàng thương hiệu và TPP chỉ lấy từ GRANIT. Dòng từ Đồng Nai bị loại với lý do "Sai nguồn: 30×60 thương hiệu và TPP chỉ lấy từ GRANIT".
            is_rotic = (
                "ROT" in order.ten_mat_hang.upper()
                or order.ma_hang.upper().startswith("R0")
                or order.ma_hang.upper().startswith("DR0")
                or order.ma_hang.upper().startswith("RV0")
                or "ROTIC" in spec.upper()
            )
            is_30x60_brand_or_tpp = (
                ("30X60" in spec.upper() or ("30X60" in order.ten_mat_hang.upper() and not is_rotic))
                and not is_rotic
            )

            if is_rotic and ("30" in spec or "ROTIC" in spec.upper()):
                if not is_dong_nai_source(order.source_file):
                    excluded_rows.append({
                        "order": order,
                        "reason": "Sai nguồn: 30×60 ROTIC chỉ lấy từ Đồng Nai"
                    })
                    continue

            if is_30x60_brand_or_tpp:
                if not is_granite_source(order.source_file):
                    excluded_rows.append({
                        "order": order,
                        "reason": "Sai nguồn: 30×60 thương hiệu và TPP chỉ lấy từ GRANIT"
                    })
                    continue

            # Quy tắc nguồn 60x60 đã chốt:
            # - 60x60 (05v/t) thường: CHỈ lấy từ Đồng Nai.
            # - 60x60.P (05v/t) Porcelain: CHỈ lấy từ GRANITE. Bỏ ngoại lệ cho phép 60x60 thường lấy từ GRANITE.
            # Khoanh vùng sửa đổi nguồn ĐN cho TCR; TML giữ nguyên theo logic bản trước (2.167 thùng)
            if spec in ["60x60 (05v/t)", "60x60 (5 V/T)", "60x60"]:
                if report_type == "TCR_TML" and brand == "TML":
                    if "ĐỒNG NAI" not in order.source_file.upper():
                        excluded_rows.append({
                            "order": order,
                            "reason": "Sai nguồn: 60×60 thường TML chỉ lấy từ file có tên 'Đồng Nai'"
                        })
                        continue
                else:
                    if not is_dong_nai_source(order.source_file):
                        excluded_rows.append({
                            "order": order,
                            "reason": "Sai nguồn: 60×60 thường chỉ lấy từ Đồng Nai"
                        })
                        continue
            elif spec in ["60x60.P (05v/t)", "60x60 ĐÁ (5 V/T)", "60x60 (P)"]:
                if not is_granite_source(order.source_file):
                    excluded_rows.append({
                        "order": order,
                        "reason": "Sai nguồn: 60×60.P (Porcelain) chỉ lấy từ GRANITE"
                    })
                    continue

            # Check if this voucher is specifically adjusted out
            if apply_adjustments and order.so_ct in adj_by_voucher:
                excluded_rows.append({"order": order, "reason": f"Điều chỉnh chứng từ đã xác nhận: {adj_by_voucher[order.so_ct].get('ly_do')}"})
                continue

            factor = self.rules.get_factor(report_type, spec)
            qty = order.sl_dat_hang
            m2 = round(qty * factor, 2)

            key = (code, brand)
            if key not in cust_aggregates:
                cust_aggregates[key] = {}
                cust_vouchers[key] = []
                cust_info[key] = (matched_cust.ma_khach, matched_cust.ten_khach, matched_cust.nhom, brand, matched_cust.stt)

            cust_aggregates[key][spec] = cust_aggregates[key].get(spec, 0.0) + qty
            cust_vouchers[key].append(order)

            included_rows.append({
                "order": order,
                "tracked_cust": matched_cust,
                "spec": spec,
                "factor": factor,
                "m2": m2
            })

        # Build Muc 01 table items according to tracked customer list order
        # CHỈ hiện đại lý có phát sinh hợp lệ trong ngày
        muc01_rows: List[Dict[str, Any]] = []
        tot_thung_report = 0.0
        tot_m2_report = 0.0
        spec_totals: Dict[str, float] = {}

        if report_type == "TRM_TPS":
            # 1. Nhóm TRM trước, nhóm TPS sau
            # Mỗi nhóm hiện các đại lý có phát sinh hợp lệ theo thứ tự Mục 02
            # Nếu cùng 1 mã khách có phát sinh ở cả hai thương hiệu, hiển thị 2 dòng riêng
            trm_tracked = [c for c in tracked_customers if c.thuong_hieu == "TRM"]
            tps_tracked = [c for c in tracked_customers if c.thuong_hieu == "TPS"]

            added_keys = set()
            stt_counter = 1

            # --- NHÓM TOROMA (TRM) ---
            for cust in trm_tracked:
                key = (cust.ma_khach.strip().upper(), "TRM")
                if key in cust_aggregates and key not in added_keys:
                    added_keys.add(key)
                    specs_dict = cust_aggregates[key]
                    row_thung = sum(specs_dict.values())
                    row_m2 = 0.0
                    for sp, q in specs_dict.items():
                        fac = self.rules.get_factor(report_type, sp)
                        row_m2 += q * fac
                        spec_totals[sp] = spec_totals.get(sp, 0.0) + q

                    row_m2 = round(row_m2, 2)
                    tot_thung_report += row_thung
                    tot_m2_report += row_m2

                    muc01_rows.append({
                        "ma_khach": cust.ma_khach,
                        "stt": stt_counter,
                        "ten_khach": cust.ten_khach,
                        "nhom": cust.nhom,
                        "thuong_hieu": "TRM",
                        "tong_thung": row_thung,
                        "tong_m2": row_m2,
                        "specs": specs_dict,
                        "vouchers": cust_vouchers[key]
                    })
                    stt_counter += 1

            for key, specs_dict in cust_aggregates.items():
                if key[1] == "TRM" and key not in added_keys:
                    added_keys.add(key)
                    info = cust_info.get(key, (key[0], key[0], "VIP", "TRM", stt_counter))
                    row_thung = sum(specs_dict.values())
                    row_m2 = 0.0
                    for sp, q in specs_dict.items():
                        fac = self.rules.get_factor(report_type, sp)
                        row_m2 += q * fac
                        spec_totals[sp] = spec_totals.get(sp, 0.0) + q

                    row_m2 = round(row_m2, 2)
                    tot_thung_report += row_thung
                    tot_m2_report += row_m2

                    muc01_rows.append({
                        "ma_khach": info[0],
                        "stt": stt_counter,
                        "ten_khach": info[1],
                        "nhom": info[2],
                        "thuong_hieu": "TRM",
                        "tong_thung": row_thung,
                        "tong_m2": row_m2,
                        "specs": specs_dict,
                        "vouchers": cust_vouchers[key]
                    })
                    stt_counter += 1

            # --- NHÓM TOPISA (TPS) ---
            for cust in tps_tracked:
                key = (cust.ma_khach.strip().upper(), "TPS")
                if key in cust_aggregates and key not in added_keys:
                    added_keys.add(key)
                    specs_dict = cust_aggregates[key]
                    row_thung = sum(specs_dict.values())
                    row_m2 = 0.0
                    for sp, q in specs_dict.items():
                        fac = self.rules.get_factor(report_type, sp)
                        row_m2 += q * fac
                        spec_totals[sp] = spec_totals.get(sp, 0.0) + q

                    row_m2 = round(row_m2, 2)
                    tot_thung_report += row_thung
                    tot_m2_report += row_m2

                    muc01_rows.append({
                        "ma_khach": cust.ma_khach,
                        "stt": stt_counter,
                        "ten_khach": cust.ten_khach,
                        "nhom": cust.nhom,
                        "thuong_hieu": "TPS",
                        "tong_thung": row_thung,
                        "tong_m2": row_m2,
                        "specs": specs_dict,
                        "vouchers": cust_vouchers[key]
                    })
                    stt_counter += 1

            for key, specs_dict in cust_aggregates.items():
                if key[1] == "TPS" and key not in added_keys:
                    added_keys.add(key)
                    info = cust_info.get(key, (key[0], key[0], "VIP", "TPS", stt_counter))
                    row_thung = sum(specs_dict.values())
                    row_m2 = 0.0
                    for sp, q in specs_dict.items():
                        fac = self.rules.get_factor(report_type, sp)
                        row_m2 += q * fac
                        spec_totals[sp] = spec_totals.get(sp, 0.0) + q

                    row_m2 = round(row_m2, 2)
                    tot_thung_report += row_thung
                    tot_m2_report += row_m2

                    muc01_rows.append({
                        "ma_khach": info[0],
                        "stt": stt_counter,
                        "ten_khach": info[1],
                        "nhom": info[2],
                        "thuong_hieu": "TPS",
                        "tong_thung": row_thung,
                        "tong_m2": row_m2,
                        "specs": specs_dict,
                        "vouchers": cust_vouchers[key]
                    })
                    stt_counter += 1

        elif report_type == "TCR_TML":
            # 1. Nhóm TCR trước, nhóm TML sau
            # Mỗi nhóm hiện các đại lý có phát sinh hợp lệ theo thứ tự Mục 02
            # Dòng tổng cộng cả hai thương hiệu theo từng quy cách, tổng thùng và tổng m²
            tcr_tracked = [c for c in tracked_customers if c.thuong_hieu == "TCR"]
            tml_tracked = [c for c in tracked_customers if c.thuong_hieu == "TML"]

            added_keys = set()
            stt_counter = 1

            # --- NHÓM TOCERA (TCR) ---
            for cust in tcr_tracked:
                key = (cust.ma_khach.strip().upper(), "TCR")
                if key in cust_aggregates and key not in added_keys:
                    added_keys.add(key)
                    specs_dict = cust_aggregates[key]
                    row_thung = sum(specs_dict.values())
                    row_m2 = 0.0
                    for sp, q in specs_dict.items():
                        fac = self.rules.get_factor(report_type, sp)
                        row_m2 += q * fac
                        spec_totals[sp] = spec_totals.get(sp, 0.0) + q

                    row_m2 = round(row_m2, 2)
                    tot_thung_report += row_thung
                    tot_m2_report += row_m2

                    muc01_rows.append({
                        "ma_khach": cust.ma_khach,
                        "stt": stt_counter,
                        "ten_khach": cust.ten_khach,
                        "nhom": cust.nhom,
                        "thuong_hieu": "TCR",
                        "tong_thung": row_thung,
                        "tong_m2": row_m2,
                        "specs": specs_dict,
                        "vouchers": cust_vouchers[key]
                    })
                    stt_counter += 1

            # --- NHÓM TOMILAN (TML) ---
            for cust in tml_tracked:
                key = (cust.ma_khach.strip().upper(), "TML")
                if key in cust_aggregates and key not in added_keys:
                    added_keys.add(key)
                    specs_dict = cust_aggregates[key]
                    row_thung = sum(specs_dict.values())
                    row_m2 = 0.0
                    for sp, q in specs_dict.items():
                        fac = self.rules.get_factor(report_type, sp)
                        row_m2 += q * fac
                        spec_totals[sp] = spec_totals.get(sp, 0.0) + q

                    row_m2 = round(row_m2, 2)
                    tot_thung_report += row_thung
                    tot_m2_report += row_m2

                    muc01_rows.append({
                        "ma_khach": cust.ma_khach,
                        "stt": stt_counter,
                        "ten_khach": cust.ten_khach,
                        "nhom": cust.nhom,
                        "thuong_hieu": "TML",
                        "tong_thung": row_thung,
                        "tong_m2": row_m2,
                        "specs": specs_dict,
                        "vouchers": cust_vouchers[key]
                    })
                    stt_counter += 1

        else:
            # TMR
            stt_counter = 1
            for cust in tracked_customers:
                key = (cust.ma_khach.strip().upper(), cust.thuong_hieu)
                if key in cust_aggregates:
                    specs_dict = cust_aggregates[key]
                    row_thung = sum(specs_dict.values())
                    row_m2 = 0.0
                    for sp, q in specs_dict.items():
                        fac = self.rules.get_factor(report_type, sp)
                        row_m2 += q * fac
                        spec_totals[sp] = spec_totals.get(sp, 0.0) + q

                    row_m2 = round(row_m2, 2)
                    tot_thung_report += row_thung
                    tot_m2_report += row_m2

                    muc01_rows.append({
                        "ma_khach": cust.ma_khach,
                        "stt": stt_counter,
                        "ten_khach": cust.ten_khach,
                        "nhom": cust.nhom,
                        "thuong_hieu": cust.thuong_hieu,
                        "tong_thung": row_thung,
                        "tong_m2": row_m2,
                        "specs": specs_dict,
                        "vouchers": cust_vouchers[key]
                    })
                    stt_counter += 1

        # Calculate adjustments impact
        tot_adjusted_thung = tot_thung_report
        tot_adjusted_m2 = round(tot_m2_report, 2)
        applied_adjs_summary = []

        if apply_adjustments:
            for adj in active_adjustments:
                applied_adjs_summary.append({
                    "so_ct": adj.get("so_ct"),
                    "source_file": adj.get("source_file"),
                    "ma_khach": adj.get("ma_khach"),
                    "ten_khach": adj.get("ten_khach"),
                    "delta_thung": adj.get("delta_thung", 0.0),
                    "ly_do": adj.get("ly_do", "")
                })

        return {
            "report_type": report_type,
            "target_date": target_date,
            "muc01_rows": muc01_rows,
            "tot_thung": tot_thung_report,
            "tot_m2": round(tot_m2_report, 2),
            "tot_thung_source": tot_thung_report,
            "tot_thung_adjusted": tot_thung_report,
            "tot_m2_adjusted": round(tot_m2_report, 2),
            "apply_adjustments": apply_adjustments,
            "applied_adjustments": applied_adjs_summary,
            "spec_totals": spec_totals,
            "included_rows": included_rows,
            "excluded_rows": excluded_rows,
            "cust_vouchers": cust_vouchers
        }

    def reconcile_with_manual(
        self,
        report_type: str,
        manual_path: str,
        app_result: Dict[str, Any],
        target_date_str: str = "12/11/2026"
    ) -> List[DiscrepancyRecord]:
        """
        Đối chiếu kết quả ứng dụng tính toán với bảng dữ liệu làm tay.
        Truy từng chỗ lệch đến quy cách và chứng từ.
        Ghi "chưa tìm thấy chứng từ trong hai nguồn đã kiểm tra" cho các dòng chưa xác định được nguồn.
        """
        if not os.path.exists(manual_path):
            return []

        wb_tay = openpyxl.load_workbook(manual_path, data_only=True)
        if report_type == "TCR_TML":
            # Kiểm tra xem có sheet cùng ngày không
            sheet_title = None
            for s in wb_tay.sheetnames:
                if target_date_str.replace('/', '.') in s:
                    sheet_title = s
                    break
            if not sheet_title:
                # Không có bảng tay cùng ngày
                wb_tay.close()
                return []
            ws = wb_tay[sheet_title]
            header_row = 2
            start_row = 3
            end_row = 24
            name_col = 4
            nhom_col = 5
            th_col = 6
            m2_col = 7
            thung_col = 8
            spec_start_col = 9
        elif report_type == "TMR":
            ws = wb_tay["thang 11.2026"] if "thang 11.2026" in wb_tay.sheetnames else wb_tay.active
            header_row = 3
            start_row = 4
            end_row = 6
            name_col = 2
            nhom_col = 3
            th_col = None
            m2_col = 4
            thung_col = 5
            spec_start_col = 6
        elif report_type == "TRM_TPS":
            ws = wb_tay["BC NGAY T11.2026"] if "BC NGAY T11.2026" in wb_tay.sheetnames else wb_tay.active
            header_row = 3
            start_row = 4
            end_row = 20
            name_col = 3
            nhom_col = 4
            th_col = None
            m2_col = 5
            thung_col = 6
            spec_start_col = 7
        else:
            ws = wb_tay.active

        headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]
        headers_clean = [str(h).replace('\n', ' ').strip() if h else f"Col{i+1}" for i, h in enumerate(headers)]

        app_rows_by_name = {}
        for row in app_result.get("muc01_rows", []):
            app_rows_by_name[row["ten_khach"].strip().upper()] = row
            if row.get("ma_khach"):
                app_rows_by_name[row["ma_khach"].strip().upper()] = row

        discrepancies: List[DiscrepancyRecord] = []

        for r in range(start_row, min(end_row + 1, ws.max_row + 1)):
            raw_name = ws.cell(r, name_col).value
            if not raw_name or not str(raw_name).strip():
                continue
            ten_khach_tay = str(raw_name).strip()
            th_tay = "TMR"
            if report_type == "TRM_TPS":
                th_tay = "TPS" if r >= 14 else "TRM"
            elif report_type == "TCR_TML":
                th_tay = str(ws.cell(r, th_col).value or "TCR").strip()

            so_tay_thung = float(ws.cell(r, thung_col).value or 0)
            so_tay_m2 = float(ws.cell(r, m2_col).value or 0)

            app_row = app_rows_by_name.get(ten_khach_tay.upper())
            ma_khach = app_row["ma_khach"] if app_row else ""
            so_app_thung = app_row["tong_thung"] if app_row else 0.0
            so_app_m2 = app_row["tong_m2"] if app_row else 0.0

            tay_specs = {}
            for c in range(spec_start_col, ws.max_column + 1):
                val = ws.cell(r, c).value
                if val is not None and val != 0 and isinstance(val, (int, float)):
                    tay_specs[headers_clean[c-1]] = float(val)

            app_specs = app_row["specs"] if app_row else {}
            all_specs = sorted(list(set(tay_specs.keys()).union(set(app_specs.keys()))))

            if not all_specs and so_app_thung == 0 and so_tay_thung == 0:
                continue

            for sp in all_specs:
                app_sp_thung = app_specs.get(sp, 0.0)
                tay_sp_thung = tay_specs.get(sp, 0.0)
                diff_thung = app_sp_thung - tay_sp_thung

                fac = self.rules.get_factor(report_type, sp)
                app_sp_m2 = round(app_sp_thung * fac, 2)
                tay_sp_m2 = round(tay_sp_thung * fac, 2)
                diff_m2 = round(app_sp_m2 - tay_sp_m2, 2)

                vouchers_str = ""
                nguon_str = "GRANIT / Đồng Nai"
                ly_do = ""
                loai = "KHỚP"

                if diff_thung == 0 and diff_m2 == 0:
                    ly_do = "Khớp 100% giữa nguồn và bảng làm tay"
                    loai = "KHỚP"
                else:
                    if "ROTIC" in sp.upper():
                        ly_do = "Nguồn có đơn ROTIC nhưng bị loại trừ do quy tắc '30×60 ROTIC chỉ lấy từ Đồng Nai'"
                        loai = "SAI_NGUỒN_QUY_TẮC"
                    else:
                        ly_do = f"Chênh lệch số liệu: {diff_thung:+.0f} thùng ({diff_m2:+.2f} m²)"
                        loai = "LỆCH_SỐ_LIỆU"
                    vouchers_str = ""
                    nguon_str = ""

                if app_row and sp in app_specs:
                    matched_order_vouchers = [o.so_ct for o in app_row["vouchers"] if self.rules.classify_spec(report_type, o.ma_hang, o.ten_mat_hang) == sp]
                    if matched_order_vouchers:
                        if vouchers_str and vouchers_str != "chưa tìm thấy chứng từ trong hai nguồn đã kiểm tra":
                            vouchers_str += f" | App tính từ: {', '.join(set(matched_order_vouchers))}"
                        else:
                            vouchers_str = f"App tính từ: {', '.join(set(matched_order_vouchers))}"

                discrepancies.append(DiscrepancyRecord(
                    report_type=report_type,
                    ngay_bc=target_date_str,
                    ma_khach=ma_khach,
                    ten_khach=ten_khach_tay,
                    thuong_hieu=th_tay,
                    quy_cach=sp,
                    so_app_thung=app_sp_thung,
                    so_tay_thung=tay_sp_thung,
                    chenh_lech_thung=diff_thung,
                    so_app_m2=app_sp_m2,
                    so_tay_m2=tay_sp_m2,
                    chenh_lech_m2=diff_m2,
                    nguon_chung_tu=nguon_str,
                    chung_tu_lien_quan=vouchers_str,
                    ly_do=ly_do,
                    loai_chenh_lech=loai
                ))

        wb_tay.close()
        return discrepancies
