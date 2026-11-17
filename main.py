import sys
import os
import argparse
from datetime import datetime, date
import shutil

# Đảm bảo đường dẫn import
if getattr(sys, 'frozen', False):
    app_dir = os.path.dirname(sys.executable)
    candidate = app_dir
    base_workspace = candidate
    for _ in range(4):
        if (os.path.exists(os.path.join(candidate, "Đồng Nai.xlsx")) or 
            os.path.exists(os.path.join(candidate, "GRANIT.xlsx")) or
            os.path.exists(os.path.join(candidate, "Đồng Nai Bảng kê đơn hàng bán 12.11-2026.xlsx"))):
            base_workspace = candidate
            break
        parent = os.path.dirname(candidate)
        if parent == candidate:
            break
        candidate = parent
    else:
        base_workspace = r"G:\APP LOC DU LIEU\Tổng hợp"
else:
    app_dir = os.path.dirname(os.path.abspath(__file__))
    base_workspace = os.path.dirname(app_dir)

if app_dir not in sys.path:
    sys.path.insert(0, app_dir)

def run_cli_export():
    parser = argparse.ArgumentParser(description="VTC Báo Cáo - Batch Export Mode")
    parser.add_argument("--export", action="store_true", help="Kích hoạt chế độ xuất báo cáo qua dòng lệnh")
    parser.add_argument("--report", default="TRM_TPS", choices=["TCR_TML", "TMR", "TRM_TPS"], help="Loại báo cáo")
    parser.add_argument("--date", default="2026-11-12", help="Ngày báo cáo (YYYY-MM-DD)")
    parser.add_argument("--file-dn", help="Đường dẫn tệp Đồng Nai")
    parser.add_argument("--file-gr", help="Đường dẫn tệp GRANITE")
    parser.add_argument("--file-mau", help="Đường dẫn tệp Excel tổng mẫu")
    parser.add_argument("--output", help="Đường dẫn tệp Excel đích xuất ra")
    parser.add_argument("--only-completed", action="store_true", help="Chỉ lấy đơn hàng Hoàn thành")
    parser.add_argument("--apply-adjustments", action="store_true", help="Áp dụng điều chỉnh chứng từ")

    args, unknown = parser.parse_known_args()
    if not args.export:
        return False

    from vtc_app.reader import ExcelReader
    from vtc_app.rules import RuleManager
    from vtc_app.engine import ReportEngine
    from vtc_app.writer import ExcelWriter

    target_d = datetime.strptime(args.date, "%Y-%m-%d").date()

    dn_path = args.file_dn.strip('"\'') if args.file_dn else os.path.join(base_workspace, "Đồng Nai Bảng kê đơn hàng bán 12.11-2026.xlsx")
    if not os.path.exists(dn_path):
        dn_path = os.path.join(base_workspace, "Đồng Nai.xlsx")

    gr_path = args.file_gr.strip('"\'') if args.file_gr else os.path.join(base_workspace, "GRANITE Bảng kê đơn hàng bán 12.11.xlsx")
    if not os.path.exists(gr_path):
        gr_path = os.path.join(base_workspace, "GRANIT.xlsx")

    default_mau = "TCR+TML.xlsx" if args.report == "TCR_TML" else ("TMR.xlsx" if args.report == "TMR" else "TRM +TPS.xlsx")
    mau_path = args.file_mau.strip('"\'') if args.file_mau else os.path.join(base_workspace, default_mau)
    out_path = args.output.strip('"\'') if args.output else os.path.join(base_workspace, "VTC_Bao_Cao", "test_run_output", default_mau)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    if not os.path.exists(out_path) or os.path.abspath(out_path) != os.path.abspath(mau_path):
        shutil.copy2(mau_path, out_path)

    config_candidates = [
        os.path.join(app_dir, "config"),
        os.path.join(app_dir, "_internal", "config"),
        os.path.join(base_workspace, "VTC_Bao_Cao", "config"),
        os.path.join(base_workspace, "config"),
    ]
    config_dir = config_candidates[0]
    for cand in config_candidates:
        if os.path.exists(cand):
            config_dir = cand
            break

    rm = RuleManager(config_dir)
    engine = ReportEngine(rm)

    tracked_custs = ExcelReader.read_tracked_customers(
        out_path, args.report, target_date=target_d, rule_manager=rm
    )

    orders = []
    if os.path.exists(dn_path):
        orders.extend(ExcelReader.read_orders(dn_path, target_date=target_d, only_completed=args.only_completed))
    if os.path.exists(gr_path):
        orders.extend(ExcelReader.read_orders(gr_path, target_date=target_d, only_completed=args.only_completed))

    res = engine.run_report(
        args.report, orders, tracked_custs, target_date=target_d,
        only_completed=args.only_completed, apply_adjustments=args.apply_adjustments
    )

    backup_dir = os.path.join(os.path.dirname(os.path.abspath(out_path)), "backups")
    ExcelWriter.update_summary_excel(out_path, args.report, res, target_d, backup_dir, rule_manager=rm)

    mode_name = "Chỉ Hoàn thành" if args.only_completed else "Hoàn thành + Duyệt"
    adj_name = "Đã bật điều chỉnh" if args.apply_adjustments else "Điều chỉnh TẮT (Số theo nguồn)"
    msg = (
        f"=== XUẤT THÀNH CÔNG [{args.report}] [{target_d}] [{mode_name}] [{adj_name}] ===\n"
        f"File đích: {out_path}\n"
        f"Tổng thùng: {res['tot_thung']:,.0f}\n"
        f"Tổng m2: {res['tot_m2']:,.2f}\n"
        f"Số đại lý phát sinh Mục 01: {len(res['muc01_rows'])}\n"
    )
    for r in res['muc01_rows']:
        msg += f"  {r['stt']}. [{r['ma_khach']}] {r['ten_khach']} ({r['thuong_hieu']} - {r['nhom']}) -> Thùng: {r['tong_thung']:,.0f}, m2: {r['tong_m2']:,.2f}\n"

    try:
        print(msg)
    except Exception:
        pass

    log_path = os.path.join(os.path.dirname(os.path.abspath(out_path)), "last_export.log")
    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(msg)
    except Exception:
        pass
    return True

def main():
    if "--export" in sys.argv:
        try:
            run_cli_export()
        except Exception as e:
            err_path = os.path.join(app_dir, "export_error.log")
            with open(err_path, "w", encoding="utf-8") as f:
                f.write(str(e))
        sys.exit(0)

    from PySide6.QtWidgets import QApplication
    from vtc_app.ui.main_window import MainWindow

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    window = MainWindow(base_workspace)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
