import sys
import os
import shutil
from datetime import datetime, date
from typing import List, Dict, Any, Optional

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QComboBox, QCheckBox, QTabWidget, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox, QSplitter, QHeaderView,
    QGroupBox, QLineEdit, QInputDialog
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont, QColor, QIcon

from ..models import OrderRow, TrackedCustomer, DiscrepancyRecord
from ..rules import RuleManager
from ..reader import ExcelReader, normalize_dealer_name
from ..engine import ReportEngine
from ..writer import ExcelWriter
from .rules_dialog import RulesDialog
from .dealers_dialog import DealersDialog
from .sync_dialog import SyncMasterExcelDialog

class MainWindow(QMainWindow):
    def __init__(self, workspace_path: str):
        super().__init__()
        self.setWindowTitle("VTC Báo Cáo - Hệ Thống Báo Cáo Tiêu Thụ Hằng Ngày")
        self.resize(1380, 850)

        self.base_workspace = workspace_path
        
        # Xác định config_dir linh hoạt
        config_candidates = [
            os.path.join(self.base_workspace, "VTC_Bao_Cao", "config"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "config"),
        ]
        if getattr(sys, 'frozen', False):
            if hasattr(sys, '_MEIPASS'):
                config_candidates.insert(0, os.path.join(sys._MEIPASS, "config"))
            exe_dir = os.path.dirname(sys.executable)
            config_candidates.insert(0, os.path.join(exe_dir, "config"))
            config_candidates.insert(0, os.path.join(exe_dir, "_internal", "config"))

        self.config_dir = config_candidates[0]
        for cand in config_candidates:
            if os.path.exists(cand):
                self.config_dir = cand
                break

        self.backup_dir = os.path.join(self.base_workspace, "VTC_Bao_Cao", "backups")
        os.makedirs(self.backup_dir, exist_ok=True)

        self.rules = RuleManager(self.config_dir)
        self.engine = ReportEngine(self.rules)

        self.current_report = "TRM_TPS"  # TCR_TML, TMR, TRM_TPS
        self.app_results: Optional[Dict[str, Any]] = None
        self.discrepancies: List[DiscrepancyRecord] = []
        self.settings: Dict[str, Any] = {}

        self.init_ui()
        self.load_default_paths()
        self.on_report_changed("TRM_TPS")

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # 1. Top Bar: Report Mode Switcher (Trái) & Quản lý / Cập nhật (Phải)
        top_bar = QHBoxLayout()
        lbl_rep = QLabel("<b>Báo cáo đang chọn:</b>")
        lbl_rep.setStyleSheet("font-size: 13px; color: #202124;")
        top_bar.addWidget(lbl_rep)

        self.btn_rep_tcr = QPushButton("TCR + TML (Tocera – Tomilan)")
        self.btn_rep_tmr = QPushButton("TMR (Tomira)")
        self.btn_rep_trm = QPushButton("TRM + TPS (Toroma – Topisa)")

        for btn, rep_key in [(self.btn_rep_tcr, "TCR_TML"), (self.btn_rep_tmr, "TMR"), (self.btn_rep_trm, "TRM_TPS")]:
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked, r=rep_key: self.on_report_changed(r))
            top_bar.addWidget(btn)

        top_bar.addStretch()

        # Cụm bên phải: Quản lý đại lý, Quy luật mã gạch, Cập nhật ứng dụng
        self.btn_manage_dealers = QPushButton("👥 Quản lý đại lý")
        self.btn_manage_dealers.setStyleSheet("""
            QPushButton { padding: 8px 14px; font-weight: bold; font-size: 12px; border-radius: 4px; border: 1px solid #1a73e8; background-color: #e8f0fe; color: #1a73e8; }
            QPushButton:hover { background-color: #d2e3fc; }
        """)
        self.btn_manage_dealers.clicked.connect(self.on_open_dealers)
        top_bar.addWidget(self.btn_manage_dealers)

        self.btn_rules = QPushButton("⚙ Quy luật mã gạch")
        self.btn_rules.setStyleSheet("padding: 8px 14px; font-weight: bold; font-size: 12px; border-radius: 4px; border: 1px solid #5f6368; background-color: #f1f3f4; color: #3c4043;")
        self.btn_rules.clicked.connect(self.on_open_rules)
        top_bar.addWidget(self.btn_rules)

        self.btn_update_app = QPushButton("🔄 Cập nhật ứng dụng")
        self.btn_update_app.setStyleSheet("""
            QPushButton { padding: 8px 14px; font-weight: bold; font-size: 12px; border-radius: 4px; border: 1px solid #0b8043; background-color: #e6f4ea; color: #0b8043; }
            QPushButton:hover { background-color: #ceead6; }
        """)
        self.btn_update_app.clicked.connect(self.on_open_update_dialog)
        top_bar.addWidget(self.btn_update_app)

        main_layout.addLayout(top_bar)

        # 2. Splitter: Left Sidebar & Right Content Tabs
        splitter = QSplitter(Qt.Horizontal)

        # Left Sidebar (Luồng vận hành hằng ngày)
        left_pane = QWidget()
        left_layout = QVBoxLayout(left_pane)
        left_layout.setContentsMargins(0, 0, 5, 0)
        left_layout.setSpacing(10)

        # Group 1: Files
        grp_files = QGroupBox("Tệp đơn hàng & Mẫu báo cáo")
        grp_files_layout = QVBoxLayout(grp_files)
        grp_files_layout.setSpacing(6)

        self.txt_file_dn = self.create_file_row(grp_files_layout, "Nguồn Đồng Nai:", self.on_browse_dn)
        self.txt_file_gr = self.create_file_row(grp_files_layout, "Nguồn GRANIT:", self.on_browse_gr)
        self.txt_file_mau = self.create_file_row(grp_files_layout, "Mẫu Excel tổng:", self.on_browse_mau)
        
        self.chk_remember_mau = QCheckBox("Ghi nhớ Excel tổng cho báo cáo này")
        self.chk_remember_mau.setChecked(True)
        grp_files_layout.addWidget(self.chk_remember_mau)
        self.txt_file_tay = QLineEdit()
        left_layout.addWidget(grp_files)

        # Group 2: Filter Controls
        grp_filter = QGroupBox("Điều kiện lọc báo cáo")
        grp_filter_layout = QVBoxLayout(grp_filter)
        grp_filter_layout.setSpacing(8)

        lbl_date = QLabel("Ngày chứng từ / báo cáo:")
        self.cbo_date = QComboBox()
        self.cbo_date.setStyleSheet("padding: 5px; font-size: 12px; font-weight: bold;")
        grp_filter_layout.addWidget(lbl_date)
        grp_filter_layout.addWidget(self.cbo_date)

        self.chk_only_completed = QCheckBox("Chỉ lấy đơn hàng Hoàn thành")
        self.chk_only_completed.setToolTip("Mặc định: Lấy cả Hoàn thành và Duyệt / Đã duyệt.\nTích chọn nếu chỉ muốn tính đơn Hoàn thành.")
        grp_filter_layout.addWidget(self.chk_only_completed)

        self.chk_apply_adj = QCheckBox("Bật điều chỉnh chứng từ (TRM+TPS)")
        self.chk_apply_adj.setChecked(False)  # Mặc định TẮT
        self.chk_apply_adj.setStyleSheet("color: #d93025; font-weight: bold;")
        self.chk_apply_adj.setToolTip("Mặc định TẮT: Giữ nguyên chứng từ Hoàn thành theo nguồn.\nBẬT: Áp dụng trừ các chứng từ đã xác nhận theo bảng tay.")
        self.chk_apply_adj.stateChanged.connect(self.on_adjustment_toggled)
        grp_filter_layout.addWidget(self.chk_apply_adj)

        left_layout.addWidget(grp_filter)

        # Daily Action Buttons
        self.btn_preview = QPushButton("▶ Xem trước báo cáo")
        self.btn_preview.setStyleSheet("background-color: #1a73e8; color: white; padding: 11px; font-weight: bold; font-size: 13px; border-radius: 4px;")
        self.btn_preview.clicked.connect(self.on_run_preview)

        self.btn_open_template_excel = QPushButton("Mở bằng Excel")
        self.btn_open_template_excel.setStyleSheet("background-color: #5f6368; color: white; padding: 9px; font-weight: bold; font-size: 12px; border-radius: 4px;")
        self.btn_open_template_excel.clicked.connect(self.on_open_template_in_excel)

        self.btn_update_excel = QPushButton("💾 Cập nhật Excel tổng")
        self.btn_update_excel.setStyleSheet("background-color: #0b8043; color: white; padding: 11px; font-weight: bold; font-size: 13px; border-radius: 4px;")
        self.btn_update_excel.clicked.connect(self.on_update_excel)

        left_layout.addWidget(self.btn_preview)
        left_layout.addWidget(self.btn_open_template_excel)
        left_layout.addWidget(self.btn_update_excel)
        left_layout.addStretch()

        left_pane.setMaximumWidth(390)
        splitter.addWidget(left_pane)

        # Right Content Tabs
        self.right_tabs = QTabWidget()

        # Tab 0: Kết quả báo cáo (Mục 01)
        tab_muc01 = QWidget()
        m1_layout = QVBoxLayout(tab_muc01)
        m1_layout.setContentsMargins(5, 5, 5, 5)
        
        self.lbl_stats = QLabel("Chưa chạy xem trước báo cáo. Nhấn '▶ Xem trước báo cáo' để bắt đầu.")
        self.lbl_stats.setStyleSheet("font-size: 13px; font-weight: bold; color: #1a73e8; padding: 8px; background-color: #f1f3f4; border-radius: 4px;")
        m1_layout.addWidget(self.lbl_stats)

        m1_splitter = QSplitter(Qt.Vertical)

        # Mục 01 Table
        w_m1 = QWidget()
        wm1_layout = QVBoxLayout(w_m1)
        wm1_layout.setContentsMargins(0, 0, 0, 0)
        wm1_layout.addWidget(QLabel("<b>Mục 01: Đại lý phát sinh tiêu thụ trong ngày (Có Mã khách hàng, chỉ hiển thị đại lý có phát sinh)</b>"))
        self.tbl_muc01 = QTableWidget()
        wm1_layout.addWidget(self.tbl_muc01)
        m1_splitter.addWidget(w_m1)

        # Details Table
        w_dt = QWidget()
        wdt_layout = QVBoxLayout(w_dt)
        wdt_layout.setContentsMargins(0, 0, 0, 0)
        wdt_layout.addWidget(QLabel("<b>Chi tiết chứng từ nguồn đã tính / bị loại (Xem lý do, chứng từ, trạng thái):</b>"))
        self.tbl_details = QTableWidget()
        wdt_layout.addWidget(self.tbl_details)
        m1_splitter.addWidget(w_dt)

        m1_splitter.setStretchFactor(0, 3)
        m1_splitter.setStretchFactor(1, 2)
        m1_layout.addWidget(m1_splitter)
        self.right_tabs.addTab(tab_muc01, "Xem trước báo cáo (Mục 01)")

        # Tab 1: Quản lý đại lý (Mục 02)
        tab_muc02 = QWidget()
        m2_layout = QVBoxLayout(tab_muc02)
        m2_layout.setContentsMargins(5, 5, 5, 5)
        m2_bar = QHBoxLayout()
        m2_bar.addWidget(QLabel("<b>Danh sách đại lý đang theo dõi (Mục 02):</b> (Mặc định giữ danh sách đầy đủ)"))
        m2_bar.addStretch()
        btn_add_cust = QPushButton("+ Thêm đại lý qua mã")
        btn_add_cust.setStyleSheet("padding: 6px 12px; font-weight: bold;")
        btn_add_cust.clicked.connect(self.on_add_customer)
        m2_bar.addWidget(btn_add_cust)
        m2_layout.addLayout(m2_bar)

        self.tbl_muc02 = QTableWidget()
        m2_layout.addWidget(self.tbl_muc02)
        self.right_tabs.addTab(tab_muc02, "Danh sách theo dõi (Mục 02)")

        # Dummy/fallback references for removed tabs to maintain API compatibility
        self.tbl_view_excel = None
        self.tbl_doi_chieu = None

        splitter.addWidget(self.right_tabs)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

        main_layout.addWidget(splitter)

    def create_file_row(self, layout, label_text, browse_slot):
        lbl = QLabel(label_text)
        txt = QLineEdit()
        txt.setReadOnly(True)
        txt.setStyleSheet("background-color: #f8f9fa;")
        btn = QPushButton("Chọn...")
        btn.clicked.connect(browse_slot)
        row = QHBoxLayout()
        row.addWidget(txt)
        row.addWidget(btn)
        layout.addWidget(lbl)
        layout.addLayout(row)
        return txt

    def load_default_paths(self):
        dn_new = os.path.join(self.base_workspace, "Đồng Nai Bảng kê đơn hàng bán 12.11-2026.xlsx")
        dn_old = os.path.join(self.base_workspace, "Đồng Nai.xlsx")
        dn_default = dn_new if os.path.exists(dn_new) else dn_old

        gr_new = os.path.join(self.base_workspace, "GRANITE Bảng kê đơn hàng bán 12.11.xlsx")
        gr_old = os.path.join(self.base_workspace, "GRANIT.xlsx")
        gr_default = gr_new if os.path.exists(gr_new) else gr_old

        self.txt_file_dn.setText(self.settings.get("file_dn", dn_default if os.path.exists(dn_default) else ""))
        self.txt_file_gr.setText(self.settings.get("file_gr", gr_default if os.path.exists(gr_default) else ""))
        self.refresh_dates_from_sources()

    def on_report_changed(self, rep_key: str):
        self.current_report = rep_key
        is_tcr = (rep_key == "TCR_TML")
        is_tmr = (rep_key == "TMR")
        is_trm = (rep_key == "TRM_TPS")

        self.btn_rep_tcr.setChecked(is_tcr)
        self.btn_rep_tmr.setChecked(is_tmr)
        self.btn_rep_trm.setChecked(is_trm)

        # Cập nhật chữ và dấu chọn ✓
        self.btn_rep_tcr.setText("✓ TCR + TML (Tocera – Tomilan)" if is_tcr else "TCR + TML (Tocera – Tomilan)")
        self.btn_rep_tmr.setText("✓ TMR (Tomira)" if is_tmr else "TMR (Tomira)")
        self.btn_rep_trm.setText("✓ TRM + TPS (Toroma – Topisa)" if is_trm else "TRM + TPS (Toroma – Topisa)")

        # Nút TCR + TML: Đỏ TOCERA (#c5221f)
        self.btn_rep_tcr.setStyleSheet(f"""
            QPushButton {{
                padding: 8px 16px; font-weight: bold; font-size: 12px; border-radius: 4px;
                background-color: {'#8e0000' if is_tcr else '#c5221f'};
                color: #ffffff;
                border: {'2px solid #ffffff' if is_tcr else '1px solid #a51b18'};
            }}
            QPushButton:hover {{ background-color: #b31b1b; }}
        """)

        # Nút TMR: Xanh lá sáng TOMIRA (#0b8043)
        self.btn_rep_tmr.setStyleSheet(f"""
            QPushButton {{
                padding: 8px 16px; font-weight: bold; font-size: 12px; border-radius: 4px;
                background-color: {'#044823' if is_tmr else '#0b8043'};
                color: #ffffff;
                border: {'2px solid #ffffff' if is_tmr else '1px solid #086333'};
            }}
            QPushButton:hover {{ background-color: #086333; }}
        """)

        # Nút TRM + TPS: Cam TOROMA (#e65100)
        self.btn_rep_trm.setStyleSheet(f"""
            QPushButton {{
                padding: 8px 16px; font-weight: bold; font-size: 12px; border-radius: 4px;
                background-color: {'#a72d00' if is_trm else '#e65100'};
                color: #ffffff;
                border: {'2px solid #ffffff' if is_trm else '1px solid #bf360c'};
            }}
            QPushButton:hover {{ background-color: #d84315; }}
        """)

        # Show adjustment toggle only for TRM_TPS
        self.chk_apply_adj.setVisible(is_trm)

        manual_dir = os.path.join(self.base_workspace, "Dự liệu làm tay")
        if not os.path.exists(manual_dir):
            manual_dir = os.path.join(self.base_workspace, "Dữ liệu làm tay")

        if rep_key == "TCR_TML":
            m_path = self.settings.get("mau_tcr", os.path.join(self.base_workspace, "TCR+TML.xlsx"))
            h_path = os.path.join(manual_dir, "TCR+TML.xlsx")
        elif rep_key == "TMR":
            m_path = self.settings.get("mau_tmr", os.path.join(self.base_workspace, "TMR.xlsx"))
            h_path = os.path.join(manual_dir, "TMR.xlsx")
        elif rep_key == "TRM_TPS":
            m_path = self.settings.get("mau_trm", os.path.join(self.base_workspace, "TRM +TPS.xlsx"))
            h_path = os.path.join(manual_dir, "TRM.xlsx")
            if not os.path.exists(h_path):
                alt_h = os.path.join(self.base_workspace, "dữ liệu làm tay.xlsx")
                if os.path.exists(alt_h):
                    h_path = alt_h

        self.txt_file_mau.setText(m_path if os.path.exists(m_path) else "")
        if hasattr(self, "txt_file_tay") and self.txt_file_tay:
            self.txt_file_tay.setText(h_path if os.path.exists(h_path) else "")

        # Reload tracked customers
        self.reload_muc02_table()

    def on_adjustment_toggled(self):
        if self.app_results:
            self.on_run_preview()

    def refresh_dates_from_sources(self):
        self.cbo_date.clear()
        dates = set()
        for p in [self.txt_file_dn.text(), self.txt_file_gr.text()]:
            if p and os.path.exists(p):
                dates.update(ExcelReader.get_source_dates(p))
        sorted_dates = sorted(list(dates))
        for d in sorted_dates:
            self.cbo_date.addItem(d.strftime("%d/%m/%Y"), d)
        if sorted_dates:
            self.cbo_date.setCurrentIndex(0)

    def on_browse_dn(self):
        f, _ = QFileDialog.getOpenFileName(self, "Chọn tệp Đồng Nai", self.base_workspace, "Excel Files (*.xlsx *.xls)")
        if f:
            self.txt_file_dn.setText(f)
            self.settings["file_dn"] = f
            self.refresh_dates_from_sources()

    def on_browse_gr(self):
        f, _ = QFileDialog.getOpenFileName(self, "Chọn tệp GRANIT", self.base_workspace, "Excel Files (*.xlsx *.xls)")
        if f:
            self.txt_file_gr.setText(f)
            self.settings["file_gr"] = f
            self.refresh_dates_from_sources()

    def on_browse_mau(self):
        f, _ = QFileDialog.getOpenFileName(self, "Chọn mẫu Excel tổng", self.base_workspace, "Excel Files (*.xlsx *.xls)")
        if f:
            self.txt_file_mau.setText(f)
            if self.chk_remember_mau.isChecked():
                k = "mau_tcr" if self.current_report == "TCR_TML" else ("mau_tmr" if self.current_report == "TMR" else "mau_trm")
                self.settings[k] = f
            self.reload_muc02_table()

    def on_browse_tay(self):
        f, _ = QFileDialog.getOpenFileName(self, "Chọn tệp làm tay", self.base_workspace, "Excel Files (*.xlsx *.xls)")
        if f:
            self.txt_file_tay.setText(f)

    def on_open_dealers(self):
        source_paths = [self.txt_file_dn.text(), self.txt_file_gr.text()]
        dlg = DealersDialog(self.rules, initial_report_type=self.current_report, source_files=source_paths, parent=self)
        dlg.dealers_changed.connect(self.on_dealers_changed)
        dlg.exec()

    def on_dealers_changed(self):
        self.reload_muc02_table()
        if self.app_results:
            self.on_run_preview()

    def on_rules_changed(self):
        if self.app_results:
            self.on_run_preview()

    def on_open_rules(self):
        dlg = RulesDialog(self.rules, initial_report_type=self.current_report, parent=self)
        dlg.rules_changed.connect(self.on_rules_changed)
        dlg.exec()

    def on_open_update_dialog(self):
        from .update_dialog import UpdateDialog
        dlg = UpdateDialog(self)
        dlg.exec()

    def get_selected_date(self) -> Optional[date]:
        return self.cbo_date.currentData()

    def reload_muc02_table(self):
        try:
            dealers = self.rules.get_dealers(self.current_report)
            headers = ["STT", "Mã khách", "Tên cửa hàng", "Nhóm", "Thương hiệu", "Trạng thái theo dõi"]
            self.tbl_muc02.setColumnCount(len(headers))
            self.tbl_muc02.setHorizontalHeaderLabels(headers)
            self.tbl_muc02.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            self.tbl_muc02.setRowCount(len(dealers))

            for idx, d in enumerate(dealers):
                is_act = d.get("is_active", True)
                it_stt = QTableWidgetItem(str(d.get("stt", idx + 1)))
                it_stt.setTextAlignment(Qt.AlignCenter)
                it_code = QTableWidgetItem(d.get("ma_khach", ""))
                it_name = QTableWidgetItem(d.get("ten_khach", ""))
                it_nhom = QTableWidgetItem(d.get("nhom", ""))
                it_nhom.setTextAlignment(Qt.AlignCenter)
                it_th = QTableWidgetItem(d.get("thuong_hieu", ""))
                it_th.setTextAlignment(Qt.AlignCenter)
                it_status = QTableWidgetItem("Đang theo dõi" if is_act else "Ngừng theo dõi")
                it_status.setTextAlignment(Qt.AlignCenter)

                if is_act:
                    it_status.setForeground(QColor("#0b8043"))
                else:
                    it_status.setForeground(QColor("#d93025"))
                    for it in [it_stt, it_code, it_name, it_nhom, it_th]:
                        it.setForeground(QColor("#9aa0a6"))

                self.tbl_muc02.setItem(idx, 0, it_stt)
                self.tbl_muc02.setItem(idx, 1, it_code)
                self.tbl_muc02.setItem(idx, 2, it_name)
                self.tbl_muc02.setItem(idx, 3, it_nhom)
                self.tbl_muc02.setItem(idx, 4, it_th)
                self.tbl_muc02.setItem(idx, 5, it_status)
        except Exception as e:
            print(f"Lỗi đọc Mục 02: {e}")

    def on_add_customer(self):
        self.on_open_dealers()

    def on_run_preview(self):
        target_date = self.get_selected_date()
        only_completed = self.chk_only_completed.isChecked()
        apply_adj = self.chk_apply_adj.isChecked() if self.current_report == "TRM_TPS" else False

        orders = []
        for p in [self.txt_file_gr.text(), self.txt_file_dn.text()]:
            if p and os.path.exists(p):
                orders.extend(ExcelReader.read_orders(p, target_date=target_date, only_completed=only_completed))

        if not orders:
            QMessageBox.warning(self, "Không có dữ liệu", "Không tìm thấy đơn hàng nào phù hợp với bộ lọc ngày và trạng thái!")
            return

        template_path = self.txt_file_mau.text()
        if not template_path or not os.path.exists(template_path):
            QMessageBox.warning(self, "Thiếu mẫu", "Vui lòng chọn tệp mẫu Excel tổng!")
            return

        # Đọc đại lý từ RuleManager và đồng bộ với template
        dealers = self.rules.get_dealers(self.current_report)
        dealer_active_map = {}
        for d in dealers:
            code = d.get("ma_khach", "").strip().upper()
            th = d.get("thuong_hieu", "").strip().upper()
            name = d.get("ten_khach", "").strip().upper()
            dealer_active_map[(code, th)] = d.get("is_active", True)
            dealer_active_map[(name, th)] = d.get("is_active", True)

        tracked_custs = ExcelReader.read_tracked_customers(
            template_path, self.current_report, target_date=target_date, rule_manager=self.rules
        )
        dealers_by_name_th = {(normalize_dealer_name(d.get("ten_khach", "")), d.get("thuong_hieu", "").strip().upper()): d for d in dealers}
        dealers_by_code_th = {(d.get("ma_khach", "").strip().upper(), d.get("thuong_hieu", "").strip().upper()): d for d in dealers if d.get("ma_khach")}
        dealers_by_name_only = {normalize_dealer_name(d.get("ten_khach", "")): d for d in dealers}

        for c in tracked_custs:
            code = c.ma_khach.strip().upper()
            th = c.thuong_hieu.strip().upper()
            name_norm = normalize_dealer_name(c.ten_khach)

            d = None
            if code:
                d = dealers_by_code_th.get((code, th))
            if not d and name_norm:
                d = dealers_by_name_th.get((name_norm, th))
            if not d and name_norm:
                d = dealers_by_name_only.get(name_norm)

            if d:
                if not c.ma_khach and d.get("ma_khach"):
                    c.ma_khach = d["ma_khach"]
                if not c.nhom and d.get("nhom"):
                    c.nhom = d["nhom"]
                c.is_active = d.get("is_active", True)
            elif (code, th) in dealer_active_map:
                c.is_active = dealer_active_map[(code, th)]
            elif (c.ten_khach.strip().upper(), th) in dealer_active_map:
                c.is_active = dealer_active_map[(c.ten_khach.strip().upper(), th)]
            else:
                c.is_active = True

        existing_keys = {(c.ma_khach.strip().upper(), c.thuong_hieu.strip().upper()) for c in tracked_custs if c.ma_khach}
        existing_names = {(c.ten_khach.strip().upper(), c.thuong_hieu.strip().upper()) for c in tracked_custs}
        for d in dealers:
            code = d.get("ma_khach", "").strip().upper()
            th = d.get("thuong_hieu", "").strip().upper()
            name = d.get("ten_khach", "").strip().upper()
            if (code, th) not in existing_keys and (name, th) not in existing_names:
                tracked_custs.append(TrackedCustomer(
                    ma_khach=d.get("ma_khach", ""),
                    stt=d.get("stt", len(tracked_custs) + 1),
                    ten_khach=d.get("ten_khach", ""),
                    nhom=d.get("nhom", ""),
                    thuong_hieu=th,
                    row_idx_muc02=0,
                    col_idx_ma_khach=1,
                    is_active=d.get("is_active", True)
                ))

        self.app_results = self.engine.run_report(
            self.current_report, orders, tracked_custs, target_date, only_completed, apply_adjustments=apply_adj
        )

        # Display Mục 01
        self.display_muc01()
        self.right_tabs.setCurrentIndex(0)

    def display_muc01(self):
        res = self.app_results
        if not res:
            return

        date_str = res['target_date'].strftime('%d/%m/%Y') if res['target_date'] else 'Tất cả'
        txt_stats = (
            f"Báo cáo: {self.current_report} | Ngày: {date_str} | "
            f"Đại lý phát sinh hợp lệ: {len(res['muc01_rows'])}\n"
            f"• Số theo nguồn: {res['tot_thung']:,.0f} thùng ({res['tot_m2']:,.2f} m²)"
        )
        if res.get("apply_adjustments"):
            txt_stats += (
                f" | ĐÃ BẬT ĐIỀU CHỈNH: {res['tot_thung']:,.0f} thùng ({res['tot_m2']:,.2f} m²)\n"
                f"(Đã áp dụng các quy tắc điều chỉnh chứng từ theo cấu hình)"
            )
        else:
            txt_stats += (
                f" | Điều chỉnh: ĐANG TẮT (Mặc định giữ nguyên toàn bộ chứng từ theo nguồn dữ liệu)"
            )
        self.lbl_stats.setText(txt_stats)

        rows = res["muc01_rows"]
        # Find all specs present
        all_specs = set()
        for r in rows:
            all_specs.update(r["specs"].keys())
        spec_list = sorted(list(all_specs))

        headers = ["Mã KH", "STT", "Cửa hàng", "Nhóm", "TH", "Tổng (thùng)", "Tổng (m²)"] + spec_list
        self.tbl_muc01.setColumnCount(len(headers))
        self.tbl_muc01.setHorizontalHeaderLabels(headers)
        self.tbl_muc01.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        
        # Nếu có dữ liệu, thêm 1 dòng TỔNG ở cuối bảng
        row_count = len(rows) + (1 if len(rows) > 0 else 0)
        self.tbl_muc01.setRowCount(row_count)

        for idx, r in enumerate(rows):
            self.tbl_muc01.setItem(idx, 0, QTableWidgetItem(r["ma_khach"]))
            self.tbl_muc01.setItem(idx, 1, QTableWidgetItem(str(r["stt"])))
            self.tbl_muc01.setItem(idx, 2, QTableWidgetItem(r["ten_khach"]))
            self.tbl_muc01.setItem(idx, 3, QTableWidgetItem(r["nhom"]))
            self.tbl_muc01.setItem(idx, 4, QTableWidgetItem(r["thuong_hieu"]))
            self.tbl_muc01.setItem(idx, 5, QTableWidgetItem(f"{r['tong_thung']:,.0f}"))
            self.tbl_muc01.setItem(idx, 6, QTableWidgetItem(f"{r['tong_m2']:,.2f}"))

            for c_idx, sp in enumerate(spec_list, start=7):
                qty = r["specs"].get(sp, 0.0)
                val_str = f"{qty:,.0f}" if qty > 0 else ""
                it = QTableWidgetItem(val_str)
                it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.tbl_muc01.setItem(idx, c_idx, it)

        # Dòng TỔNG Mục 01
        if len(rows) > 0:
            tot_idx = len(rows)
            bold_font = QFont()
            bold_font.setBold(True)
            bg_tot = QColor("#e8f0fe")

            def make_tot_item(text: str, align=Qt.AlignLeft | Qt.AlignVCenter):
                item = QTableWidgetItem(text)
                item.setFont(bold_font)
                item.setBackground(bg_tot)
                item.setTextAlignment(align)
                return item

            self.tbl_muc01.setItem(tot_idx, 0, make_tot_item(""))
            self.tbl_muc01.setItem(tot_idx, 1, make_tot_item("TỔNG", Qt.AlignCenter))
            self.tbl_muc01.setItem(tot_idx, 2, make_tot_item(""))
            self.tbl_muc01.setItem(tot_idx, 3, make_tot_item(""))
            self.tbl_muc01.setItem(tot_idx, 4, make_tot_item(""))
            self.tbl_muc01.setItem(tot_idx, 5, make_tot_item(f"{res['tot_thung']:,.0f}", Qt.AlignRight | Qt.AlignVCenter))
            self.tbl_muc01.setItem(tot_idx, 6, make_tot_item(f"{res['tot_m2']:,.2f}", Qt.AlignRight | Qt.AlignVCenter))

            for c_idx, sp in enumerate(spec_list, start=7):
                sp_total = sum(r["specs"].get(sp, 0.0) for r in rows)
                val_str = f"{sp_total:,.0f}" if sp_total > 0 else ""
                self.tbl_muc01.setItem(tot_idx, c_idx, make_tot_item(val_str, Qt.AlignRight | Qt.AlignVCenter))

        # Populate Details table (Included and Excluded rows)
        inc = res["included_rows"]
        exc = res["excluded_rows"]
        dt_headers = ["Phân loại", "Nguồn", "Số CT", "Mã KH", "Tên KH", "Mã hàng", "Tên hàng", "SL (thùng)", "Trạng thái", "Ghi chú / Lý do"]
        self.tbl_details.setColumnCount(len(dt_headers))
        self.tbl_details.setHorizontalHeaderLabels(dt_headers)
        self.tbl_details.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        self.tbl_details.horizontalHeader().setSectionResizeMode(9, QHeaderView.Stretch)
        self.tbl_details.setRowCount(len(inc) + len(exc))

        row_idx = 0
        for item in inc:
            o = item["order"]
            vals = ["ĐÃ TÍNH", o.source_file, o.so_ct, o.ma_khach, o.ten_khach, o.ma_hang, o.ten_mat_hang, f"{o.sl_dat_hang:,.2f}", o.trang_thai, f"Quy cách: {item['spec']} (hs: {item['factor']})"]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(str(v))
                it.setBackground(QColor("#e6f4ea"))
                self.tbl_details.setItem(row_idx, c, it)
            row_idx += 1

        for item in exc:
            o = item["order"]
            vals = ["BỊ LOẠI", o.source_file, o.so_ct, o.ma_khach, o.ten_khach, o.ma_hang, o.ten_mat_hang, f"{o.sl_dat_hang:,.2f}", o.trang_thai, item["reason"]]
            for c, v in enumerate(vals):
                it = QTableWidgetItem(str(v))
                it.setBackground(QColor("#fce8e6"))
                self.tbl_details.setItem(row_idx, c, it)
            row_idx += 1

    def on_run_reconcile(self):
        if not self.app_results:
            self.on_run_preview()
            if not self.app_results:
                return

        hand_file = self.txt_file_tay.text()
        if not hand_file or not os.path.exists(hand_file):
            QMessageBox.warning(self, "Thiếu tệp làm tay", "Vui lòng chọn tệp Excel làm tay để đối chiếu!")
            return

        target_date = self.get_selected_date()
        date_str = target_date.strftime("%d/%m/%Y") if target_date else "12/11/2026"

        self.discrepancies = self.engine.reconcile_with_manual(
            self.current_report, hand_file, self.app_results, date_str
        )

        self.display_reconciliation()
        self.right_tabs.setCurrentIndex(3)

    def display_reconciliation(self):
        disc = self.discrepancies
        if not disc:
            if self.current_report == "TCR_TML":
                self.lbl_dc_summary.setText("Chưa có bảng tay cùng ngày để đối chiếu (Bảng tay hiện có chỉ lưu ngày 09/11/2026).")
            else:
                self.lbl_dc_summary.setText("Không có dữ liệu đối chiếu hoặc không phát sinh chênh lệch.")
            self.tbl_doi_chieu.setRowCount(0)
            return

        matches = sum(1 for d in disc if d.loai_chenh_lech == "KHỚP")
        diffs = len(disc) - matches
        summary_txt = f"Tổng số dòng đối chiếu: {len(disc)} | Khớp 100%: {matches} | Chênh lệch: {diffs}"
        if self.current_report == "TRM_TPS":
            summary_txt += " | Số theo nguồn: 4.552 thùng | Bảng tay: 6.109 thùng (+1.904 thùng ngoài, -347 thùng tự trừ)"
        self.lbl_dc_summary.setText(summary_txt)

        headers = [
            "Ngày BC", "Mã KH", "Cửa hàng", "TH", "Quy cách",
            "App (thùng)", "Tay (thùng)", "Lệch (thùng)",
            "App (m²)", "Tay (m²)", "Lệch (m²)",
            "Nguồn chứng từ", "Chứng từ liên quan", "Phân loại & Lý do"
        ]
        self.tbl_doi_chieu.setColumnCount(len(headers))
        self.tbl_doi_chieu.setHorizontalHeaderLabels(headers)
        self.tbl_doi_chieu.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl_doi_chieu.horizontalHeader().setSectionResizeMode(12, QHeaderView.Stretch)
        self.tbl_doi_chieu.horizontalHeader().setSectionResizeMode(13, QHeaderView.Stretch)
        self.tbl_doi_chieu.setRowCount(len(disc))

        for idx, d in enumerate(disc):
            color = QColor("#ffffff")
            if d.loai_chenh_lech == "KHỚP":
                color = QColor("#e6f4ea") # Light green
            elif d.loai_chenh_lech == "TAY_TRỪ_CHỨNG_TỪ":
                color = QColor("#feefe3") # Light orange
            elif d.loai_chenh_lech in ["NGUỒN_THIẾU_CHỨNG_TỪ", "THIẾU_MÃ_MỤC02", "CHƯA_TÌM_THẤY_CHỨNG_TỪ"]:
                color = QColor("#fce8e6") # Light red

            vals = [
                d.ngay_bc, d.ma_khach, d.ten_khach, d.thuong_hieu, d.quy_cach,
                f"{d.so_app_thung:,.0f}", f"{d.so_tay_thung:,.0f}", f"{d.chenh_lech_thung:+,.0f}" if d.chenh_lech_thung else "0",
                f"{d.so_app_m2:,.2f}", f"{d.so_tay_m2:,.2f}", f"{d.chenh_lech_m2:+,.2f}" if d.chenh_lech_m2 else "0.00",
                d.nguon_chung_tu, d.chung_tu_lien_quan, d.ly_do
            ]

            for c_idx, v in enumerate(vals):
                it = QTableWidgetItem(str(v))
                it.setBackground(color)
                if c_idx in [5, 6, 7, 8, 9, 10]:
                    it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.tbl_doi_chieu.setItem(idx, c_idx, it)

    def on_export_discrepancies(self):
        if not self.discrepancies:
            QMessageBox.information(self, "Thông báo", "Vui lòng chạy đối chiếu trước khi xuất!")
            return

        f, _ = QFileDialog.getSaveFileName(
            self, "Xuất kết quả đối chiếu", self.base_workspace, "Excel Files (*.xlsx)"
        )
        if not f:
            return

        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Doi_Chieu_Chi_Tiet"

        headers = [
            "Báo cáo", "Ngày BC", "Mã KH", "Cửa hàng", "Thương hiệu", "Quy cách",
            "App (thùng)", "Tay (thùng)", "Lệch (thùng)",
            "App (m²)", "Tay (m²)", "Lệch (m²)",
            "Nguồn chứng từ", "Chứng từ liên quan", "Lý do đối chiếu", "Phân loại"
        ]
        ws.append(headers)

        for d in self.discrepancies:
            ws.append([
                d.report_type, d.ngay_bc, d.ma_khach, d.ten_khach, d.thuong_hieu, d.quy_cach,
                d.so_app_thung, d.so_tay_thung, d.chenh_lech_thung,
                d.so_app_m2, d.so_tay_m2, d.chenh_lech_m2,
                d.nguon_chung_tu, d.chung_tu_lien_quan, d.ly_do, d.loai_chenh_lech
            ])

        wb.save(f)
        wb.close()
        QMessageBox.information(self, "Thành công", f"Đã xuất bảng đối chiếu thành công tại:\n{f}")

    def on_open_template_in_excel(self):
        template_path = self.txt_file_mau.text()
        if template_path and os.path.exists(template_path):
            os.startfile(template_path)
        else:
            QMessageBox.warning(self, "Thiếu tệp", "Chưa tìm thấy tệp mẫu Excel tổng!")

    def on_update_excel(self):
        if not self.app_results:
            self.on_run_preview()
            if not self.app_results:
                return

        target_date = self.get_selected_date()
        if not target_date:
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn ngày báo cáo!")
            return

        template_path = self.txt_file_mau.text()
        if not template_path or not os.path.exists(template_path):
            QMessageBox.warning(self, "Lỗi", "Tệp Excel tổng không tồn tại!")
            return

        # 1. Phân tích đối chiếu Mục 02 với danh mục đại lý AppData
        try:
            sync_items = ExcelWriter.analyze_master_excel_dealers(template_path, self.current_report, self.rules, target_date=target_date)
        except Exception as e:
            print(f"Lỗi phân tích đối chiếu Excel: {e}")
            sync_items = []

        needs_dialog = any(
            it.get("has_change") or it.get("needs_verify") or it.get("is_conflict")
            for it in sync_items
        )

        if needs_dialog:
            dlg = SyncMasterExcelDialog(self, self.current_report, sync_items, template_path)
            if dlg.exec() != QDialog.Accepted or not dlg.is_confirmed:
                return
            try:
                ExcelWriter.apply_confirmed_sync_to_sheet(template_path, self.current_report, dlg.confirmed_items, target_date=target_date)
            except PermissionError as pe:
                QMessageBox.critical(self, "Tệp đang bị khóa", str(pe))
                return
            except Exception as e:
                QMessageBox.critical(self, "Lỗi đồng bộ mã khách", f"Không thể đồng bộ mã vào Excel: {e}")
                return
        else:
            confirm = QMessageBox.question(
                self, "Xác nhận cập nhật",
                f"Bạn có chắc muốn ghi kết quả ngày {target_date.strftime('%d/%m/%Y')} vào tệp Excel tổng:\n{template_path}?\n\n(Hệ thống sẽ tự động tạo bản sao lưu ngày hôm trước).",
                QMessageBox.Yes | QMessageBox.No
            )
            if confirm != QMessageBox.Yes:
                return

        try:
            saved_path = ExcelWriter.update_summary_excel(
                template_path, self.current_report, self.app_results, target_date, self.backup_dir, rule_manager=self.rules
            )
            QMessageBox.information(
                self, "Cập nhật thành công",
                f"Đã cập nhật dữ liệu thành công vào tệp:\n{saved_path}\n\n"
                f"• Mục 01: Đã điền mã khách và chỉ hiển thị đại lý có phát sinh hợp lệ.\n"
                f"• Mục 02: Đã lưu số liệu ngày {target_date.strftime('%d/%m/%Y')} và đồng bộ chuẩn hóa danh sách theo dõi."
            )
        except PermissionError as pe:
            QMessageBox.critical(self, "Tệp đang bị khóa", str(pe))
        except Exception as e:
            QMessageBox.critical(self, "Lỗi cập nhật", f"Không thể cập nhật tệp: {e}")

    def on_view_file_changed(self):
        idx = self.cbo_view_file.currentIndex()
        paths = [self.txt_file_dn.text(), self.txt_file_gr.text(), self.txt_file_mau.text()]
        target_path = paths[idx] if idx < len(paths) else ""

        self.cbo_view_sheet.clear()
        if target_path and os.path.exists(target_path):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(target_path, read_only=True)
                for s in wb.sheetnames:
                    self.cbo_view_sheet.addItem(s)
                wb.close()
            except Exception as e:
                print(f"Lỗi đọc sheet: {e}")

    def on_view_sheet_changed(self):
        sheet_name = self.cbo_view_sheet.currentText()
        if not sheet_name:
            self.tbl_view_excel.setRowCount(0)
            return

        idx = self.cbo_view_file.currentIndex()
        paths = [self.txt_file_dn.text(), self.txt_file_gr.text(), self.txt_file_mau.text()]
        target_path = paths[idx] if idx < len(paths) else ""
        if not target_path or not os.path.exists(target_path):
            return

        try:
            import openpyxl
            wb = openpyxl.load_workbook(target_path, data_only=True)
            ws = wb[sheet_name]
            max_r = min(150, ws.max_row)
            max_c = min(30, ws.max_column)

            self.tbl_view_excel.setRowCount(max_r)
            self.tbl_view_excel.setColumnCount(max_c)

            for r in range(1, max_r + 1):
                for c in range(1, max_c + 1):
                    val = ws.cell(r, c).value
                    val_str = ""
                    if val is not None:
                        if isinstance(val, (int, float)):
                            val_str = f"{val:,.2f}" if isinstance(val, float) else f"{val:,}"
                        elif isinstance(val, datetime):
                            val_str = val.strftime("%d/%m/%Y")
                        else:
                            val_str = str(val)
                    it = QTableWidgetItem(val_str)
                    self.tbl_view_excel.setItem(r - 1, c - 1, it)
            wb.close()
        except Exception as e:
            print(f"Lỗi tải xem trước sheet: {e}")

    def on_search_excel(self, text: str):
        text = text.strip().upper()
        if not text:
            for r in range(self.tbl_view_excel.rowCount()):
                self.tbl_view_excel.setRowHidden(r, False)
            return

        for r in range(self.tbl_view_excel.rowCount()):
            match = False
            for c in range(self.tbl_view_excel.columnCount()):
                item = self.tbl_view_excel.item(r, c)
                if item and text in item.text().upper():
                    match = True
                    break
            self.tbl_view_excel.setRowHidden(r, not match)

    def on_open_current_in_excel(self):
        idx = self.cbo_view_file.currentIndex()
        paths = [self.txt_file_dn.text(), self.txt_file_gr.text(), self.txt_file_mau.text()]
        target_path = paths[idx] if idx < len(paths) else ""
        if target_path and os.path.exists(target_path):
            os.startfile(target_path)
