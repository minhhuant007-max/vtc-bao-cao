import os
from typing import List, Dict, Any, Optional
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QComboBox, QCheckBox, QMessageBox, QGroupBox, QInputDialog,
    QFormLayout, QDialogButtonBox
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QFont

from ..rules import RuleManager
from ..reader import ExcelReader

class AddEditDealerDialog(QDialog):
    def __init__(self, report_type: str, rule_manager: RuleManager, source_files: List[str], dealer_data: Optional[Dict[str, Any]] = None, parent=None):
        super().__init__(parent)
        self.report_type = report_type
        self.rule_manager = rule_manager
        self.source_files = source_files
        self.dealer_data = dealer_data or {}
        self.is_edit = bool(dealer_data)

        self.setWindowTitle("Sửa đại lý theo dõi" if self.is_edit else "Thêm đại lý theo dõi mới")
        self.resize(500, 320)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()

        # 1. Report Type (read-only display)
        lbl_rep = QLabel(self.report_type)
        lbl_rep.setStyleSheet("font-weight: bold; color: #1a73e8;")
        form.addRow("Nhóm báo cáo:", lbl_rep)

        # 2. Customer Code with Auto-suggest
        self.txt_code = QLineEdit()
        self.txt_code.setPlaceholderText("Ví dụ: KH0043, CH0770...")
        if self.dealer_data.get("ma_khach"):
            self.txt_code.setText(self.dealer_data.get("ma_khach", ""))
        self.txt_code.textChanged.connect(self.on_code_changed)
        form.addRow("Mã khách hàng (*):", self.txt_code)

        # 3. Dealer Name
        self.txt_name = QLineEdit()
        self.txt_name.setPlaceholderText("Tên đại lý / Cửa hàng")
        if self.dealer_data.get("ten_khach"):
            self.txt_name.setText(self.dealer_data.get("ten_khach", ""))
        form.addRow("Tên đại lý (*):", self.txt_name)

        # 4. Group (VIP, GPP, D1)
        self.cbo_nhom = QComboBox()
        self.cbo_nhom.addItems(["VIP", "GPP", "D1"])
        curr_nhom = self.dealer_data.get("nhom", "VIP").upper()
        idx_nhom = self.cbo_nhom.findText(curr_nhom)
        if idx_nhom >= 0:
            self.cbo_nhom.setCurrentIndex(idx_nhom)
        form.addRow("Nhóm đại lý:", self.cbo_nhom)

        # 5. Brand (Thương hiệu)
        self.cbo_th = QComboBox()
        if self.report_type == "TCR_TML":
            self.cbo_th.addItems(["TCR", "TML"])
        elif self.report_type == "TMR":
            self.cbo_th.addItems(["TMR"])
        elif self.report_type == "TRM_TPS":
            self.cbo_th.addItems(["TRM", "TPS"])
        else:
            self.cbo_th.addItems(["TCR", "TML", "TMR", "TRM", "TPS"])

        curr_th = self.dealer_data.get("thuong_hieu", "").upper()
        idx_th = self.cbo_th.findText(curr_th)
        if idx_th >= 0:
            self.cbo_th.setCurrentIndex(idx_th)
        form.addRow("Thương hiệu (*):", self.cbo_th)

        # 6. Active tracking status
        self.chk_active = QCheckBox("Đang theo dõi trong báo cáo Mục 01 và tính toán hằng ngày")
        self.chk_active.setChecked(self.dealer_data.get("is_active", True))
        form.addRow("Trạng thái:", self.chk_active)

        layout.addLayout(form)

        # Auto-suggest info label
        self.lbl_suggest = QLabel("")
        self.lbl_suggest.setStyleSheet("color: #0b8043; font-style: italic;")
        layout.addWidget(self.lbl_suggest)

        # Buttons
        btn_box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.on_save)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def on_code_changed(self, text: str):
        if self.is_edit:
            return
        code = text.strip().upper()
        if not code or len(code) < 3:
            self.lbl_suggest.setText("")
            return

        # Scan source files for auto-suggest
        found_name = ""
        found_nhom = ""
        for p in self.source_files:
            if p and os.path.exists(p):
                try:
                    orders = ExcelReader.read_orders(p)
                    for o in orders:
                        if o.ma_khach.strip().upper() == code:
                            found_name = o.ten_khach
                            t_upper = o.ten_khach.upper()
                            if "(VIP)" in t_upper:
                                found_nhom = "VIP"
                            elif "(GPP)" in t_upper:
                                found_nhom = "GPP"
                            elif "(D1)" in t_upper:
                                found_nhom = "D1"
                            break
                    if found_name:
                        break
                except Exception:
                    pass

        if found_name:
            if not self.txt_name.text().strip():
                self.txt_name.setText(found_name)
            if found_nhom:
                idx = self.cbo_nhom.findText(found_nhom)
                if idx >= 0:
                    self.cbo_nhom.setCurrentIndex(idx)
            self.lbl_suggest.setText(f"💡 Đã tìm thấy từ nguồn: {found_name} ({found_nhom or 'Nhóm'})")
        else:
            self.lbl_suggest.setText("")

    def on_save(self):
        code = self.txt_code.text().strip().upper()
        name = self.txt_name.text().strip()
        nhom = self.cbo_nhom.currentText()
        th = self.cbo_th.currentText()
        is_act = self.chk_active.isChecked()

        if not name:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng nhập Tên đại lý!")
            return

        self.result_data = {
            "ma_khach": code,
            "ten_khach": name,
            "nhom": nhom,
            "thuong_hieu": th,
            "is_active": is_act,
            "stt": self.dealer_data.get("stt", 0)
        }
        self.accept()

class DealersDialog(QDialog):
    dealers_changed = Signal()

    def __init__(self, rule_manager: RuleManager, initial_report_type: str = "TRM_TPS", source_files: Optional[List[str]] = None, parent=None):
        super().__init__(parent)
        self.rule_manager = rule_manager
        self.current_report = initial_report_type
        self.source_files = source_files or []

        self.setWindowTitle("Quản lý đại lý theo dõi (Mục 02)")
        self.resize(880, 600)
        self.setup_ui()
        self.load_dealers_table()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # Top Bar: Report Group & Brand Filter
        top_bar = QHBoxLayout()
        top_bar.addWidget(QLabel("<b>Nhóm báo cáo:</b>"))

        self.cbo_rep = QComboBox()
        self.cbo_rep.addItem("TRM + TPS (Toroma – Topisa)", "TRM_TPS")
        self.cbo_rep.addItem("TCR + TML (Tocera – Tomilan)", "TCR_TML")
        self.cbo_rep.addItem("TMR (Tomira)", "TMR")
        idx = self.cbo_rep.findData(self.current_report)
        if idx >= 0:
            self.cbo_rep.setCurrentIndex(idx)
        self.cbo_rep.currentIndexChanged.connect(self.on_report_changed)
        top_bar.addWidget(self.cbo_rep)

        top_bar.addSpacing(15)
        top_bar.addWidget(QLabel("<b>Lọc thương hiệu:</b>"))
        self.cbo_brand_filter = QComboBox()
        self.cbo_brand_filter.currentIndexChanged.connect(self.load_dealers_table)
        top_bar.addWidget(self.cbo_brand_filter)

        top_bar.addStretch()

        # Search box
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("🔍 Tìm theo mã hoặc tên...")
        self.txt_search.setFixedWidth(200)
        self.txt_search.textChanged.connect(self.filter_dealers_display)
        top_bar.addWidget(self.txt_search)

        layout.addLayout(top_bar)
        self.update_brand_filter_items()

        # Dealers Table
        self.table = QTableWidget()
        headers = ["STT", "Mã khách hàng", "Tên đại lý (Cửa hàng)", "Nhóm", "Thương hiệu", "Trạng thái", "Thao tác"]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.table)

        # Action Buttons
        btn_bar = QHBoxLayout()
        btn_add = QPushButton("➕ Thêm đại lý mới")
        btn_add.setStyleSheet("background-color: #1a73e8; color: white; padding: 7px 15px; font-weight: bold; border-radius: 4px;")
        btn_add.clicked.connect(self.on_add_dealer)
        btn_bar.addWidget(btn_add)

        btn_edit = QPushButton("✏ Sửa đại lý đang chọn")
        btn_edit.setStyleSheet("padding: 7px 15px; font-weight: bold;")
        btn_edit.clicked.connect(self.on_edit_dealer)
        btn_bar.addWidget(btn_edit)

        btn_toggle = QPushButton("🔄 Đổi trạng thái Theo dõi / Ngừng")
        btn_toggle.setStyleSheet("padding: 7px 15px; font-weight: bold;")
        btn_toggle.clicked.connect(self.on_toggle_active)
        btn_bar.addWidget(btn_toggle)

        btn_bar.addStretch()

        btn_save = QPushButton("💾 Lưu danh sách")
        btn_save.setStyleSheet("background-color: #0b8043; color: white; padding: 7px 20px; font-weight: bold; border-radius: 4px;")
        btn_save.clicked.connect(self.on_save_all)
        btn_bar.addWidget(btn_save)

        btn_close = QPushButton("Đóng")
        btn_close.clicked.connect(self.reject)
        btn_bar.addWidget(btn_close)

        layout.addLayout(btn_bar)

    def on_report_changed(self):
        self.current_report = self.cbo_rep.currentData()
        self.update_brand_filter_items()
        self.load_dealers_table()

    def update_brand_filter_items(self):
        self.cbo_brand_filter.blockSignals(True)
        self.cbo_brand_filter.clear()
        self.cbo_brand_filter.addItem("Tất cả", "")
        if self.current_report == "TCR_TML":
            self.cbo_brand_filter.addItem("TCR", "TCR")
            self.cbo_brand_filter.addItem("TML", "TML")
        elif self.current_report == "TMR":
            self.cbo_brand_filter.addItem("TMR", "TMR")
        elif self.current_report == "TRM_TPS":
            self.cbo_brand_filter.addItem("TRM", "TRM")
            self.cbo_brand_filter.addItem("TPS", "TPS")
        self.cbo_brand_filter.blockSignals(False)

    def load_dealers_table(self):
        brand_filter = self.cbo_brand_filter.currentData()
        dealers = self.rule_manager.get_dealers(self.current_report, brand=brand_filter if brand_filter else None)

        self.table.setRowCount(len(dealers))
        for r, d in enumerate(dealers):
            stt = d.get("stt", r + 1)
            code = d.get("ma_khach", "")
            name = d.get("ten_khach", "")
            nhom = d.get("nhom", "")
            th = d.get("thuong_hieu", "")
            is_active = d.get("is_active", True)

            item_stt = QTableWidgetItem(str(stt))
            item_stt.setTextAlignment(Qt.AlignCenter)
            item_code = QTableWidgetItem(code)
            item_name = QTableWidgetItem(name)
            item_nhom = QTableWidgetItem(nhom)
            item_nhom.setTextAlignment(Qt.AlignCenter)
            item_th = QTableWidgetItem(th)
            item_th.setTextAlignment(Qt.AlignCenter)

            status_str = "Đang theo dõi" if is_active else "Ngừng theo dõi"
            item_status = QTableWidgetItem(status_str)
            item_status.setTextAlignment(Qt.AlignCenter)
            if is_active:
                item_status.setForeground(QColor("#0b8043"))
            else:
                item_status.setForeground(QColor("#d93025"))

            item_action = QTableWidgetItem("Bấm đúp để sửa")
            item_action.setTextAlignment(Qt.AlignCenter)
            item_action.setForeground(QColor("#888888"))

            if not is_active:
                for it in [item_stt, item_code, item_name, item_nhom, item_th]:
                    it.setForeground(QColor("#9aa0a6"))

            self.table.setItem(r, 0, item_stt)
            self.table.setItem(r, 1, item_code)
            self.table.setItem(r, 2, item_name)
            self.table.setItem(r, 3, item_nhom)
            self.table.setItem(r, 4, item_th)
            self.table.setItem(r, 5, item_status)
            self.table.setItem(r, 6, item_action)

            # Store dealer data in hidden role of item_code
            item_code.setData(Qt.UserRole, d)

        self.filter_dealers_display()

    def filter_dealers_display(self):
        query = self.txt_search.text().strip().upper()
        for r in range(self.table.rowCount()):
            code = (self.table.item(r, 1).text() if self.table.item(r, 1) else "").upper()
            name = (self.table.item(r, 2).text() if self.table.item(r, 2) else "").upper()
            match = not query or (query in code or query in name)
            self.table.setRowHidden(r, not match)

    def on_add_dealer(self):
        dlg = AddEditDealerDialog(self.current_report, self.rule_manager, self.source_files, parent=self)
        if dlg.exec():
            new_data = dlg.result_data
            self.rule_manager.add_or_update_dealer(self.current_report, new_data)
            self.load_dealers_table()
            self.dealers_changed.emit()

    def on_edit_dealer(self):
        selected_row = self.table.currentRow()
        if selected_row < 0:
            QMessageBox.information(self, "Thông báo", "Vui lòng chọn đại lý cần sửa trong bảng!")
            return

        item_code = self.table.item(selected_row, 1)
        dealer_data = item_code.data(Qt.UserRole) if item_code else None
        if not dealer_data:
            return

        dlg = AddEditDealerDialog(self.current_report, self.rule_manager, self.source_files, dealer_data=dealer_data, parent=self)
        if dlg.exec():
            updated_data = dlg.result_data
            self.rule_manager.add_or_update_dealer(self.current_report, updated_data)
            self.load_dealers_table()
            self.dealers_changed.emit()

    def on_toggle_active(self):
        selected_row = self.table.currentRow()
        if selected_row < 0:
            QMessageBox.information(self, "Thông báo", "Vui lòng chọn đại lý cần đổi trạng thái!")
            return

        item_code = self.table.item(selected_row, 1)
        dealer_data = item_code.data(Qt.UserRole) if item_code else None
        if not dealer_data:
            return

        curr_active = dealer_data.get("is_active", True)
        new_active = not curr_active
        code = dealer_data.get("ma_khach", "")
        th = dealer_data.get("thuong_hieu", "")

        self.rule_manager.set_dealer_active(self.current_report, code, th, new_active)
        self.load_dealers_table()
        self.dealers_changed.emit()

    def on_save_all(self):
        self.rule_manager.save_dealers()
        self.dealers_changed.emit()
        QMessageBox.information(self, "Thành công", "Đã lưu danh sách đại lý thành công!")
        self.accept()
