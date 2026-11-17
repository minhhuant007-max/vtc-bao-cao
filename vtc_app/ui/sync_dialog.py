import os
from typing import Dict, List, Any, Optional, Tuple
import openpyxl
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox,
    QMessageBox, QFrame, QLineEdit
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont

class SyncMasterExcelDialog(QDialog):
    """
    Hộp thoại đối chiếu và xác nhận đồng bộ Mục 02 vào tệp Excel tổng.
    Cho phép người dùng kiểm tra từng dòng, phát hiện xung đột mã,
    xác minh riêng mã CH9999 của Bảo Việt, và quyết định ghi hay không.
    """
    def __init__(self, parent, report_type: str, sync_items: List[Dict[str, Any]], template_path: str):
        super().__init__(parent)
        self.report_type = report_type
        self.sync_items = sync_items
        self.template_path = template_path
        self.confirmed_items = []
        self.is_confirmed = False

        self.setWindowTitle(f"Đối chiếu & Đồng bộ Mục 02 Excel tổng - {report_type}")
        self.resize(1000, 600)
        self.setMinimumSize(850, 480)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Header card
        lbl_title = QLabel("ĐỐI CHIẾU VÀ ĐỒNG BỘ MÃ KHÁCH & NHÓM VÀO MỤC 02")
        lbl_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #1a73e8;")
        layout.addWidget(lbl_title)

        lbl_desc = QLabel(
            f"Tệp Excel đích: <b>{os.path.basename(self.template_path)}</b><br>"
            "Vui lòng kiểm tra các thay đổi bên dưới trước khi ghi vào Excel tổng. "
            "Công thức và số liệu các ngày cũ được giữ nguyên 100%."
        )
        lbl_desc.setWordWrap(True)
        lbl_desc.setStyleSheet("color: #444; font-size: 13px;")
        layout.addWidget(lbl_desc)

        # Thống kê
        needs_verify = sum(1 for it in self.sync_items if it.get("needs_verify"))
        conflicts = sum(1 for it in self.sync_items if it.get("is_conflict"))
        updates = sum(1 for it in self.sync_items if it.get("has_change"))

        stat_text = f"• Tổng số đại lý: <b>{len(self.sync_items)}</b> | • Cập nhật mã/nhóm: <b>{updates}</b>"
        if needs_verify > 0:
            stat_text += f" | <span style='color: #e65100; font-weight: bold;'>⚠️ Cần xác minh (Bảo Việt): {needs_verify}</span>"
        if conflicts > 0:
            stat_text += f" | <span style='color: #c5221f; font-weight: bold;'>⚠️ Xung đột mã: {conflicts}</span>"

        lbl_stat = QLabel(stat_text)
        lbl_stat.setStyleSheet("background-color: #f8f9fa; border: 1px solid #dadce0; border-radius: 4px; padding: 8px;")
        layout.addWidget(lbl_stat)

        # Bảng đối chiếu
        self.tbl_sync = QTableWidget()
        headers = [
            "Đồng bộ", "Dòng", "TH", "Tên đại lý", 
            "Mã hiện tại (Excel)", "Mã sẽ ghi (AppData)", 
            "Nhóm (Excel)", "Nhóm sẽ ghi", "Trạng thái & Ghi chú"
        ]
        self.tbl_sync.setColumnCount(len(headers))
        self.tbl_sync.setHorizontalHeaderLabels(headers)
        self.tbl_sync.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.tbl_sync.horizontalHeader().setSectionResizeMode(8, QHeaderView.ResizeToContents)
        self.tbl_sync.setRowCount(len(self.sync_items))

        bold_font = QFont()
        bold_font.setBold(True)

        for row_idx, it in enumerate(self.sync_items):
            # Cột 0: Checkbox
            chk = QCheckBox()
            # Mặc định checked nếu không xung đột nghiêm trọng
            chk.setChecked(it.get("is_checked", True))
            chk_widget = QFrame()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.addWidget(chk)
            chk_layout.setAlignment(Qt.AlignCenter)
            chk_layout.setContentsMargins(0, 0, 0, 0)
            self.tbl_sync.setCellWidget(row_idx, 0, chk_widget)
            it["_chk_widget"] = chk

            # Cột 1: Dòng
            item_row = QTableWidgetItem(str(it.get("row_num", "")))
            item_row.setTextAlignment(Qt.AlignCenter)
            self.tbl_sync.setItem(row_idx, 1, item_row)

            # Cột 2: TH
            item_th = QTableWidgetItem(str(it.get("th", "")))
            item_th.setTextAlignment(Qt.AlignCenter)
            self.tbl_sync.setItem(row_idx, 2, item_th)

            # Cột 3: Tên đại lý
            item_name = QTableWidgetItem(str(it.get("name", "")))
            item_name.setFont(bold_font)
            self.tbl_sync.setItem(row_idx, 3, item_name)

            # Cột 4: Mã hiện tại trên Excel
            ex_code = str(it.get("excel_code") or "").strip()
            item_ex_code = QTableWidgetItem(ex_code if ex_code else "(Trống)")
            item_ex_code.setTextAlignment(Qt.AlignCenter)
            if not ex_code:
                item_ex_code.setForeground(QColor("#70757a"))
            self.tbl_sync.setItem(row_idx, 4, item_ex_code)

            # Cột 5: Mã sẽ ghi (có thể sửa được)
            app_code = str(it.get("app_code") or "").strip()
            item_app_code = QTableWidgetItem(app_code)
            item_app_code.setTextAlignment(Qt.AlignCenter)
            item_app_code.setFont(bold_font)
            if it.get("needs_verify"):
                item_app_code.setForeground(QColor("#e65100"))
                item_app_code.setBackground(QColor("#fff8e1"))
            elif it.get("is_conflict"):
                item_app_code.setForeground(QColor("#c5221f"))
                item_app_code.setBackground(QColor("#fce8e6"))
            elif it.get("has_code_change"):
                item_app_code.setForeground(QColor("#0b8043"))
                item_app_code.setBackground(QColor("#e6f4ea"))
            self.tbl_sync.setItem(row_idx, 5, item_app_code)

            # Cột 6: Nhóm Excel
            ex_grp = str(it.get("excel_group") or "").strip()
            item_ex_grp = QTableWidgetItem(ex_grp if ex_grp else "(Trống)")
            item_ex_grp.setTextAlignment(Qt.AlignCenter)
            self.tbl_sync.setItem(row_idx, 6, item_ex_grp)

            # Cột 7: Nhóm sẽ ghi
            app_grp = str(it.get("app_group") or "").strip()
            item_app_grp = QTableWidgetItem(app_grp)
            item_app_grp.setTextAlignment(Qt.AlignCenter)
            if it.get("has_group_change"):
                item_app_grp.setForeground(QColor("#0b8043"))
                item_app_grp.setBackground(QColor("#e6f4ea"))
                item_app_grp.setFont(bold_font)
            self.tbl_sync.setItem(row_idx, 7, item_app_grp)

            # Cột 8: Trạng thái & Ghi chú
            note = it.get("note", "")
            item_note = QTableWidgetItem(note)
            if it.get("needs_verify"):
                item_note.setForeground(QColor("#e65100"))
                item_note.setFont(bold_font)
            elif it.get("is_conflict"):
                item_note.setForeground(QColor("#c5221f"))
                item_note.setFont(bold_font)
            elif it.get("has_change"):
                item_note.setForeground(QColor("#0b8043"))
            self.tbl_sync.setItem(row_idx, 8, item_note)

        layout.addWidget(self.tbl_sync)

        # Quick actions bar
        act_bar = QHBoxLayout()
        btn_sel_all = QPushButton("Chọn tất cả")
        btn_sel_all.clicked.connect(lambda: self.toggle_all_checks(True))
        btn_desel_all = QPushButton("Bỏ chọn tất cả")
        btn_desel_all.clicked.connect(lambda: self.toggle_all_checks(False))
        act_bar.addWidget(btn_sel_all)
        act_bar.addWidget(btn_desel_all)
        act_bar.addStretch()

        layout.addLayout(act_bar)

        # Bottom buttons
        btn_box = QHBoxLayout()
        btn_box.addStretch()

        self.btn_cancel = QPushButton("Hủy bỏ")
        self.btn_cancel.setStyleSheet("padding: 8px 16px;")
        self.btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(self.btn_cancel)

        self.btn_confirm = QPushButton("✓ Xác nhận & Cập nhật Excel")
        self.btn_confirm.setStyleSheet(
            "background-color: #1a73e8; color: white; font-weight: bold; padding: 8px 24px; border-radius: 4px;"
        )
        self.btn_confirm.clicked.connect(self.on_confirm_clicked)
        btn_box.addWidget(self.btn_confirm)

        layout.addLayout(btn_box)

    def toggle_all_checks(self, checked: bool):
        for it in self.sync_items:
            chk = it.get("_chk_widget")
            if chk:
                chk.setChecked(checked)

    def on_confirm_clicked(self):
        # Thu thập các dòng được chọn và giá trị mã/nhóm mới nhất từ bảng
        self.confirmed_items = []
        for row_idx, it in enumerate(self.sync_items):
            chk = it.get("_chk_widget")
            if chk and chk.isChecked():
                # Lấy mã người dùng đã nhập/chỉnh sửa trong bảng
                item_code = self.tbl_sync.item(row_idx, 5)
                final_code = item_code.text().strip() if item_code else it.get("app_code", "")

                # Lấy nhóm
                item_grp = self.tbl_sync.item(row_idx, 7)
                final_grp = item_grp.text().strip() if item_grp else it.get("app_group", "")

                item_copy = dict(it)
                item_copy["final_code"] = final_code
                item_copy["final_group"] = final_grp
                self.confirmed_items.append(item_copy)

        # Kiểm tra riêng mã CH9999 của Bảo Việt
        for it in self.confirmed_items:
            if "BẢO VIỆT" in it.get("name", "").upper() and it.get("final_code") == "CH9999":
                res = QMessageBox.question(
                    self,
                    "Xác minh mã Bảo Việt",
                    "Đại lý BẢO VIỆT - ĐẮK NÔNG hiện đang dùng mã tạm 'CH9999'.\n\n"
                    "Bạn có xác nhận ghi mã 'CH9999' vào Excel tổng không?\n"
                    "• Bấm 'Yes' để ghi CH9999.\n"
                    "• Bấm 'No' để để trống mã khách (không ghi CH9999).",
                    QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
                )
                if res == QMessageBox.Cancel:
                    return
                elif res == QMessageBox.No:
                    it["final_code"] = ""

        self.is_confirmed = True
        self.accept()
