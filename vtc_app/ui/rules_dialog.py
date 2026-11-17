import os
import re
from typing import List, Dict, Any, Optional
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QComboBox, QMessageBox, QGroupBox, QWidget, QSplitter
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QFont

from ..rules import RuleManager

class RulesDialog(QDialog):
    rules_changed = Signal()

    def __init__(self, rule_manager: RuleManager, initial_report_type: str = "TRM_TPS", parent=None):
        super().__init__(parent)
        self.rule_manager = rule_manager
        self.current_report = initial_report_type

        self.setWindowTitle("Quy luật mã gạch & Hệ số tính m²")
        self.resize(1050, 680)
        self.setup_ui()
        self.load_specs_table()

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
        top_bar.addWidget(QLabel("<b>Thương hiệu:</b>"))
        self.cbo_brand = QComboBox()
        self.cbo_brand.currentIndexChanged.connect(self.on_brand_changed)
        top_bar.addWidget(self.cbo_brand)

        top_bar.addStretch()

        lbl_guide = QLabel("<i>Đuôi 'XX' là ký tự đại diện (ví dụ T66XX). Mã không đuôi XX là khớp chính xác.</i>")
        lbl_guide.setStyleSheet("color: #555555;")
        top_bar.addWidget(lbl_guide)

        layout.addLayout(top_bar)
        self.update_brand_items()

        # Alert banner for conflicts
        self.lbl_conflict_banner = QLabel("")
        self.lbl_conflict_banner.setStyleSheet("background-color: #fce8e6; color: #c5221f; font-weight: bold; padding: 6px; border-radius: 4px;")
        self.lbl_conflict_banner.setVisible(False)
        layout.addWidget(self.lbl_conflict_banner)

        # Main Table of Spec Columns
        self.table = QTableWidget()
        headers = ["STT", "Cột quy cách (Excel)", "Nguồn áp dụng", "Hệ số m²", "Mẫu mã nhận diện / Tiền tố (Cách nhau bằng dấu phẩy)"]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        layout.addWidget(self.table)

        # Bottom Testing Box ("Thử mã hàng")
        grp_test = QGroupBox("🔍 Thử mã hàng & Kiểm tra tính hợp lệ quy tắc")
        grp_test_layout = QVBoxLayout(grp_test)

        test_inputs = QHBoxLayout()
        test_inputs.addWidget(QLabel("Mã hàng:"))
        self.txt_test_code = QLineEdit()
        self.txt_test_code.setPlaceholderText("Ví dụ: T66108L1, T6166NL1, R020-AA, TTP447NL1...")
        self.txt_test_code.returnPressed.connect(self.on_run_test)
        test_inputs.addWidget(self.txt_test_code)

        test_inputs.addWidget(QLabel("Tên mặt hàng (tùy chọn):"))
        self.txt_test_name = QLineEdit()
        self.txt_test_name.setPlaceholderText("Ví dụ: PORCELAIN, GẠCH MEN, ROTIC...")
        self.txt_test_name.returnPressed.connect(self.on_run_test)
        test_inputs.addWidget(self.txt_test_name)

        test_inputs.addWidget(QLabel("Nguồn:"))
        self.cbo_test_source = QComboBox()
        self.cbo_test_source.addItems(["Đồng Nai", "GRANITE"])
        test_inputs.addWidget(self.cbo_test_source)

        btn_test = QPushButton("Kiểm tra")
        btn_test.setStyleSheet("background-color: #1a73e8; color: white; padding: 5px 15px; font-weight: bold; border-radius: 4px;")
        btn_test.clicked.connect(self.on_run_test)
        test_inputs.addWidget(btn_test)

        grp_test_layout.addLayout(test_inputs)

        # Test results display
        self.lbl_test_result = QLabel("Nhập mã hàng và bấm 'Kiểm tra' để xem kết quả phân loại cột đích và kiểm tra nguồn.")
        self.lbl_test_result.setStyleSheet("font-size: 13px; padding: 6px; background-color: #f8f9fa; border-radius: 4px;")
        grp_test_layout.addWidget(self.lbl_test_result)

        layout.addWidget(grp_test)

        # Action Buttons
        btn_bar = QHBoxLayout()
        btn_bar.addStretch()

        btn_save = QPushButton("💾 Lưu quy tắc")
        btn_save.setStyleSheet("background-color: #0b8043; color: white; padding: 8px 24px; font-weight: bold; font-size: 13px; border-radius: 4px;")
        btn_save.clicked.connect(self.on_save_all)
        btn_bar.addWidget(btn_save)

        btn_close = QPushButton("Đóng")
        btn_close.clicked.connect(self.reject)
        btn_bar.addWidget(btn_close)

        layout.addLayout(btn_bar)

    def on_report_changed(self):
        self.current_report = self.cbo_rep.currentData()
        self.update_brand_items()
        self.load_specs_table()

    def update_brand_items(self):
        self.cbo_brand.blockSignals(True)
        self.cbo_brand.clear()
        if self.current_report == "TCR_TML":
            self.cbo_brand.addItem("TCR", "TCR")
            self.cbo_brand.addItem("TML", "TML")
        elif self.current_report == "TMR":
            self.cbo_brand.addItem("TMR", "TMR")
        elif self.current_report == "TRM_TPS":
            self.cbo_brand.addItem("TRM", "TRM")
            self.cbo_brand.addItem("TPS", "TPS")
        self.cbo_brand.blockSignals(False)

    def on_brand_changed(self):
        self.load_specs_table()

    def load_specs_table(self):
        specs = self.rule_manager.get_spec_rules(self.current_report)
        self.table.setRowCount(len(specs))

        for r, sp in enumerate(specs):
            name = sp.get("name", "")
            source = sp.get("source", "BOTH")
            factor = sp.get("factor", 1.0)
            patterns = sp.get("patterns", [])
            pattern_str = ", ".join(patterns)

            # 0. STT
            item_stt = QTableWidgetItem(str(r + 1))
            item_stt.setTextAlignment(Qt.AlignCenter)
            item_stt.setFlags(item_stt.flags() & ~Qt.ItemIsEditable)

            # 1. Spec Name
            item_name = QTableWidgetItem(name)
            item_name.setFont(QFont("Segoe UI", 9, QFont.Bold))
            item_name.setFlags(item_name.flags() & ~Qt.ItemIsEditable)

            # 2. Source (Cho phép chỉnh Đồng Nai hoặc GRANITE cho 40x80)
            if "40X80" in name.upper():
                cbo_src = QComboBox()
                cbo_src.addItems(["GRANITE", "Đồng Nai"])
                if source == "DONG_NAI":
                    cbo_src.setCurrentText("Đồng Nai")
                else:
                    cbo_src.setCurrentText("GRANITE")
                cbo_src.setStyleSheet("font-weight: bold; color: #b05000; padding: 2px;")
                self.table.setCellWidget(r, 2, cbo_src)
            else:
                if source == "DONG_NAI":
                    src_label = "Đồng Nai"
                    badge_color = QColor("#1a73e8")
                elif source == "GRANITE":
                    src_label = "GRANIT"
                    badge_color = QColor("#e37400")
                else:
                    src_label = "Cả hai nguồn"
                    badge_color = QColor("#5f6368")

                item_source = QTableWidgetItem(src_label)
                item_source.setTextAlignment(Qt.AlignCenter)
                item_source.setForeground(badge_color)
                item_source.setFont(QFont("Segoe UI", 9, QFont.Bold))
                item_source.setFlags(item_source.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(r, 2, item_source)

            # 3. Factor m² (editable)
            item_factor = QTableWidgetItem(str(factor))
            item_factor.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            # 4. Patterns (editable)
            item_patterns = QTableWidgetItem(pattern_str)

            self.table.setItem(r, 0, item_stt)
            self.table.setItem(r, 1, item_name)
            self.table.setItem(r, 3, item_factor)
            self.table.setItem(r, 4, item_patterns)

        self.check_conflicts()

    def check_conflicts(self):
        # Scan patterns across rows to detect identical priority conflicts
        pattern_map: Dict[str, List[Tuple[int, str]]] = {}
        conflicts = []

        for r in range(self.table.rowCount()):
            sp_name = self.table.item(r, 1).text().strip()
            patterns_text = self.table.item(r, 4).text() if self.table.item(r, 4) else ""
            raw_patterns = [p.strip().upper() for p in re.split(r"[,;\n]", patterns_text) if p.strip()]

            for p in raw_patterns:
                if p not in pattern_map:
                    pattern_map[p] = []
                pattern_map[p].append((r, sp_name))

        for p, occurrences in pattern_map.items():
            specs_matched = {occ[1] for occ in occurrences}
            if len(specs_matched) > 1:
                conflicts.append(f"Mẫu '{p}' đang trỏ vào nhiều cột: {', '.join(specs_matched)}")

        if conflicts:
            self.lbl_conflict_banner.setText("⚠️ CẢNH BÁO XUNG ĐỘT QUY TẮC:\n• " + "\n• ".join(conflicts))
            self.lbl_conflict_banner.setVisible(True)
        else:
            self.lbl_conflict_banner.setVisible(False)

    def on_run_test(self):
        code = self.txt_test_code.text().strip().upper()
        if not code:
            QMessageBox.warning(self, "Thông báo", "Vui lòng nhập Mã hàng cần kiểm tra!")
            return

        name = self.txt_test_name.text().strip()
        source = self.cbo_test_source.currentText()
        brand = self.cbo_brand.currentText()

        # Temporarily apply current table patterns to RuleManager for accurate testing
        self.apply_table_to_manager()

        res = self.rule_manager.test_item_code(
            self.current_report, brand, code, name, source
        )

        spec = res["target_spec"]
        pat = res["matched_pattern"]
        factor = res["factor"]
        prio = res["priority"]
        is_src_valid = res["is_source_valid"]
        src_reason = res["source_reason"]
        is_conflict = res["is_conflict"]
        conflict_specs = res["conflict_specs"]

        src_status_html = (
            f"<b style='color: #0b8043;'>✅ HỢP LỆ THEO NGUỒN ({source})</b>"
            if is_src_valid
            else f"<b style='color: #d93025;'>❌ TỪ CHỐI NGUỒN: {src_reason} ({source})</b>"
        )

        conflict_html = ""
        if is_conflict:
            conflict_html = f"<br><b style='color: #c5221f;'>⚠️ Xung đột quy tắc: Cùng độ ưu tiên trỏ tới nhiều cột ({', '.join(conflict_specs)})</b>"

        prio_desc = "Khớp chính xác (9999)" if prio == 9999 else (f"Khớp tiền tố dài {prio} ký tự" if prio > 0 else "Quy tắc mặc định")

        result_html = (
            f"<b>Mã hàng:</b> <span style='color: #1a73e8; font-weight: bold;'>{code}</span> | "
            f"<b>Cột đích:</b> <span style='color: #188038; font-weight: bold;'>{spec}</span> | "
            f"<b>Mẫu khớp:</b> {pat} (<i>{prio_desc}</i>) | "
            f"<b>Hệ số m²:</b> {factor} m²/thùng<br>"
            f"<b>Kiểm tra nguồn:</b> {src_status_html}{conflict_html}"
        )
        self.lbl_test_result.setText(result_html)

    def apply_table_to_manager(self):
        specs = self.rule_manager.spec_rules.setdefault(self.current_report, [])
        for r in range(self.table.rowCount()):
            sp_name = self.table.item(r, 1).text().strip()
            fac_str = self.table.item(r, 3).text().strip()
            pat_str = self.table.item(r, 4).text().strip() if self.table.item(r, 4) else ""

            try:
                fac = float(fac_str)
            except ValueError:
                fac = 1.0

            patterns = [p.strip().upper() for p in re.split(r"[,;\n]", pat_str) if p.strip()]

            widget_src = self.table.cellWidget(r, 2)
            src_val = None
            if isinstance(widget_src, QComboBox):
                src_val = "DONG_NAI" if "ĐỒNG NAI" in widget_src.currentText().upper() else "GRANITE"

            # Find in specs
            found = False
            for sp in specs:
                if sp.get("name", "").strip().lower() == sp_name.lower():
                    sp["factor"] = fac
                    sp["patterns"] = patterns
                    if src_val:
                        sp["source"] = src_val
                    found = True
                    break
            if not found:
                specs.append({
                    "name": sp_name,
                    "factor": fac,
                    "source": src_val or "BOTH",
                    "patterns": patterns
                })

    def on_save_all(self):
        self.apply_table_to_manager()
        self.rule_manager.save_spec_rules()
        self.check_conflicts()
        self.rules_changed.emit()
        QMessageBox.information(self, "Thành công", "Đã lưu quy luật mã gạch và hệ số thành công!")
        self.accept()
