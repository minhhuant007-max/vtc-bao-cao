import os
import sys
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTextEdit, QProgressBar, QMessageBox, QGroupBox, QLineEdit
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QColor

from ..version import APP_VERSION, DEFAULT_GITHUB_REPO
from ..updater import check_for_updates, download_asset, create_and_launch_updater_script

class CheckUpdateWorker(QThread):
    finished_check = Signal(dict)

    def __init__(self, repo_name: str):
        super().__init__()
        self.repo_name = repo_name

    def run(self):
        res = check_for_updates(self.repo_name)
        self.finished_check.emit(res)

class DownloadWorker(QThread):
    progress = Signal(int, int, int)
    finished_download = Signal(str)
    error = Signal(str)

    def __init__(self, download_url: str):
        super().__init__()
        self.download_url = download_url

    def run(self):
        try:
            import tempfile
            dest_dir = os.path.join(tempfile.gettempdir(), "vtc_update_downloads")
            file_path = download_asset(
                self.download_url, dest_dir,
                progress_callback=lambda p, d, t: self.progress.emit(p, d, t)
            )
            self.finished_download.emit(file_path)
        except Exception as e:
            self.error.emit(str(e))

class UpdateDialog(QDialog):
    def __init__(self, parent=None, repo_name: str = DEFAULT_GITHUB_REPO):
        super().__init__(parent)
        self.repo_name = repo_name
        self.update_info = None

        self.setWindowTitle("Cập nhật ứng dụng VTC Báo Cáo")
        self.resize(560, 480)
        self.init_ui()
        self.on_check_update()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # Header Title
        title_box = QHBoxLayout()
        lbl_icon = QLabel("🔄")
        lbl_icon.setStyleSheet("font-size: 28px;")
        title_box.addWidget(lbl_icon)

        v_title = QVBoxLayout()
        lbl_title = QLabel("CẬP NHẬT ỨNG DỤNG VTC BÁO CÁO")
        lbl_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #1a73e8;")
        lbl_sub = QLabel("Kiểm tra và nâng cấp trực tiếp qua GitHub Releases")
        lbl_sub.setStyleSheet("color: #5f6368; font-size: 12px;")
        v_title.addWidget(lbl_title)
        v_title.addWidget(lbl_sub)
        title_box.addLayout(v_title)
        title_box.addStretch()
        layout.addLayout(title_box)

        # Version Info Box
        info_group = QGroupBox("Thông tin phiên bản")
        info_layout = QVBoxLayout(info_group)

        row_v = QHBoxLayout()
        row_v.addWidget(QLabel("<b>Phiên bản đang dùng:</b>"))
        self.lbl_curr_v = QLabel(f"<span style='color: #202124; font-weight: bold;'>v{APP_VERSION}</span>")
        row_v.addWidget(self.lbl_curr_v)
        row_v.addSpacing(25)
        row_v.addWidget(QLabel("<b>Bản mới nhất trên GitHub:</b>"))
        self.lbl_latest_v = QLabel("<i>Đang kiểm tra...</i>")
        row_v.addWidget(self.lbl_latest_v)
        row_v.addStretch()
        info_layout.addLayout(row_v)

        # Status badge
        self.lbl_status = QLabel("Đang kết nối GitHub để kiểm tra...")
        self.lbl_status.setStyleSheet("padding: 6px; border-radius: 4px; font-weight: bold; background-color: #e8eaed;")
        info_layout.addWidget(self.lbl_status)

        layout.addWidget(info_group)

        # Release Notes
        lbl_notes_hdr = QLabel("<b>Nội dung cập nhật (Release Notes):</b>")
        layout.addWidget(lbl_notes_hdr)

        self.txt_notes = QTextEdit()
        self.txt_notes.setReadOnly(True)
        self.txt_notes.setPlaceholderText("Chưa có thông tin phát hành...")
        layout.addWidget(self.txt_notes)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        # Button Bar
        btn_bar = QHBoxLayout()
        self.btn_check = QPushButton("🔍 Kiểm tra lại")
        self.btn_check.setStyleSheet("padding: 7px 14px; font-weight: bold;")
        self.btn_check.clicked.connect(self.on_check_update)
        btn_bar.addWidget(self.btn_check)

        btn_bar.addStretch()

        self.btn_download_update = QPushButton("⬇ Tải và cập nhật ngay")
        self.btn_download_update.setEnabled(False)
        self.btn_download_update.setStyleSheet("""
            QPushButton { background-color: #0b8043; color: white; padding: 8px 18px; font-weight: bold; border-radius: 4px; }
            QPushButton:hover { background-color: #086333; }
            QPushButton:disabled { background-color: #dadce0; color: #80868b; }
        """)
        self.btn_download_update.clicked.connect(self.on_start_download)
        btn_bar.addWidget(self.btn_download_update)

        btn_close = QPushButton("Đóng")
        btn_close.setStyleSheet("padding: 7px 14px;")
        btn_close.clicked.connect(self.reject)
        btn_bar.addWidget(btn_close)

        layout.addLayout(btn_bar)

    def on_check_update(self):
        self.btn_check.setEnabled(False)
        self.btn_download_update.setEnabled(False)
        self.lbl_latest_v.setText("<i>Đang kiểm tra...</i>")
        self.lbl_status.setText("Đang kết nối GitHub Releases...")
        self.lbl_status.setStyleSheet("padding: 6px; border-radius: 4px; font-weight: bold; background-color: #e8eaed;")

        self.check_worker = CheckUpdateWorker(self.repo_name)
        self.check_worker.finished_check.connect(self.on_check_finished)
        self.check_worker.start()

    def on_check_finished(self, res: dict):
        self.btn_check.setEnabled(True)
        self.update_info = res

        err = res.get("error")
        if err:
            self.lbl_latest_v.setText("<span style='color: #d93025;'>Không xác định</span>")
            self.lbl_status.setText(f"❌ {err}")
            self.lbl_status.setStyleSheet("padding: 6px; border-radius: 4px; font-weight: bold; background-color: #fce8e6; color: #c5221f;")
            self.txt_notes.setPlainText(
                f"Không thể lấy thông tin cập nhật từ repo '{self.repo_name}'.\n\n"
                f"Chi tiết: {err}\n\n"
                "Ghi chú: Nếu repo là riêng tư (Private) hoặc chưa công bố Release đầu tiên, vui lòng liên hệ quản trị viên."
            )
            return

        latest_v = res.get("latest_version", APP_VERSION)
        self.lbl_latest_v.setText(f"<span style='color: #1a73e8; font-weight: bold;'>v{latest_v}</span>")
        self.txt_notes.setPlainText(res.get("release_notes", "") or "Không có mô tả chi tiết.")

        if res.get("has_update"):
            self.lbl_status.setText("🚀 Đã có phiên bản mới sẵn sàng cập nhật!")
            self.lbl_status.setStyleSheet("padding: 6px; border-radius: 4px; font-weight: bold; background-color: #e6f4ea; color: #137333;")
            if res.get("download_url"):
                self.btn_download_update.setEnabled(True)
        else:
            self.lbl_status.setText("✅ Bạn đang sử dụng phiên bản mới nhất!")
            self.lbl_status.setStyleSheet("padding: 6px; border-radius: 4px; font-weight: bold; background-color: #e8f0fe; color: #1a73e8;")

    def on_start_download(self):
        if not self.update_info or not self.update_info.get("download_url"):
            QMessageBox.warning(self, "Thông báo", "Không tìm thấy liên kết tải bản cập nhật!")
            return

        url = self.update_info["download_url"]
        confirm = QMessageBox.question(
            self,
            "Xác nhận cập nhật",
            f"Ứng dụng sẽ tải bản cập nhật mới nhất (v{self.update_info['latest_version']}), "
            "tự động đóng phần mềm và cập nhật phiên bản mới.\n\n"
            "Dữ liệu khách hàng, thiết lập quy luật mã gạch và các bản sao lưu sẽ được bảo toàn nguyên vẹn 100%.\n\n"
            "Bạn có muốn tiếp tục?",
            QMessageBox.Yes | QMessageBox.No
        )
        if confirm != QMessageBox.Yes:
            return

        self.btn_download_update.setEnabled(False)
        self.btn_check.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.lbl_status.setText("Đang tải gói cập nhật từ GitHub...")

        self.dl_worker = DownloadWorker(url)
        self.dl_worker.progress.connect(self.on_download_progress)
        self.dl_worker.finished_download.connect(self.on_download_finished)
        self.dl_worker.error.connect(self.on_download_error)
        self.dl_worker.start()

    def on_download_progress(self, percent: int, downloaded: int, total: int):
        self.progress_bar.setValue(percent)
        mb_down = downloaded / (1024 * 1024)
        mb_tot = total / (1024 * 1024)
        self.lbl_status.setText(f"Đang tải: {percent}% ({mb_down:.1f} MB / {mb_tot:.1f} MB)...")

    def on_download_finished(self, file_path: str):
        self.lbl_status.setText("Tải thành công! Đang kích hoạt tiến trình cập nhật...")
        self.lbl_status.setStyleSheet("padding: 6px; border-radius: 4px; font-weight: bold; background-color: #e6f4ea; color: #137333;")

        app_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        if not getattr(sys, 'frozen', False):
            # Development mode
            app_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

        # Launch updater process
        create_and_launch_updater_script(app_dir, file_path, os.getpid())

        QMessageBox.information(
            self,
            "Khởi động cập nhật",
            "Bản cập nhật đã sẵn sàng! Ứng dụng sẽ đóng lại để hoàn tất quá trình cập nhật và tự động mở lại."
        )
        sys.exit(0)

    def on_download_error(self, err_msg: str):
        self.btn_check.setEnabled(True)
        self.btn_download_update.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.lbl_status.setText(f"Lỗi tải về: {err_msg}")
        self.lbl_status.setStyleSheet("padding: 6px; border-radius: 4px; font-weight: bold; background-color: #fce8e6; color: #c5221f;")
        QMessageBox.critical(self, "Lỗi tải cập nhật", f"Không thể tải bản cập nhật: {err_msg}")
