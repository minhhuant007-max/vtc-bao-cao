# VTC Báo Cáo - Hệ Thống Báo Cáo Tiêu Thụ Gạch Men VTC

Hệ thống xử lý, tổng hợp và lập báo cáo bán hàng hằng ngày tự động cho các thương hiệu:
- **TCR + TML** (Tocera & Tomilan)
- **TMR** (Tomira)
- **TRM + TPS** (Toroma & Topisa)

## 1. Yêu cầu môi trường
- Python 3.11+ (Khuyến nghị Python 3.13 x64)
- Các thư viện phụ thuộc:
  ```bash
  pip install PySide6 openpyxl
  pip install pyinstaller
  ```

## 2. Cấu trúc thư mục
- `vtc_app/`: Mã nguồn ứng dụng (Engine, Reader, Writer, Rules, Models, UI PySide6).
- `config/`: Cấu hình danh mục đại lý (`dealers.json`), quy luật mã gạch (`tile_rules.json`), hệ số và ánh xạ (`rules.json`).
- `main.py`: Điểm khởi chạy ứng dụng GUI hoặc CLI batch export.
- `build_x64.py`: Script đóng gói PyInstaller x64 (onedir, windowed).

## 3. Hướng dẫn đóng gói ứng dụng (Build EXE)
Chạy script:
```bash
python build_x64.py
```
Tệp thực thi sau khi hoàn tất sẽ nằm tại:
`dist/VTC_Bao_Cao/VTC_Bao_Cao.exe`
Thư mục `config/` mặc định được đóng gói tự động đi kèm.
