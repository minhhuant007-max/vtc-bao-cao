from dataclasses import dataclass, field
from datetime import datetime, date
from typing import Optional, List, Dict, Any

@dataclass
class OrderRow:
    source_file: str
    sheet_name: str
    row_num: int
    so_ct: str
    ngay_ct: Optional[date]
    ma_khach: str
    ten_khach: str
    trang_thai: str
    ma_hang: str
    ten_mat_hang: str
    ma_lo: Optional[str] = None
    dvt: str = "thùng"
    sl_dat_hang: float = 0.0
    don_gia: float = 0.0
    thanh_tien: float = 0.0
    dien_giai: str = ""
    ghi_chu: str = ""

@dataclass
class TrackedCustomer:
    ma_khach: str
    stt: int
    ten_khach: str
    nhom: str
    thuong_hieu: str  # TCR, TML, TMR, TRM, TPS
    row_idx_muc02: int = 0
    col_idx_ma_khach: int = 1
    is_active: bool = True

@dataclass
class SpecDefinition:
    name: str
    col_idx: int
    factor_m2: float
    box_pieces: int = 0
    description: str = ""

@dataclass
class DiscrepancyRecord:
    report_type: str
    ngay_bc: str
    ma_khach: str
    ten_khach: str
    thuong_hieu: str
    quy_cach: str
    so_app_thung: float
    so_tay_thung: float
    chenh_lech_thung: float
    so_app_m2: float
    so_tay_m2: float
    chenh_lech_m2: float
    nguon_chung_tu: str
    chung_tu_lien_quan: str
    ly_do: str
    loai_chenh_lech: str  # KHỚP, TAY_TRỪ_CHỨNG_TỪ, NGUỒN_THIẾU_CHỨNG_TỪ, THIẾU_MÃ_MỤC02, CHƯA_XÁC_NHẬN
