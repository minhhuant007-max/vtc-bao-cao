import os
import json
import re
from typing import Dict, Any, List, Optional, Tuple

DEFAULT_FACTORS = {
    "TCR_TML": {
        "20X40": 1.0,
        "25X40 (15 V/T)": 1.5,
        "25X50": 1.0,
        "30X45  (12 V/T)": 1.62,
        "30X60 ROTIC": 1.8,
        "30X60 TTP": 1.44,
        "30X60 TH": 1.44,
        "40X80": 1.92,
        "30X30 (11 V/T)": 0.99,
        "30X30 (16 V/T)": 1.44,
        "40X40": 0.96,
        "40X40 (12V/T)": 1.92,
        "40X60": 1.92,
        "50X50 (6 V/T)": 1.5,
        "60x60 (5 V/T)": 1.8,
        "60x60 ĐÁ (5 V/T)": 1.8,
        "80x80 ĐÁ": 1.92
    },
    "TMR": {
        "20x40": 1.0,
        "25x40": 1.5,
        "25x50": 1.0,
        "30x45": 1.62,
        "30x60 ROTIC": 1.8,
        "30x60 TTP(8v)": 1.44,
        "30x60 TTP(10v)": 1.8,
        "30x60 TH(8v)": 1.44,
        "30x60 TH(10v)": 1.8,
        "40x80": 1.92,
        "30x30": 1.44,
        "40x40": 1.44,
        "40x60": 1.92,
        "50x50": 1.5,
        "60x60": 1.8,
        "60x60 (P)": 1.8,
        "80x80 (P)": 1.92
    },
    "TRM_TPS": {
        "20x40": 0.96,
        "25x40 (15v/t)": 1.5,
        "25x50": 1.0,
        "30x60 ROTIC": 1.8,
        "30x60 TTP(10v)": 1.8,
        "30x60(10v) TRM/TPS": 1.8,
        "40x80": 1.92,
        "30x30 (16v/t)": 1.44,
        "40x40": 0.96,
        "40x40 (12v/t)": 1.92,
        "40x60": 1.8,
        "50x50 (06v/t)": 1.5,
        "60x60 (05v/t)": 1.8,
        "60x60.P (05v/t)": 1.8,
        "80x80.P (03v/t)": 1.92
    }
}

DEFAULT_SOURCE_EXCEPTIONS = {}
DEFAULT_VOUCHER_ADJUSTMENTS = []
DEFAULT_PENDING_DEALERS = []

DEFAULT_DEALERS = {
    "TCR_TML": [
        {"ma_khach": "KH0001", "stt": 1, "ten_khach": "ĐẠI LÝ MẪU 01", "nhom": "VIP", "thuong_hieu": "TCR", "is_active": True}
    ],
    "TMR": [
        {"ma_khach": "KH0002", "stt": 1, "ten_khach": "ĐẠI LÝ MẪU 02", "nhom": "VIP", "thuong_hieu": "TMR", "is_active": True}
    ],
    "TRM_TPS": [
        {"ma_khach": "KH0003", "stt": 1, "ten_khach": "ĐẠI LÝ MẪU 03", "nhom": "VIP", "thuong_hieu": "TRM", "is_active": True}
    ]
}

DEFAULT_SPEC_RULES = {
    "TCR_TML": [
        {"name": "20X40", "factor": 1.0, "source": "BOTH", "patterns": ["20X40*", "TTP24XX"]},
        {"name": "25X40 (15 V/T)", "factor": 1.5, "source": "BOTH", "patterns": ["25X40*", "TTP220XX", "TTPV220XX", "TTP250XX", "TTP251XX"]},
        {"name": "25X50", "factor": 1.0, "source": "BOTH", "patterns": ["25X50*", "TTP25XX"]},
        {"name": "30X45  (12 V/T)", "factor": 1.62, "source": "BOTH", "patterns": ["30X45*", "TTP34XX", "DTTP34XX"]},
        {"name": "30X60 ROTIC", "factor": 1.8, "source": "DONG_NAI", "patterns": ["R0XX", "RV0XX", "DR0XX", "*ROTIC*"]},
        {"name": "30X60 TTP", "factor": 1.44, "source": "GRANITE", "patterns": ["TTP36XX", "DTTP36XX", "TTP38XX", "DTTP38XX"]},
        {"name": "30X60 TH", "factor": 1.44, "source": "GRANITE", "patterns": ["TH36XX", "TV36XX", "AV36XX", "PV36XX", "FV36XX", "DT36XX", "DA36XX", "P36XX", "F36XX", "DP36XX", "DF36XX"]},
        {"name": "40X80", "factor": 1.92, "source": "BOTH", "patterns": ["40X80*", "48XX", "P48XX", "F48XX", "TTP48XX"]},
        {"name": "30X30 (11 V/T)", "factor": 0.99, "source": "BOTH", "patterns": ["30X30*", "F34XX", "F30XX", "P30XX", "P34XX"]},
        {"name": "30X30 (16 V/T)", "factor": 1.44, "source": "BOTH", "patterns": ["*16V*", "*16 VIÊN*", "*X*"]},
        {"name": "40X40", "factor": 0.96, "source": "BOTH", "patterns": ["40X40*", "TTP4XX"]},
        {"name": "40X40 (12V/T)", "factor": 1.92, "source": "BOTH", "patterns": ["TTP423XX", "TTP425XX", "TTP4268XX", "TTP473XX", "TTP447XX"]},
        {"name": "40X60", "factor": 1.92, "source": "BOTH", "patterns": ["40X60*"]},
        {"name": "50X50 (6 V/T)", "factor": 1.5, "source": "BOTH", "patterns": ["50X50*", "PZ57XX", "TTP5XX", "F5XX", "M5XX"]},
        {"name": "60x60 (5 V/T)", "factor": 1.8, "source": "DONG_NAI", "patterns": ["60X60*", "P6XX", "F6XX", "TTP6XX"]},
        {"name": "60x60 ĐÁ (5 V/T)", "factor": 1.8, "source": "GRANITE", "patterns": ["P66XX", "F66XX", "TTP66XX"]},
        {"name": "80x80 ĐÁ", "factor": 1.92, "source": "BOTH", "patterns": ["80X80*", "P8XX", "T8XX", "A8XX", "F8XX", "V8XX", "M8XX", "TTP8XX"]}
    ],
    "TMR": [
        {"name": "20x40", "factor": 1.0, "source": "BOTH", "patterns": ["20X40*", "TTP24XX"]},
        {"name": "25x40", "factor": 1.5, "source": "BOTH", "patterns": ["25X40*", "TTP220XX", "TTPV220XX", "TTP250XX", "TTP251XX"]},
        {"name": "25x50", "factor": 1.0, "source": "BOTH", "patterns": ["25X50*", "TTP25XX"]},
        {"name": "30x45", "factor": 1.62, "source": "BOTH", "patterns": ["30X45*", "TTP34XX", "DTTP34XX"]},
        {"name": "30x60 ROTIC", "factor": 1.8, "source": "DONG_NAI", "patterns": ["R0XX", "RV0XX", "DR0XX", "*ROTIC*"]},
        {"name": "30x60 TTP(8v)", "factor": 1.44, "source": "GRANITE", "patterns": ["*HL*", "*_H*", "*-H*", "*8V*"]},
        {"name": "30x60 TTP(10v)", "factor": 1.8, "source": "GRANITE", "patterns": ["TTP36XX", "DTTP36XX", "TTP38XX", "DTTP38XX"]},
        {"name": "30x60 TH(8v)", "factor": 1.44, "source": "GRANITE", "patterns": ["*8V*"]},
        {"name": "30x60 TH(10v)", "factor": 1.8, "source": "GRANITE", "patterns": ["TH36XX", "TV36XX", "MV36XX", "DV36XX", "M36XX", "DM36XX"]},
        {"name": "40x80", "factor": 1.92, "source": "BOTH", "patterns": ["40X80*", "48XX", "M48XX", "TTP48XX"]},
        {"name": "30x30", "factor": 1.44, "source": "BOTH", "patterns": ["30X30*", "*35*X*", "*38*X*"]},
        {"name": "40x40", "factor": 1.44, "source": "BOTH", "patterns": ["40X40*", "TTP4XX"]},
        {"name": "40x60", "factor": 1.92, "source": "BOTH", "patterns": ["40X60*"]},
        {"name": "50x50", "factor": 1.5, "source": "BOTH", "patterns": ["50X50*", "M55XX", "TTP5XX"]},
        {"name": "60x60", "factor": 1.8, "source": "DONG_NAI", "patterns": ["60X60*", "M6XX", "TTP6XX"]},
        {"name": "60x60 (P)", "factor": 1.8, "source": "GRANITE", "patterns": ["M66XX", "TTP66XX"]},
        {"name": "80x80 (P)", "factor": 1.92, "source": "BOTH", "patterns": ["80X80*", "M8XX", "T8XX", "P8XX", "A8XX", "F8XX", "V8XX", "TTP8XX"]}
    ],
    "TRM_TPS": [
        {"name": "20x40", "factor": 0.96, "source": "BOTH", "patterns": ["20X40*", "TTP24XX"]},
        {"name": "25x40 (15v/t)", "factor": 1.5, "source": "BOTH", "patterns": ["25X40*", "TTP220XX", "TTPV220XX"]},
        {"name": "25x50", "factor": 1.0, "source": "BOTH", "patterns": ["25X50*", "TTP25XX"]},
        {"name": "30x60 ROTIC", "factor": 1.8, "source": "DONG_NAI", "patterns": ["R0XX", "RV0XX", "DR0XX", "*ROTIC*"]},
        {"name": "30x60 TTP(10v)", "factor": 1.8, "source": "GRANITE", "patterns": ["TTP36XX", "DTTP36XX", "TTP38XX", "DTTP38XX", "TTPV38XX", "TTPV36XX"]},
        {"name": "30x60(10v) TRM/TPS", "factor": 1.8, "source": "GRANITE", "patterns": ["T36XX", "A36XX", "TV36XX", "AV36XX", "DT36XX", "DA36XX", "TH36XX"]},
        {"name": "40x80", "factor": 1.92, "source": "BOTH", "patterns": ["40X80*", "48XX", "T48XX", "A48XX", "TTP48XX"]},
        {"name": "30x30 (16v/t)", "factor": 1.44, "source": "BOTH", "patterns": ["30X30*", "T30XX", "A30XX", "T38XX", "A38XX"]},
        {"name": "40x40", "factor": 0.96, "source": "BOTH", "patterns": ["40X40*", "TTP4XX"]},
        {"name": "40x40 (12v/t)", "factor": 1.92, "source": "BOTH", "patterns": ["TTP423XX", "TTP425XX", "TTP4268XX", "TTP473XX", "TTP447XX"]},
        {"name": "40x60", "factor": 1.8, "source": "BOTH", "patterns": ["40X60*"]},
        {"name": "50x50 (06v/t)", "factor": 1.5, "source": "BOTH", "patterns": ["50X50*", "A50XX", "TTP5XX", "TZ51XX"]},
        {"name": "60x60 (05v/t)", "factor": 1.8, "source": "DONG_NAI", "patterns": ["60X60*", "T6XX", "A6XX", "TTP6XX", "T6166NL1"]},
        {"name": "60x60.P (05v/t)", "factor": 1.8, "source": "GRANITE", "patterns": ["T66XX", "A66XX", "TTP66XX"]},
        {"name": "80x80.P (03v/t)", "factor": 1.92, "source": "BOTH", "patterns": ["80X80*", "T8XX", "A8XX", "P8XX", "F8XX", "V8XX", "M8XX", "TTP8XX"]}
    ]
}

def get_user_config_dir(bundled_config_dir: Optional[str] = None) -> str:
    """
    Trả về thư mục cấu hình người dùng tại %APPDATA%/VTC_Bao_Cao/config.
    Tách biệt hoàn toàn dữ liệu khách hàng thật và quy luật mã gạch khỏi thư mục cài đặt app
    để không bao giờ bị ghi đè hay mất dữ liệu khi cập nhật phiên bản mới.
    """
    app_data = os.environ.get("APPDATA")
    if app_data:
        user_cfg = os.path.join(app_data, "VTC_Bao_Cao", "config")
    else:
        user_cfg = os.path.expanduser("~/.vtc_bao_cao/config")

    os.makedirs(user_cfg, exist_ok=True)

    if bundled_config_dir and os.path.exists(bundled_config_dir):
        for fname in ["dealers.json", "tile_rules.json", "rules.json"]:
            dest = os.path.join(user_cfg, fname)
            src = os.path.join(bundled_config_dir, fname)
            if not os.path.exists(dest) and os.path.exists(src):
                try:
                    shutil.copy2(src, dest)
                except Exception:
                    pass
    return user_cfg

class RuleManager:
    def __init__(self, config_dir: str):
        self.bundled_config_dir = config_dir
        self.user_config_dir = get_user_config_dir(bundled_config_dir=config_dir)
        self.config_dir = self.user_config_dir

        self.rules_path = os.path.join(self.config_dir, "rules.json")
        self.dealers_path = os.path.join(self.config_dir, "dealers.json")
        self.tile_rules_path = os.path.join(self.config_dir, "tile_rules.json")

        self.factors = {k: v.copy() for k, v in DEFAULT_FACTORS.items()}
        self.source_exceptions = DEFAULT_SOURCE_EXCEPTIONS.copy()
        self.voucher_adjustments = list(DEFAULT_VOUCHER_ADJUSTMENTS)
        self.pending_dealers = list(DEFAULT_PENDING_DEALERS)

        self.dealers = {k: [dict(item) for item in v] for k, v in DEFAULT_DEALERS.items()}
        self.spec_rules = {k: [dict(item) for item in v] for k, v in DEFAULT_SPEC_RULES.items()}

        self.load_rules()
        self.load_dealers()
        self.load_spec_rules()

    def load_rules(self):
        if os.path.exists(self.rules_path):
            try:
                with open(self.rules_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if "factors" in data:
                        self.factors.update(data["factors"])
                    if "source_exceptions" in data:
                        self.source_exceptions.update(data["source_exceptions"])
                    if "voucher_adjustments" in data:
                        self.voucher_adjustments = data["voucher_adjustments"]
                    if "pending_dealers" in data:
                        self.pending_dealers = data["pending_dealers"]
            except Exception as e:
                print(f"Lỗi đọc {self.rules_path}: {e}")

    def save_rules(self):
        os.makedirs(self.config_dir, exist_ok=True)
        data = {
            "factors": self.factors,
            "source_exceptions": self.source_exceptions,
            "voucher_adjustments": self.voucher_adjustments,
            "pending_dealers": self.pending_dealers
        }
        with open(self.rules_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load_dealers(self):
        if os.path.exists(self.dealers_path):
            try:
                with open(self.dealers_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for k in ["TCR_TML", "TMR", "TRM_TPS"]:
                        if k in data:
                            self.dealers[k] = data[k]
            except Exception as e:
                print(f"Lỗi đọc {self.dealers_path}: {e}")
        else:
            self.save_dealers()

    def save_dealers(self):
        os.makedirs(self.config_dir, exist_ok=True)
        with open(self.dealers_path, "w", encoding="utf-8") as f:
            json.dump(self.dealers, f, ensure_ascii=False, indent=2)
        # Danh sách khách hàng thật chỉ lưu bền vững trong user_config_dir (%APPDATA%),
        # tuyệt đối không ghi đè vào bundled_config_dir để tránh rò rỉ dữ liệu lên Git/gói cài đặt.

    def get_dealers(self, report_type: str, brand: Optional[str] = None, only_active: bool = False) -> List[Dict[str, Any]]:
        dealers = self.dealers.get(report_type, [])
        res = []
        for d in dealers:
            if brand and d.get("thuong_hieu") != brand:
                continue
            if only_active and not d.get("is_active", True):
                continue
            res.append(dict(d))
        return res

    def add_or_update_dealer(self, report_type: str, dealer_data: Dict[str, Any]):
        dealers = self.dealers.setdefault(report_type, [])
        code = str(dealer_data.get("ma_khach", "")).strip().upper()
        name = str(dealer_data.get("ten_khach", "")).strip().upper()
        brand = str(dealer_data.get("thuong_hieu", "")).strip().upper()

        found = False
        for i, d in enumerate(dealers):
            d_code = str(d.get("ma_khach", "")).strip().upper()
            d_name = str(d.get("ten_khach", "")).strip().upper()
            d_brand = str(d.get("thuong_hieu", "")).strip().upper()

            # Match by code & brand or name & brand
            if (code and d_code == code and d_brand == brand) or (name and d_name == name and d_brand == brand):
                dealers[i].update(dealer_data)
                found = True
                break

        if not found:
            if "stt" not in dealer_data or not dealer_data["stt"]:
                dealer_data["stt"] = len(dealers) + 1
            if "is_active" not in dealer_data:
                dealer_data["is_active"] = True
            dealers.append(dealer_data)

        self.save_dealers()

    def set_dealer_active(self, report_type: str, ma_khach: str, thuong_hieu: str, is_active: bool):
        dealers = self.dealers.get(report_type, [])
        code = ma_khach.strip().upper()
        th = thuong_hieu.strip().upper()
        for d in dealers:
            if d.get("ma_khach", "").strip().upper() == code and d.get("thuong_hieu", "").strip().upper() == th:
                d["is_active"] = is_active
                break
        self.save_dealers()

    def delete_dealer(self, report_type: str, ma_khach: str, thuong_hieu: Optional[str] = None):
        dealers = self.dealers.get(report_type, [])
        code = ma_khach.strip().upper()
        th = thuong_hieu.strip().upper() if thuong_hieu else None
        self.dealers[report_type] = [
            d for d in dealers
            if not (d.get("ma_khach", "").strip().upper() == code and (th is None or d.get("thuong_hieu", "").strip().upper() == th))
        ]
        self.save_dealers()

    def load_spec_rules(self):
        if os.path.exists(self.tile_rules_path):
            try:
                with open(self.tile_rules_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    for k in ["TCR_TML", "TMR", "TRM_TPS"]:
                        if k in data:
                            self.spec_rules[k] = data[k]
                            # Sync factors dictionary
                            for sp in self.spec_rules[k]:
                                sp_name = sp.get("name")
                                sp_factor = sp.get("factor")
                                if sp_name and sp_factor is not None:
                                    self.factors.setdefault(k, {})[sp_name] = float(sp_factor)
            except Exception as e:
                print(f"Lỗi đọc {self.tile_rules_path}: {e}")
        else:
            self.save_spec_rules()

    def save_spec_rules(self):
        os.makedirs(self.config_dir, exist_ok=True)
        with open(self.tile_rules_path, "w", encoding="utf-8") as f:
            json.dump(self.spec_rules, f, ensure_ascii=False, indent=2)
        if self.bundled_config_dir and os.path.abspath(self.bundled_config_dir) != os.path.abspath(self.config_dir):
            try:
                os.makedirs(self.bundled_config_dir, exist_ok=True)
                with open(os.path.join(self.bundled_config_dir, "tile_rules.json"), "w", encoding="utf-8") as f:
                    json.dump(self.spec_rules, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
        # Also sync to factors in rules.json
        for k, specs in self.spec_rules.items():
            for sp in specs:
                sp_name = sp.get("name")
                sp_fac = sp.get("factor")
                if sp_name and sp_fac is not None:
                    self.factors.setdefault(k, {})[sp_name] = float(sp_fac)
        self.save_rules()

    def get_spec_rules(self, report_type: str) -> List[Dict[str, Any]]:
        return self.spec_rules.get(report_type, [])

    def get_factor(self, report_type: str, spec_name: str) -> float:
        # Check in spec_rules first
        specs = self.spec_rules.get(report_type, [])
        for sp in specs:
            if sp.get("name", "").strip().lower() == spec_name.strip().lower():
                return float(sp.get("factor", 1.0))
        # Fallback to factors dict
        rep_dict = self.factors.get(report_type, {})
        for k, v in rep_dict.items():
            if k.strip().lower() == spec_name.strip().lower():
                return float(v)
        return 1.0

    def check_source_exception(self, ma_khach: str, source_file: str, report_type: str, ten_khach: str = "") -> Tuple[bool, str]:
        code = ma_khach.strip().upper()
        if code in self.source_exceptions:
            file_ex = self.source_exceptions[code].get(source_file)
            if file_ex:
                allowed_reps = file_ex.get("is_valid_for_reports", [])
                if report_type in allowed_reps:
                    return True, file_ex.get("mapped_name", "")
                else:
                    return False, f"Mã {code} tại nguồn {source_file} thuộc '{file_ex.get('dealer_name')}', không hợp lệ cho báo cáo {report_type}"
        return True, ""

    def get_adjustments_for_date(self, target_date_str: str) -> List[Dict[str, Any]]:
        return [adj for adj in self.voucher_adjustments if adj.get("ngay_ct") == target_date_str]

    @staticmethod
    def identify_brand(report_type: str, item_code: str, customer_brand: str) -> str:
        ic = item_code.upper().strip()
        if report_type == "TCR_TML":
            if ic.startswith("P"):
                return "TCR"
            elif ic.startswith("F"):
                return "TML"
            else:
                return customer_brand if customer_brand in ["TCR", "TML"] else "TCR"
        elif report_type == "TMR":
            return "TMR"
        elif report_type == "TRM_TPS":
            if ic.startswith("T") and not ic.startswith("TTP"):
                return "TRM"
            elif ic.startswith("A"):
                return "TPS"
            else:
                return customer_brand if customer_brand in ["TRM", "TPS"] else "TRM"
        return customer_brand

    def match_tile_pattern(self, report_type: str, brand: str, item_code: str, item_name: str) -> Tuple[Optional[str], Optional[str], float, int, bool, List[str]]:
        """
        Tìm quy cách khớp nhất dựa trên danh sách quy tắc mẫu:
        Trả về: (spec_name, matched_pattern, factor, priority, is_conflict, conflict_specs)
        """
        ic = item_code.strip().upper()
        nm = item_name.strip().upper()
        specs = self.spec_rules.get(report_type, [])

        matches: List[Dict[str, Any]] = []

        for sp in specs:
            sp_name = sp.get("name", "")
            fac = float(sp.get("factor", 1.0))
            patterns = sp.get("patterns", [])

            for p_str in patterns:
                p = p_str.strip().upper()
                if not p:
                    continue

                # 1. Exact match
                if p == ic:
                    matches.append({
                        "spec": sp_name,
                        "pattern": p_str,
                        "factor": fac,
                        "priority": 9999,
                        "type": "exact"
                    })
                    continue

                # 2. Wildcard prefix ending with XX (e.g. T66XX)
                if p.endswith("XX") and len(p) >= 3:
                    prefix = p[:-2]
                    if ic.startswith(prefix):
                        matches.append({
                            "spec": sp_name,
                            "pattern": p_str,
                            "factor": fac,
                            "priority": len(prefix),
                            "type": "prefix"
                        })
                    continue

                # 3. Wildcard with *
                if "*" in p:
                    regex = "^" + re.escape(p).replace(r"\*", ".*") + "$"
                    if re.search(regex, ic) or re.search(regex, nm):
                        matches.append({
                            "spec": sp_name,
                            "pattern": p_str,
                            "factor": fac,
                            "priority": len(p.replace("*", "")),
                            "type": "wildcard"
                        })

        if not matches:
            return None, None, 1.0, 0, False, []

        # Find highest priority
        max_priority = max(m["priority"] for m in matches)
        top_matches = [m for m in matches if m["priority"] == max_priority]

        distinct_specs = list({m["spec"] for m in top_matches})
        if len(distinct_specs) > 1:
            # Conflict detected!
            best = top_matches[0]
            return best["spec"], best["pattern"], best["factor"], best["priority"], True, distinct_specs

        best = top_matches[0]
        return best["spec"], best["pattern"], best["factor"], best["priority"], False, []

    def classify_spec(self, report_type: str, item_code: str, item_name: str) -> str:
        """
        Quy chuẩn tên cột quy cách.
        Ưu tiên dùng pattern matching động cấu hình bởi người dùng.
        Nếu không khớp, dùng logic phân loại gốc để đảm bảo tương thích 100%.
        """
        matched_spec, _, _, priority, _, _ = self.match_tile_pattern(report_type, "", item_code, item_name)
        if matched_spec and priority > 0:
            return matched_spec

        # Fallback to hardcoded heuristics if patterns didn't match
        ic = item_code.upper().strip()
        name = item_name.upper()

        if report_type == "TCR_TML":
            if "20X40" in name or ic.startswith("TTP24"):
                return "20X40"
            if "25X40" in name or ic.startswith("TTP220") or ic.startswith("TTPV220") or ic.startswith("TTP250") or ic.startswith("TTP251") or (ic.startswith("TTP25") and "X" in ic):
                return "25X40 (15 V/T)"
            if "25X50" in name or (ic.startswith("TTP25") and not ic.startswith("TTP250") and not ic.startswith("TTP251") and "X" not in ic):
                return "25X50"
            if "30X45" in name or ic.startswith("TTP34") or ic.startswith("DTTP34"):
                return "30X45  (12 V/T)"
            if ic.startswith("R0") or ic.startswith("RV0") or ic.startswith("DR0") or "ROT" in name:
                return "30X60 ROTIC"
            if ("30X60" in name and "TH" in ic) or ic.startswith("TH36") or ic.startswith("TV36") or ic.startswith("AV36") or ic.startswith("PV36") or ic.startswith("FV36") or ic.startswith("DT36") or ic.startswith("DA36") or ic.startswith("P36") or ic.startswith("F36") or ic.startswith("DP36") or ic.startswith("DF36"):
                return "30X60 TH"
            if ("30X60" in name and "TTP" in ic) or (((ic.startswith("TTP36") or ic.startswith("DTTP36") or ic.startswith("TTP38") or ic.startswith("DTTP38"))) and "X" not in ic):
                return "30X60 TTP"
            if "40X80" in name or "48" in ic[:5]:
                return "40X80"
            if "30X30" in name or ic.startswith("F34") or ic.startswith("F30") or ic.startswith("P30") or ic.startswith("P34") or ("30" in ic[:5] and "X" in ic) or ("36" in ic[:5] and "X" in ic) or ("38" in ic[:5] and "X" in ic):
                if "16V" in name or "16 VIÊN" in name or "X" in ic:
                    return "30X30 (16 V/T)"
                return "30X30 (11 V/T)"
            if "40X40" in name or ic.startswith("TTP4"):
                if "12V" in name or "12 VIÊN" in name or ic.startswith("TTP423") or ic.startswith("TTP425") or ic.startswith("TTP4268") or ic.startswith("TTP473") or ic.startswith("TTP447"):
                    return "40X40 (12V/T)"
                return "40X40"
            if "40X60" in name:
                return "40X60"
            if "50X50" in name or ic.startswith("PZ57") or ic.startswith("TTP5") or ic.startswith("F5") or ic.startswith("M5"):
                return "50X50 (6 V/T)"
            is_80_code = (
                ic.startswith(("P8", "T8", "A8", "F8", "V8", "M8", "TTP8", "DP8", "DT8", "DA8", "DF8", "DV8", "DM8", "DTTP8"))
                or ("80X80" in name)
            ) and not ic.startswith(("P6", "T6", "A6", "F6", "V6", "M6", "TTP6", "DP6", "DT6", "DA6", "DF6", "DV6", "DM6", "DTTP6"))
            if is_80_code:
                return "80x80 ĐÁ"
            if "60X60" in name or ic.startswith("P6") or ic.startswith("F6") or ic.startswith("TTP6"):
                m = re.match(r"^([A-Za-z]+)(\d+)", ic)
                digits = m.group(2) if m else ""
                if "PORCELAIN" in name or "ĐÁ" in name or (len(digits) >= 4 and digits.startswith("66")):
                    return "60x60 ĐÁ (5 V/T)"
                return "60x60 (5 V/T)"

        elif report_type == "TMR":
            if "20X40" in name or ic.startswith("TTP24"):
                return "20x40"
            if "25X40" in name or ic.startswith("TTP220") or ic.startswith("TTPV220") or ic.startswith("TTP250") or ic.startswith("TTP251") or (ic.startswith("TTP25") and "X" in ic):
                return "25x40"
            if "25X50" in name or (ic.startswith("TTP25") and not ic.startswith("TTP250") and not ic.startswith("TTP251") and "X" not in ic):
                return "25x50"
            if "30X45" in name or ic.startswith("TTP34") or ic.startswith("DTTP34"):
                return "30x45"
            if ic.startswith("R0") or ic.startswith("RV0") or ic.startswith("DR0") or "ROT" in name:
                return "30x60 ROTIC"
            if ("30X60" in name and "TTP" in ic) or (((ic.startswith("TTP36") or ic.startswith("DTTP36") or ic.startswith("TTP38") or ic.startswith("DTTP38"))) and "X" not in ic):
                is_8v = ("8V" in name or "8 VIÊN" in name or "8V/T" in name or "8 VIÊN/THÙNG" in name or "8V" in name.replace(" ", "") or "HL" in ic or "_H" in ic or "-H" in ic)
                if is_8v:
                    return "30x60 TTP(8v)"
                return "30x60 TTP(10v)"
            if ("30X60" in name and ("TH" in ic or "TMR" in name)) or ic.startswith("TH36") or ic.startswith("TV36") or ic.startswith("MV36") or ic.startswith("DV36") or ic.startswith("M36") or ic.startswith("DM36"):
                is_8v = ("8V" in name or "8 VIÊN" in name or "8V/T" in name or "8 VIÊN/THÙNG" in name or "8V" in name.replace(" ", "") or "HL" in ic or "_H" in ic or "-H" in ic)
                if is_8v:
                    return "30x60 TH(8v)"
                return "30x60 TH(10v)"
            if "40X80" in name or "48" in ic[:5]:
                return "40x80"
            if "30X30" in name or ("35" in ic[:5] and "X" in ic) or ("38" in ic[:5] and "X" in ic):
                return "30x30"
            if "40X40" in name or ic.startswith("TTP4"):
                return "40x40"
            if "40X60" in name:
                return "40x60"
            if "50X50" in name or ic.startswith("M55") or ic.startswith("TTP5"):
                return "50x50"
            is_80_tmr = (
                ic.startswith(("M8", "T8", "P8", "A8", "F8", "V8", "TTP8", "DM8", "DT8", "DP8", "DA8", "DF8", "DV8", "DTTP8"))
                or ("80X80" in name)
            ) and not ic.startswith(("M6", "T6", "P6", "A6", "F6", "V6", "TTP6", "DM6", "DT6", "DP6", "DA6", "DF6", "DV6", "DTTP6"))
            if is_80_tmr:
                return "80x80 (P)"
            if "60X60" in name or ic.startswith("M6") or ic.startswith("TTP6"):
                m = re.match(r"^([A-Za-z]+)(\d+)", ic)
                digits = m.group(2) if m else ""
                if "PORCELAIN" in name or "ĐÁ" in name or (len(digits) >= 4 and digits.startswith("66")):
                    return "60x60 (P)"
                return "60x60"

        elif report_type == "TRM_TPS":
            if "20X40" in name or ic.startswith("TTP24"):
                return "20x40"
            if "25X40" in name or ic.startswith("TTP220") or ic.startswith("TTPV220"):
                return "25x40 (15v/t)"
            if "25X50" in name or ic.startswith("TTP25"):
                return "25x50"
            if ic.startswith("R0") or ic.startswith("RV0") or ic.startswith("DR0") or "ROT" in name:
                return "30x60 ROTIC"
            if ("30X60" in name and "TTP" in ic) or ic.startswith("TTP36") or ic.startswith("DTTP36") or ((ic.startswith("TTP38") or ic.startswith("DTTP38") or ic.startswith("TTPV38") or ic.startswith("TTPV36")) and "X" not in ic):
                return "30x60 TTP(10v)"
            if ("30X60" in name and ("TRM" in name or "TPS" in name)) or ic.startswith("T36") or ic.startswith("A36") or ic.startswith("TV36") or ic.startswith("AV36") or ic.startswith("DT36") or ic.startswith("DA36") or ic.startswith("TH36"):
                return "30x60(10v) TRM/TPS"
            if "40X80" in name or "48" in ic[:5]:
                return "40x80"
            if "30X30" in name or ("30" in ic[:5] and "X" in ic) or ("38" in ic[:5] and "X" in ic):
                return "30x30 (16v/t)"
            if "40X40" in name or ic.startswith("TTP4"):
                if "12V" in name or "12 VIÊN" in name or ic.startswith("TTP423") or ic.startswith("TTP425") or ic.startswith("TTP4268") or ic.startswith("TTP473") or ic.startswith("TTP447"):
                    return "40x40 (12v/t)"
                return "40x40"
            if "40X60" in name:
                return "40x60"
            if "50X50" in name or ic.startswith("A50") or ic.startswith("TTP5") or ic.startswith("TZ51"):
                return "50x50 (06v/t)"
            if "80X80" in name or (len(ic) > 2 and (ic[0] in "TAFPVM" or ic.startswith("TTP")) and "8" in ic[1:4]):
                return "80x80.P (03v/t)"
            if "60X60" in name or ic.startswith("T6") or ic.startswith("A6") or ic.startswith("TTP6"):
                m = re.match(r"^([A-Za-z]+)(\d+)", ic)
                digits = m.group(2) if m else ""
                if "PORCELAIN" in name or "ĐÁ" in name or (len(digits) >= 4 and digits.startswith("66")):
                    return "60x60.P (05v/t)"
                return "60x60 (05v/t)"

        return "KHÁC"

    def test_item_code(
        self,
        report_type: str,
        brand: str,
        item_code: str,
        item_name: str,
        source_name_or_file: str
    ) -> Dict[str, Any]:
        """
        Kiểm tra mã hàng cho hộp thoại 'Thử mã hàng':
        Trả về chi tiết: cột đích, mẫu khớp, hệ số, xung đột, tính hợp lệ nguồn, lý do từ chối.
        """
        ic = item_code.strip().upper()
        matched_spec, matched_pattern, factor, priority, is_conflict, conflict_specs = self.match_tile_pattern(
            report_type, brand, ic, item_name
        )

        if not matched_spec or priority <= 0:
            matched_spec = self.classify_spec(report_type, ic, item_name)
            factor = self.get_factor(report_type, matched_spec)
            matched_pattern = "Mặc định hệ thống"

        src_upper = source_name_or_file.strip().upper()
        is_dn = any(k in src_upper for k in ["ĐỒNG NAI", "DONG NAI", "ĐN", "DN"])
        is_gr = "GRANIT" in src_upper

        is_source_valid = True
        source_reason = "Hợp lệ theo nguồn"

        # Check source constraints
        is_rotic = "ROTIC" in matched_spec.upper() or ic.startswith("R0") or ic.startswith("DR0")
        is_30x60_brand_or_tpp = ("30X60" in matched_spec.upper() and not is_rotic)
        is_60_thuong = matched_spec in ["60x60 (05v/t)", "60x60 (5 V/T)", "60x60"]
        is_60_da = matched_spec in ["60x60.P (05v/t)", "60x60 ĐÁ (5 V/T)", "60x60 (P)"]

        if is_rotic and ("30" in matched_spec or "ROTIC" in matched_spec.upper()):
            if not is_dn:
                is_source_valid = False
                source_reason = "Sai nguồn: 30×60 ROTIC chỉ lấy từ Đồng Nai"
        elif is_30x60_brand_or_tpp:
            if not is_gr:
                is_source_valid = False
                source_reason = "Sai nguồn: 30×60 thương hiệu và TPP chỉ lấy từ GRANIT"
        elif is_60_thuong:
            if not is_dn:
                is_source_valid = False
                source_reason = "Sai nguồn: 60×60 thường chỉ lấy từ Đồng Nai"
        elif is_60_da:
            if not is_gr:
                is_source_valid = False
                source_reason = "Sai nguồn: 60×60.P (Porcelain) chỉ lấy từ GRANITE"

        return {
            "item_code": ic,
            "target_spec": matched_spec,
            "matched_pattern": matched_pattern,
            "factor": factor,
            "priority": priority,
            "is_conflict": is_conflict,
            "conflict_specs": conflict_specs,
            "is_source_valid": is_source_valid,
            "source_reason": source_reason
        }
