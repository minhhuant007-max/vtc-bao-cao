import os
import sys
import json
import urllib.request
import urllib.error
import subprocess
import shutil
import tempfile
from typing import Dict, Any, Optional, Tuple
from .version import APP_VERSION, DEFAULT_GITHUB_REPO

def parse_version_tuple(v_str: str) -> Tuple[int, ...]:
    """Parse version string like 'v1.0.1' or '1.2.3' into tuple (1, 0, 1)"""
    v_clean = v_str.strip().lower().lstrip('v')
    parts = []
    for part in v_clean.split('.'):
        digits = "".join(ch for ch in part if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)

def get_local_access_token() -> Optional[str]:
    """
    Lấy token truy cập GitHub từ biến môi trường máy hoặc cấu hình cá nhân tại %APPDATA%.
    Tuyệt đối không nhúng token vào EXE để đảm bảo an toàn bảo mật.
    """
    for env_k in ["GITHUB_TOKEN", "GH_TOKEN"]:
        val = os.environ.get(env_k)
        if val:
            return val.strip()

    appdata = os.environ.get("APPDATA")
    if appdata:
        cfg_file = os.path.join(appdata, "VTC_Bao_Cao", "config", "settings.json")
        if os.path.exists(cfg_file):
            try:
                with open(cfg_file, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    tok = cfg.get("github_token")
                    if tok:
                        return tok.strip()
            except Exception:
                pass
    return None

def get_release_via_gh_cli(repo: str) -> Optional[Dict[str, Any]]:
    """Truy vấn GitHub Releases thông qua GitHub CLI (gh) đã đăng nhập trên máy mà không cần token nhúng."""
    try:
        cmd = ["gh", "release", "view", "--repo", repo, "--json", "tagName,body,publishedAt,assets"]
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=8)
        if p.returncode == 0 and p.stdout:
            return json.loads(p.stdout)
    except Exception:
        pass
    return None

def check_for_updates(repo: str = DEFAULT_GITHUB_REPO, timeout: int = 8) -> Dict[str, Any]:
    """
    Kiểm tra phiên bản mới nhất từ GitHub Releases.
    Hỗ trợ cả kho lưu trữ công khai và kho riêng tư (thông qua GitHub CLI hoặc token máy cá nhân).
    """
    res = {
        "has_update": False,
        "current_version": APP_VERSION,
        "latest_version": APP_VERSION,
        "release_notes": "",
        "download_url": "",
        "asset_name": "",
        "published_at": "",
        "error": None
    }

    # 1. Thử qua GitHub CLI (gh) nếu máy đã cài và đăng nhập
    gh_data = get_release_via_gh_cli(repo)
    if gh_data:
        tag = gh_data.get("tagName", "").strip()
        res["latest_version"] = tag.lstrip('v')
        res["release_notes"] = gh_data.get("body", "")
        res["published_at"] = gh_data.get("publishedAt", "")[:10]

        curr_tuple = parse_version_tuple(APP_VERSION)
        latest_tuple = parse_version_tuple(tag)
        if latest_tuple > curr_tuple:
            res["has_update"] = True

        assets = gh_data.get("assets", [])
        chosen_asset = None
        for ast in assets:
            name = ast.get("name", "").lower()
            if "x64" in name or "update" in name or name.endswith(".zip") or name.endswith(".exe"):
                chosen_asset = ast
                break
        if not chosen_asset and assets:
            chosen_asset = assets[0]

        if chosen_asset:
            res["download_url"] = chosen_asset.get("url", "")
            res["asset_name"] = chosen_asset.get("name", "")
        elif res["has_update"]:
            res["download_url"] = f"https://github.com/{repo}/releases/tag/{tag}"
            res["asset_name"] = "Xem trên GitHub"
        return res

    # 2. Thử qua GitHub API với token máy hoặc không token
    api_url = f"https://api.github.com/repos/{repo}/releases/latest"
    headers = {
        "User-Agent": f"VTC-Bao-Cao-App/{APP_VERSION}",
        "Accept": "application/vnd.github.v3+json"
    }
    tok = get_local_access_token()
    if tok:
        headers["Authorization"] = f"Bearer {tok}"

    req = urllib.request.Request(api_url, headers=headers)
    
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            if response.status != 200:
                res["error"] = f"GitHub trả về mã lỗi HTTP {response.status}"
                return res
            data = json.loads(response.read().decode("utf-8"))
            
            tag = data.get("tag_name", "").strip()
            res["latest_version"] = tag.lstrip('v')
            res["release_notes"] = data.get("body", "")
            res["published_at"] = data.get("published_at", "")[:10]
            
            # So sánh phiên bản
            curr_tuple = parse_version_tuple(APP_VERSION)
            latest_tuple = parse_version_tuple(tag)
            
            if latest_tuple > curr_tuple:
                res["has_update"] = True
                
            # Tìm asset Windows x64 (.zip hoặc .exe)
            assets = data.get("assets", [])
            chosen_asset = None
            for ast in assets:
                name = ast.get("name", "").lower()
                if "x64" in name or "update" in name or name.endswith(".zip") or name.endswith(".exe"):
                    chosen_asset = ast
                    break
            if not chosen_asset and assets:
                chosen_asset = assets[0]
                
            if chosen_asset:
                res["download_url"] = chosen_asset.get("browser_download_url", "")
                res["asset_name"] = chosen_asset.get("name", "")
            elif res["has_update"]:
                res["download_url"] = data.get("html_url", "")
                res["asset_name"] = "Xem trên GitHub"

    except urllib.error.HTTPError as e:
        if e.code == 404:
            res["error"] = (
                "Kho lưu trữ GitHub riêng tư yêu cầu quyền truy cập.\n"
                "Bạn có thể:\n"
                "1. Đăng nhập GitHub CLI qua lệnh 'gh auth login' trên máy này, HOẶC\n"
                "2. Thêm biến môi trường GITHUB_TOKEN (hoặc cấu hình tại %APPDATA%), HOẶC\n"
                f"3. Truy cập trực tiếp qua trình duyệt: https://github.com/{repo}/releases"
            )
        else:
            res["error"] = f"Lỗi kết nối GitHub (HTTP {e.code}): {e.reason}"
    except urllib.error.URLError as e:
        res["error"] = f"Không thể kết nối Internet hoặc GitHub: {e.reason}"
    except Exception as e:
        res["error"] = f"Lỗi kiểm tra bản cập nhật: {str(e)}"
        
    return res

def download_asset(download_url: str, dest_dir: str, repo: str = DEFAULT_GITHUB_REPO, tag: Optional[str] = None, progress_callback=None) -> str:
    """
    Tải gói cập nhật từ GitHub về thư mục tạm.
    Hỗ trợ GitHub CLI (gh) hoặc token máy cá nhân hoặc tải trực tiếp.
    """
    os.makedirs(dest_dir, exist_ok=True)
    filename = download_url.split("/")[-1] or "update_package.zip"
    dest_path = os.path.join(dest_dir, filename)

    # Tự động trích xuất tag từ download_url nếu chưa truyền vào
    if not tag and "/download/" in download_url:
        try:
            parts = download_url.split("/download/")[1].split("/")
            if len(parts) >= 1:
                tag = parts[0]
        except Exception:
            pass

    # 1. Thử tải qua gh release download nếu gh có sẵn trên máy và có tag
    if tag:
        try:
            gh_cmd = ["gh", "release", "download", tag, "--pattern", "*.zip", "--dir", dest_dir, "--repo", repo, "--clobber"]
            p = subprocess.run(gh_cmd, capture_output=True, text=True, timeout=120)
            if p.returncode == 0:
                for f in os.listdir(dest_dir):
                    if f.endswith(".zip") and (tag in f or "VTC" in f):
                        return os.path.join(dest_dir, f)
        except Exception:
            pass

    # 2. Tải qua HTTP với token máy cá nhân nếu có
    headers = {
        "User-Agent": f"VTC-Bao-Cao-App/{APP_VERSION}",
        "Accept": "application/octet-stream"
    }
    tok = get_local_access_token()
    if tok:
        headers["Authorization"] = f"Bearer {tok}"

    req = urllib.request.Request(download_url, headers=headers)
    with urllib.request.urlopen(req, timeout=60) as resp, open(dest_path, "wb") as out_file:
        total_size = int(resp.headers.get("Content-Length", 0))
        downloaded = 0
        block_size = 65536
        while True:
            buffer = resp.read(block_size)
            if not buffer:
                break
            downloaded += len(buffer)
            out_file.write(buffer)
            if progress_callback and total_size > 0:
                percent = int(downloaded * 100 / total_size)
                progress_callback(percent, downloaded, total_size)

    return dest_path

def create_and_launch_updater_script(app_dir: str, update_file_path: str, pid: int) -> bool:
    """
    Tạo script PowerShell cập nhật độc lập:
    1. Đợi ứng dụng chính (PID) thoát hẳn.
    2. Sao lưu phiên bản hiện tại sang thư mục `_backup_prev`.
    3. Giải nén hoặc sao chép phiên bản mới đè lên `app_dir`.
    4. Nếu thất bại, khôi phục lại từ `_backup_prev`.
    5. Khởi động lại `VTC_Bao_Cao.exe`.
    """
    updater_dir = os.path.join(tempfile.gettempdir(), "vtc_app_updater")
    os.makedirs(updater_dir, exist_ok=True)
    ps1_path = os.path.join(updater_dir, "run_update.ps1")
    
    exe_target = os.path.join(app_dir, "VTC_Bao_Cao.exe")
    backup_dir = os.path.join(app_dir, "..", "_backup_prev")
    
    ps_content = f"""
$ErrorActionPreference = 'Stop'
$pidToWait = {pid}
$appDir = "{app_dir.replace('\\', '/')}"
$updateFile = "{update_file_path.replace('\\', '/')}"
$exeTarget = "{exe_target.replace('\\', '/')}"
$backupDir = "{backup_dir.replace('\\', '/')}"
$logFile = "$env:TEMP/vtc_updater.log"

"Starting update process at $(Get-Date)..." | Out-File -FilePath $logFile -Encoding utf8

# 1. Chờ ứng dụng chính đóng hẳn
"Waiting for PID $pidToWait to exit..." | Out-File -FilePath $logFile -Append -Encoding utf8
$timeout = 20
while ($timeout -gt 0) {{
    $proc = Get-Process -Id $pidToWait -ErrorAction SilentlyContinue
    if (-not $proc) {{ break }}
    Start-Sleep -Seconds 1
    $timeout--
}}

Start-Sleep -Seconds 2

try {{
    # 2. Tạo bản sao lưu ứng dụng cũ
    "Backing up current app to $backupDir..." | Out-File -FilePath $logFile -Append -Encoding utf8
    if (Test-Path $backupDir) {{ Remove-Item -Path $backupDir -Recurse -Force }}
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    Copy-Item -Path "$appDir/*" -Destination $backupDir -Recurse -Force -Exclude "config"

    # 3. Áp dụng bản cập nhật
    "Applying update from $updateFile..." | Out-File -FilePath $logFile -Append -Encoding utf8
    if ($updateFile.EndsWith(".zip")) {{
        Expand-Archive -Path $updateFile -DestinationPath $appDir -Force
    }} elseif ($updateFile.EndsWith(".exe") -and $updateFile -ne $exeTarget) {{
        Copy-Item -Path $updateFile -Destination $exeTarget -Force
    }}

    "Update applied successfully!" | Out-File -FilePath $logFile -Append -Encoding utf8
}} catch {{
    "ERROR during update: $_. Restoring previous version..." | Out-File -FilePath $logFile -Append -Encoding utf8
    if (Test-Path $backupDir) {{
        Copy-Item -Path "$backupDir/*" -Destination $appDir -Recurse -Force
    }}
}} finally {{
    # 4. Khởi động lại ứng dụng
    "Relaunching $exeTarget..." | Out-File -FilePath $logFile -Append -Encoding utf8
    if (Test-Path $exeTarget) {{
        Start-Process -FilePath $exeTarget
    }}
}}
"""
    with open(ps1_path, "w", encoding="utf-8") as f:
        f.write(ps_content)
        
    cmd = [
        "powershell.exe",
        "-ExecutionPolicy", "Bypass",
        "-WindowStyle", "Hidden",
        "-File", ps1_path
    ]
    subprocess.Popen(cmd, creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0)
    return True
