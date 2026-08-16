"""
Auto and Manual Update Module for ULN / JSON to DOCX Converter.
Integrates with GitHub Releases (CallMeMrPenguin/Docx-Converter).
"""

import os
import sys
import json
import time
import urllib.request
import urllib.error
import subprocess
import tempfile
import webbrowser
from typing import Dict, Any, Optional, Tuple, Callable

DEFAULT_REPO = "CallMeMrPenguin/Docx-Converter"
FALLBACK_VERSION = "1.0.0"

def get_bundle_dir() -> str:
    """Returns the base directory whether running from source or frozen executable."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

def get_current_version() -> str:
    """Reads the current version from the bundled VERSION file or fallback."""
    try:
        ver_file = os.path.join(get_bundle_dir(), "VERSION")
        if os.path.exists(ver_file):
            with open(ver_file, "r", encoding="utf-8") as f:
                ver = f.read().strip()
                if ver:
                    return ver
    except Exception:
        pass
    
    # Check parent/cwd if running from source
    try:
        local_ver = os.path.join(os.path.dirname(os.path.abspath(__file__)), "VERSION")
        if os.path.exists(local_ver):
            with open(local_ver, "r", encoding="utf-8") as f:
                ver = f.read().strip()
                if ver:
                    return ver
    except Exception:
        pass

    return FALLBACK_VERSION

def parse_version(v_str: str) -> Tuple[int, ...]:
    """Parses a version string like 'v1.2.3' or '1.2' into a tuple of ints."""
    clean = v_str.strip().lstrip("vV")
    parts = []
    for chunk in clean.split("."):
        try:
            num = int(''.join(filter(str.isdigit, chunk)) or 0)
            parts.append(num)
        except ValueError:
            parts.append(0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts)

def is_newer_version(latest_ver: str, current_ver: str) -> bool:
    """Checks if latest_ver is strictly greater than current_ver."""
    return parse_version(latest_ver) > parse_version(current_ver)

def check_for_updates(repo: str = DEFAULT_REPO) -> Dict[str, Any]:
    """
    Checks GitHub Releases for updates.
    Returns a dict with update details.
    """
    current_ver = get_current_version()
    result: Dict[str, Any] = {
        "success": False,
        "has_update": False,
        "current_version": current_ver,
        "latest_version": current_ver,
        "release_title": "",
        "release_notes": "",
        "release_url": f"https://github.com/{repo}/releases",
        "download_url": None,
        "asset_name": None,
        "error": None
    }

    # 1. Try GitHub API for Latest Release
    api_url = f"https://api.github.com/repos/{repo}/releases/latest"
    headers = {
        "User-Agent": "Docx-Converter-Updater/1.0",
        "Accept": "application/vnd.github.v3+json"
    }

    req = urllib.request.Request(api_url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=6) as response:
            if response.status == 200:
                data = json.loads(response.read().decode("utf-8"))
                tag_name = data.get("tag_name", "").strip()
                latest_ver = tag_name.lstrip("vV") if tag_name else current_ver
                
                result["latest_version"] = latest_ver
                result["release_title"] = data.get("name", f"Release {tag_name}")
                result["release_notes"] = data.get("body", "").strip() or "Không có mô tả chi tiết."
                result["release_url"] = data.get("html_url", result["release_url"])
                result["success"] = True

                # Look for downloadable .exe asset
                for asset in data.get("assets", []):
                    name = asset.get("name", "")
                    if name.lower().endswith(".exe"):
                        result["download_url"] = asset.get("browser_download_url")
                        result["asset_name"] = name
                        break

                if is_newer_version(latest_ver, current_ver):
                    result["has_update"] = True
                return result
    except urllib.error.HTTPError as he:
        # If 404 (no releases yet) or 403 (API rate limit), try raw VERSION check
        pass
    except Exception as e:
        result["error"] = str(e)

    # 2. Fallback: Check raw VERSION file from main branch
    raw_url = f"https://raw.githubusercontent.com/{repo}/main/VERSION"
    try:
        raw_req = urllib.request.Request(raw_url, headers={"User-Agent": "Docx-Converter-Updater/1.0"})
        with urllib.request.urlopen(raw_req, timeout=5) as response:
            if response.status == 200:
                raw_ver = response.read().decode("utf-8").strip()
                if raw_ver:
                    result["latest_version"] = raw_ver
                    result["release_title"] = f"Phiên bản {raw_ver}"
                    result["release_notes"] = f"Bản cập nhật mới nhất {raw_ver} đã có sẵn trên GitHub."
                    result["success"] = True
                    if is_newer_version(raw_ver, current_ver):
                        result["has_update"] = True
                    return result
    except Exception as e:
        if not result["error"]:
            result["error"] = str(e)

    return result

def download_and_install_update(
    download_url: Optional[str],
    release_url: str,
    progress_callback: Optional[Callable[[int, int], None]] = None
) -> Tuple[bool, str]:
    """
    Downloads the updated .exe and replaces the current running binary via transient batch script.
    If not running as frozen exe or no direct download asset, opens the browser to the release page.
    """
    is_frozen = getattr(sys, 'frozen', False)
    current_exe = os.path.abspath(sys.executable)

    # If running from python script or no direct exe url, open browser
    if not is_frozen or not download_url:
        try:
            webbrowser.open(release_url)
            return True, "Đã mở trang tải bản cập nhật trên trình duyệt."
        except Exception as e:
            return False, f"Không thể mở trình duyệt: {e}"

    temp_dir = tempfile.gettempdir()
    download_target = os.path.join(temp_dir, f"DocxConverter_Update_{int(time.time())}.exe")

    try:
        # Download new executable with progress reporting
        req = urllib.request.Request(download_url, headers={"User-Agent": "Docx-Converter-Updater/1.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            total_size = int(resp.headers.get('content-length', 0))
            downloaded = 0
            block_size = 65536

            with open(download_target, "wb") as out_f:
                while True:
                    chunk = resp.read(block_size)
                    if not chunk:
                        break
                    out_f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total_size > 0:
                        progress_callback(downloaded, total_size)

        if not os.path.exists(download_target) or os.path.getsize(download_target) < 1000:
            return False, "File tải về không hợp lệ hoặc bị lỗi."

        # Create transient batch script to swap running exe and restart
        bat_path = os.path.join(temp_dir, f"update_swap_{int(time.time())}.bat")
        current_pid = os.getpid()

        bat_script = f"""@echo off
setlocal
chcp 65001 >nul
timeout /t 1 /nobreak >nul

:wait_pid
tasklist /fi "PID eq {current_pid}" 2>nul | find "{current_pid}" >nul
if %ERRORLEVEL% equ 0 (
    timeout /t 1 /nobreak >nul
    goto wait_pid
)

timeout /t 1 /nobreak >nul
copy /y "{download_target}" "{current_exe}" >nul
if %ERRORLEVEL% neq 0 (
    move /y "{download_target}" "{current_exe}" >nul
)

del /f /q "{download_target}" >nul 2>&1
start "" "{current_exe}"
del /f /q "%~f0" >nul 2>&1
exit
"""
        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(bat_script)

        # Launch detached updater script
        subprocess.Popen(
            ["cmd.exe", "/c", bat_path],
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0,
            close_fds=True
        )

        # Exit current app immediately so the script can replace it
        sys.exit(0)

    except Exception as e:
        if os.path.exists(download_target):
            try:
                os.remove(download_target)
            except Exception:
                pass
        return False, f"Lỗi khi cập nhật tự động: {e}"
