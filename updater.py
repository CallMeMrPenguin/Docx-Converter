"""
Auto and Manual Update Module for ULN / JSON to DOCX Converter.
Integrates with GitHub Releases, Tags, and Raw Repository (CallMeMrPenguin/Docx-Converter).
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
from typing import Dict, Any, Optional, Tuple, Callable, List

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

def get_github_token() -> Optional[str]:
    """Retrieves GitHub token from env or git credentials if available."""
    # Check env vars first
    for key in ("GITHUB_TOKEN", "GH_TOKEN", "DOCX_CONVERTER_GITHUB_TOKEN"):
        t = os.environ.get(key)
        if t:
            return t.strip()
    
    # Try Git Credential Manager (works on developer/authenticated machines)
    try:
        proc = subprocess.Popen(
            ['git', 'credential', 'fill'],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        out, _ = proc.communicate(input='protocol=https\nhost=github.com\n\n', timeout=2)
        for line in out.strip().split('\n'):
            if line.startswith('password='):
                val = line.split('=', 1)[1].strip()
                if val:
                    return val
    except Exception:
        pass
    return None

def parse_version(v_str: str) -> Tuple[int, ...]:
    """Parses a version string like 'v1.2.3' or '1.2' into a tuple of ints."""
    if not v_str:
        return (0, 0, 0)
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
    Checks GitHub Releases, Tags, and raw VERSION file for updates.
    Returns a dict with comprehensive update details.
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

    token = get_github_token()
    headers = {
        "User-Agent": "Docx-Converter-Updater/1.0",
        "Accept": "application/vnd.github.v3+json"
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"

    raw_main_ver: Optional[str] = None
    releases_data: List[Dict[str, Any]] = []
    tags_data: List[str] = []
    errors: List[str] = []

    # 1. Fetch raw VERSION from main branch (bypasses GitHub API rate limits)
    try:
        raw_url = f"https://raw.githubusercontent.com/{repo}/main/VERSION"
        raw_req = urllib.request.Request(raw_url, headers={"User-Agent": "Docx-Converter-Updater/1.0"})
        with urllib.request.urlopen(raw_req, timeout=5) as resp:
            if resp.status == 200:
                val = resp.read().decode("utf-8").strip()
                if val:
                    raw_main_ver = val
    except Exception as e:
        errors.append(f"Raw VERSION: {e}")

    # 2. Fetch GitHub Releases list (finds all releases, including latest and assets)
    try:
        api_url = f"https://api.github.com/repos/{repo}/releases"
        req = urllib.request.Request(api_url, headers=headers)
        with urllib.request.urlopen(req, timeout=6) as resp:
            if resp.status == 200:
                releases_data = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        errors.append(f"Releases API: {e}")

    # 3. Fetch GitHub Tags list (in case tags exist without explicit releases)
    try:
        tags_url = f"https://api.github.com/repos/{repo}/tags"
        req = urllib.request.Request(tags_url, headers=headers)
        with urllib.request.urlopen(req, timeout=6) as resp:
            if resp.status == 200:
                tags_json = json.loads(resp.read().decode("utf-8"))
                tags_data = [t.get("name", "").strip().lstrip("vV") for t in tags_json if t.get("name")]
    except Exception as e:
        errors.append(f"Tags API: {e}")

    # Find highest version among releases
    best_release: Optional[Dict[str, Any]] = None
    highest_rel_ver: Optional[str] = None
    for rel in releases_data:
        tag = rel.get("tag_name", "").strip().lstrip("vV")
        if not tag:
            continue
        if highest_rel_ver is None or is_newer_version(tag, highest_rel_ver):
            highest_rel_ver = tag
            best_release = rel

    # Find highest version among tags
    highest_tag_ver: Optional[str] = None
    for t_ver in tags_data:
        if highest_tag_ver is None or is_newer_version(t_ver, highest_tag_ver):
            highest_tag_ver = t_ver

    # Determine absolute highest candidate version
    candidates: List[str] = []
    if raw_main_ver:
        candidates.append(raw_main_ver)
    if highest_rel_ver:
        candidates.append(highest_rel_ver)
    if highest_tag_ver:
        candidates.append(highest_tag_ver)

    if not candidates:
        if errors:
            result["error"] = " | ".join(errors)
        return result

    best_ver = max(candidates, key=parse_version)
    result["latest_version"] = best_ver
    result["success"] = True

    # If the best version matches or is covered by a release with an executable asset
    matched_release = None
    for rel in releases_data:
        tag = rel.get("tag_name", "").strip().lstrip("vV")
        if tag == best_ver:
            matched_release = rel
            break

    # If exact release not found for best_ver, fallback to the latest release available
    target_rel = matched_release or best_release

    if target_rel:
        tag_name = target_rel.get("tag_name", f"v{best_ver}")
        result["release_title"] = target_rel.get("name", f"Release {tag_name}")
        result["release_notes"] = target_rel.get("body", "").strip() or "Bản cập nhật mới trên GitHub."
        result["release_url"] = target_rel.get("html_url", result["release_url"])

        # Look for downloadable .exe asset
        for asset in target_rel.get("assets", []):
            name = asset.get("name", "")
            if name.lower().endswith(".exe"):
                # Use public browser_download_url by default for direct reliable downloads
                result["download_url"] = asset.get("browser_download_url") or asset.get("url")
                result["asset_name"] = name
                break
    else:
        result["release_title"] = f"Phiên bản v{best_ver}"
        result["release_notes"] = f"Đã có phiên bản mới v{best_ver} trên GitHub.\n\nBản cập nhật mã nguồn mới nhất đã sẵn sàng."
        result["release_url"] = f"https://github.com/{repo}/releases"

    if is_newer_version(best_ver, current_ver):
        result["has_update"] = True

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

    token = get_github_token()
    headers = {
        "User-Agent": "Docx-Converter-Updater/1.0",
        "Accept": "*/*"
    }
    # Only attach Authorization if hitting api.github.com, not public github.com release assets
    if token and "api.github.com" in download_url:
        headers["Authorization"] = f"Bearer {token}"
        headers["Accept"] = "application/octet-stream"

    try:
        # Download new executable with progress reporting
        req = urllib.request.Request(download_url, headers=headers)
        with urllib.request.urlopen(req, timeout=40) as resp:
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
