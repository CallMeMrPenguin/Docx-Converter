"""
Automated GitHub Release & Binary Asset Publisher for ULN / Docx Converter.
Builds the standalone portable executable and publishes the release with direct assets to GitHub.
"""

import os
import sys
import json
import time
import urllib.request
import urllib.error
import subprocess
from typing import Optional, Dict, Any

import updater
import build_portable

DEFAULT_REPO = "CallMeMrPenguin/Docx-Converter"

def publish_release(repo: str = DEFAULT_REPO, build_first: bool = True):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(base_dir, "dist")
    exe_path = os.path.join(dist_dir, "DocxConverter_Portable.exe")
    version_str = updater.get_current_version()
    tag_name = f"v{version_str}"

    print("=" * 65)
    print(f"  PUBLISHING RELEASE: {tag_name} to GitHub ({repo})")
    print("=" * 65)

    # 1. Build portable exe if requested or missing
    if build_first or not os.path.exists(exe_path):
        print(f"[*] Compiling Portable Executable (v{version_str})...")
        build_portable.build_portable()
    
    if not os.path.exists(exe_path):
        print(f"[!] Error: Executable not found at {exe_path}")
        sys.exit(1)

    file_size = os.path.getsize(exe_path)
    file_size_mb = file_size / (1024 * 1024)
    print(f"[*] Executable ready: {exe_path} ({file_size_mb:.2f} MB)")

    # 2. Get GitHub Token
    token = updater.get_github_token()
    if not token:
        print("[!] Error: No GitHub authentication token found.")
        print("    Please ensure Git Credential Manager is active or set GITHUB_TOKEN environment variable.")
        sys.exit(1)

    headers = {
        "User-Agent": "Docx-Converter-Publisher/1.0",
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    # 3. Check if release already exists
    print(f"[*] Checking existing GitHub Releases for {tag_name}...")
    release_info: Optional[Dict[str, Any]] = None
    try:
        req = urllib.request.Request(
            f"https://api.github.com/repos/{repo}/releases/tags/{tag_name}",
            headers=headers
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            if resp.status == 200:
                release_info = json.loads(resp.read().decode("utf-8"))
                print(f"[+] Found existing release ID: {release_info.get('id')}")
    except urllib.error.HTTPError as he:
        if he.code != 404:
            print(f"[!] HTTP Error checking release: {he}")
    except Exception as e:
        print(f"[!] Error checking release: {e}")

    # 4. Create release if not exists
    if not release_info:
        print(f"[*] Creating GitHub Release for {tag_name}...")
        payload = {
            "tag_name": tag_name,
            "target_commitish": "main",
            "name": f"Docx Converter {tag_name} - Portable Edition",
            "body": f"## ULN to DOCX Converter {tag_name} (Portable)\n\n"
                    f"- Standalone Portable Executable (bấm là chạy không cần cài đặt).\n"
                    f"- Tự động đồng bộ và hỗ trợ cập nhật trực tiếp trong ứng dụng.\n"
                    f"- Tối ưu hoá căn lề bảng, typography và trích xuất hình ảnh.",
            "draft": False,
            "prerelease": False
        }
        create_req = urllib.request.Request(
            f"https://api.github.com/repos/{repo}/releases",
            data=json.dumps(payload).encode("utf-8"),
            headers=headers,
            method="POST"
        )
        try:
            with urllib.request.urlopen(create_req, timeout=15) as resp:
                release_info = json.loads(resp.read().decode("utf-8"))
                print(f"[+] Created Release ID: {release_info.get('id')}")
        except Exception as e:
            print(f"[!] Failed to create release: {e}")
            sys.exit(1)

    release_id = release_info.get("id")
    upload_url_raw = release_info.get("upload_url", "")
    upload_url_base = upload_url_raw.split("{")[0] if "{" in upload_url_raw else f"https://uploads.github.com/repos/{repo}/releases/{release_id}/assets"

    # 5. Delete existing asset with same name if already present
    asset_name = "DocxConverter_Portable.exe"
    for asset in release_info.get("assets", []):
        if asset.get("name") == asset_name:
            asset_id = asset.get("id")
            print(f"[*] Deleting previous asset ID {asset_id} to re-upload...")
            del_req = urllib.request.Request(
                f"https://api.github.com/repos/{repo}/releases/assets/{asset_id}",
                headers=headers,
                method="DELETE"
            )
            try:
                with urllib.request.urlopen(del_req, timeout=10) as resp:
                    print("[+] Previous asset deleted successfully.")
            except Exception as e:
                print(f"[!] Warning deleting old asset: {e}")

    # 6. Upload Binary Asset
    print(f"[*] Uploading {asset_name} ({file_size_mb:.2f} MB) to GitHub Release...")
    upload_url = f"{upload_url_base}?name={asset_name}"
    
    upload_headers = {
        "User-Agent": "Docx-Converter-Publisher/1.0",
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream",
        "Content-Length": str(file_size)
    }

    with open(exe_path, "rb") as f:
        data = f.read()

    upload_req = urllib.request.Request(
        upload_url,
        data=data,
        headers=upload_headers,
        method="POST"
    )

    try:
        with urllib.request.urlopen(upload_req, timeout=120) as resp:
            if resp.status in (200, 201):
                asset_res = json.loads(resp.read().decode("utf-8"))
                print("=" * 65)
                print("[OK] RELEASE PUBLISHED & ASSET UPLOADED SUCCESSFULLY!")
                print(f"[OK] Release URL: {release_info.get('html_url')}")
                print(f"[OK] Download URL: {asset_res.get('browser_download_url')}")
                print("=" * 65)
            else:
                print(f"[!] Upload returned status {resp.status}")
    except Exception as e:
        print(f"[!] Failed to upload release asset: {e}")
        sys.exit(1)

if __name__ == "__main__":
    rebuild = "--no-build" not in sys.argv
    publish_release(build_first=rebuild)
