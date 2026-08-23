"""
Build script to compile JSON/ULN to DOCX Converter into a standalone Portable Executable.
"""

import os
import sys
import subprocess
import shutil

def build_portable():
    print("=" * 60)
    print("  BUILDING PORTABLE STANDALONE EXECUTABLE (DOCX CONVERTER)  ")
    print("=" * 60)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(base_dir, "dist")
    build_dir = os.path.join(base_dir, "build")

    # Read Version
    version_path = os.path.join(base_dir, "VERSION")
    version_str = "1.0.0"
    if os.path.exists(version_path):
        with open(version_path, "r", encoding="utf-8") as f:
            version_str = f.read().strip()
    print(f"[*] Target Application Version: v{version_str}")

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name", "DocxConverter_Portable",
        "--add-data", f"{os.path.join(base_dir, 'VERSION')};.",
        "--add-data", f"{os.path.join(base_dir, 'uln_test.txt')};.",
        "--add-data", f"{os.path.join(base_dir, 'prompt.txt')};.",
        "--hidden-import", "pythoncom",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "win32timezone",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "pillow_heif",
        "--hidden-import", "pillow_avif",
        "--collect-all", "pillow_heif",
        "--collect-all", "pillow_avif",
        "--hidden-import", "updater",
        "--hidden-import", "uln_parser",
        "--hidden-import", "uln_renderer",
        "--hidden-import", "uln_compiler",
        "--hidden-import", "renderer_utils",
        "--hidden-import", "renderer_blocks",
        "--hidden-import", "gui_styles",
        "--hidden-import", "gui_prompt_editor",
        "--hidden-import", "gui_update_modal",
        "--hidden-import", "gui_docx_picker",
        "--hidden-import", "gui_image_preview",
        "--collect-all", "renderers",
        os.path.join(base_dir, "gui_app.py")
    ]

    print(f"[*] Running PyInstaller...")
    ret = subprocess.run(cmd, cwd=base_dir)
    if ret.returncode != 0:
        print("[!] PyInstaller build failed!")
        sys.exit(ret.returncode)

    exe_path = os.path.join(dist_dir, "DocxConverter_Portable.exe")
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print("=" * 60)
        print("[OK] BUILD SUCCESSFUL!")
        print(f"[OK] Portable Executable: {exe_path}")
        print(f"[OK] File Size: {size_mb:.2f} MB")
        print("=" * 60)
    else:
        print("[!] Executable was not found in dist/ folder.")
        sys.exit(1)

if __name__ == "__main__":
    build_portable()
