"""
start.py - Trình khởi chạy ứng dụng ULN to DOCX Studio (Localhost Web App).
Tự động mở máy chủ Web FastAPI và giao diện người dùng trên trình duyệt mặc định.
"""

import os
import sys
import argparse

# Reconfigure standard streams to UTF-8 on Windows
if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Đảm bảo thư mục gốc nằm trong sys.path để các module hoạt động đồng bộ
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from server.app import start_server


def main():
    parser = argparse.ArgumentParser(description="Start ULN to DOCX Studio Web Server")
    parser.add_argument("--port", "-p", type=int, default=None, help="Port to run server (default: 8899 or next free port)")
    parser.add_argument("--host", type=str, default="127.0.0.1", help="Host address (default: 127.0.0.1)")
    parser.add_argument("--no-browser", action="store_true", help="Do not automatically open default web browser")
    args = parser.parse_args()

    start_server(port=args.port, host=args.host, open_browser=not args.no_browser)


if __name__ == "__main__":
    main()
