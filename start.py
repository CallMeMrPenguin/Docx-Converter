"""
start.py - Trình khởi chạy ứng dụng ULN to DOCX Studio (Localhost Web App).
Tự động mở máy chủ Web FastAPI và giao diện người dùng trên trình duyệt mặc định.
"""

import os
import sys

# Đảm bảo thư mục gốc nằm trong sys.path để các module hoạt động đồng bộ
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from server.app import start_server


if __name__ == "__main__":
    start_server()
