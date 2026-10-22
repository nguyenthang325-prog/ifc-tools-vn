"""Cấu hình chung: nơi đọc đường dẫn tương đối và nơi ghi kết quả.

- Đường dẫn IFC tương đối được tính từ thư mục hiện tại.
- Kết quả ghi vào biến môi trường IFC_TOOLS_OUTPUT nếu có,
  mặc định ./ifc_tools_output/ — mỗi ngày 1 thư mục con YYYY-MM-DD.
"""
import os
from datetime import date
from pathlib import Path

BASE        = Path.cwd()
KET_QUA_DIR = Path(os.environ.get("IFC_TOOLS_OUTPUT") or BASE / "ifc_tools_output")
OUTPUT_DIR  = KET_QUA_DIR / date.today().strftime("%Y-%m-%d")
