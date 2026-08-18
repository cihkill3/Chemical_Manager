"""Runtime resource lookup for source and PyInstaller one-file execution."""

import os
import sys


def resource_path(filename: str) -> str:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    local = os.path.join(app_root, filename)
    if os.path.exists(local):
        return local
    return os.path.join(os.path.dirname(app_root), filename)
