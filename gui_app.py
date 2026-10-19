"""
gui_app.py - Main Desktop Entrypoint & Backward-Compatible Facade.
Launches the Flat Minimal Light UI ULN Formatter Desktop Application.
The implementation lives under gui/ (gui/app.py, gui/theme.py, gui/components/, gui/dialogs/).
"""

import sys
import tkinter as tk
from gui.app import ULNFormatterApp, main

__all__ = ["ULNFormatterApp", "main"]

if __name__ == "__main__":
    main()
