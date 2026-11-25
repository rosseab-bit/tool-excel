"""Punto de entrada de PyInstaller (no puede arrancar con `python -m`)."""

import multiprocessing
import sys

from excel_tool import app_entry

if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(app_entry.main())
