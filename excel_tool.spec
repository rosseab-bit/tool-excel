# -*- mode: python ; coding: utf-8 -*-
# Spec unico para Windows y Linux. Los scripts build_* lo llaman con pyinstaller.
# Modo carpeta por defecto; onefile si EXCEL_TOOL_ONEFILE=1.
# Este archivo no es codigo de la app: es la unica excepcion que consulta sys.platform
# (solo para elegir el icono: .ico en Windows, .png en Linux).
import os
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_all

raiz = Path(SPECPATH)
paquete = raiz / "src" / "excel_tool"
onefile = os.environ.get("EXCEL_TOOL_ONEFILE") == "1"
icono = str(paquete / "assets" / ("icono.ico" if sys.platform == "win32" else "icono.png"))

# python_calamine trae una extension compilada.
datas, binaries, hiddenimports = collect_all("python_calamine")
# Archivos propios, para importlib.resources.files("excel_tool").
datas += [
    (str(paquete / "limites_default.json"), "excel_tool"),
    (str(paquete / "assets"), "excel_tool/assets"),
]
hiddenimports += ["pandas", "openpyxl", "tkinter"]

excludes = [
    "matplotlib", "IPython", "scipy", "pytest", "PyQt5", "PyQt6", "PySide2", "PySide6",
    "tkinter.test", "test", "unittest", "sphinx", "jupyter", "notebook", "numba",
    "sqlalchemy", "lxml", "bs4", "pyarrow", "tables", "xlsxwriter", "xlwt", "setuptools",
]

a = Analysis(
    [str(raiz / "packaging" / "launcher.py")],
    pathex=[str(raiz / "src")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
)
pyz = PYZ(a.pure)

if onefile:
    exe = EXE(
        pyz, a.scripts, a.binaries, a.datas, [],
        name="ExcelJoinTool", console=False, upx=False, icon=icono,
    )
else:
    exe = EXE(
        pyz, a.scripts, [],
        exclude_binaries=True, name="ExcelJoinTool", console=False, upx=False, icon=icono,
    )
    coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="ExcelJoinTool")
