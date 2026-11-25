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
hiddenimports += ["pandas", "openpyxl", "tkinter", "pyarrow"]

excludes = [
    "matplotlib", "IPython", "scipy", "pytest", "PyQt5", "PyQt6", "PySide2", "PySide6",
    "tkinter.test", "test", "unittest", "sphinx", "jupyter", "notebook", "numba",
    "sqlalchemy", "lxml", "bs4",
    # pyarrow: solo lo que se probo que no hace falta (el resto lo importa pyarrow/__init__)
    "pyarrow.flight", "pyarrow.orc", "pyarrow.cuda", "pyarrow.tests", "pyarrow.include",
    "tables", "xlsxwriter", "xlwt", "setuptools",
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
# pyarrow se usa solo para textos compactos, pero todas sus extensiones enlazan libarrow_python,
# que a su vez necesita libarrow, compute, dataset, acero, substrait y parquet: esas SE QUEDAN.
# Se sacan solo las partes independientes (nadie las enlaza): Flight, S3, ORC, Gandiva, CUDA.
# Se compara por prefijo del nombre sin "lib" y sin extension de plataforma, asi sirve igual
# para Linux (libarrow_flight.so.2600), y para Windows (arrow_flight.dll, _flight.cp312-win_amd64.pyd).
_sin = (
    "arrow_flight", "arrow_python_flight", "_flight",
    "arrow_s3", "_s3fs",
    "arrow_gandiva", "gandiva",
    "arrow_cuda", "_cuda",
    "_orc", "_dataset_orc",
)
def _sobra(nombre):
    partes = Path(nombre).parts
    if "pyarrow" not in partes:
        return False
    base = Path(nombre).name.lower()
    if base.startswith("lib"):
        base = base[3:]
    return base.startswith(_sin)
a.binaries = [b for b in a.binaries if not _sobra(b[0])]
# Fuentes de Cython y cabeceras de C++ (solo sirven para compilar extensiones, ~7 MB) y tests.
def _sobra_dato(nombre):
    p = Path(nombre)
    if "pyarrow" not in p.parts:
        return False
    return (
        _sobra(nombre)
        or p.suffix.lower() in (".pyx", ".pxd", ".pxi", ".h", ".hpp")
        or "include" in p.parts
        or "tests" in p.parts
    )
a.datas = [d for d in a.datas if not _sobra_dato(d[0])]
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
