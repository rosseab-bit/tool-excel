"""Punto de entrada: app, --version y --self-test."""

import importlib
import json
import logging
import sys
import tempfile
from importlib import resources
from logging.handlers import RotatingFileHandler
from pathlib import Path

from excel_tool import APP_NAME, __version__, platform_utils

logger = logging.getLogger("excel_tool")
_HANDLER_ATTR = "_excel_tool_handler"


def _setup_logging() -> None:
    """Log a archivo. Si no se puede, sigue sin log."""
    try:
        if any(getattr(h, _HANDLER_ATTR, False) for h in logger.handlers):
            return
        handler = RotatingFileHandler(
            platform_utils.log_dir() / "excel_tool.log",
            maxBytes=1_000_000,
            backupCount=3,
            encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        setattr(handler, _HANDLER_ATTR, True)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
    except OSError:
        pass


def _print(texto: str, *, error: bool = False) -> None:
    stream = sys.stderr if error else sys.stdout
    if stream is not None:
        print(texto, file=stream)


def _paso_imports() -> str:
    for nombre in (
        "pandas",
        "numpy",
        "openpyxl",
        "python_calamine",
        "tkinter",
        "excel_tool.ui.app",
    ):
        importlib.import_module(nombre)
    return "librerías importadas"


def _paso_limites() -> str:
    ruta = resources.files("excel_tool") / "limites_default.json"
    datos = json.loads(ruta.read_text(encoding="utf-8"))
    if "excel" not in datos or "google_sheets" not in datos:
        raise ValueError("faltan secciones en limites_default.json")
    return "limites_default.json leído"


def _paso_archivos() -> str:
    import pandas as pd

    from excel_tool.core.io import ReadOptions, WriteOptions, read_table, write_table

    original = pd.DataFrame({"nombre": ["Ñandú", "acción"], "codigo": ["00123", "00124"]})
    with tempfile.TemporaryDirectory() as tmp:
        carpeta = Path(tmp)
        for formato, nombre in (("csv_excel", "prueba.csv"), ("xlsx", "prueba.xlsx")):
            ruta = carpeta / nombre
            write_table(original, ruta, WriteOptions(format=formato))
            leido = read_table(ruta, ReadOptions(header_row=0)).df
            if (
                list(leido.columns) != list(original.columns)
                or leido.values.tolist() != original.values.tolist()
            ):
                raise ValueError(f"{nombre} no volvió igual")
    return "CSV y xlsx escritos y releídos"


def _paso_ventana() -> str:
    import tkinter as tk

    raiz = tk.Tk()
    try:
        raiz.update()
    finally:
        raiz.destroy()
    return "ventana creada"


def _self_test() -> int:
    pasos = [
        ("imports", _paso_imports),
        ("limites", _paso_limites),
        ("archivos", _paso_archivos),
        ("ventana", _paso_ventana),
    ]
    fallas = 0
    for nombre, funcion in pasos:
        try:
            detalle = funcion()
            logger.info("SELF-TEST OK %s: %s", nombre, detalle)
        except Exception as exc:
            fallas += 1
            logger.error("SELF-TEST FALLA %s: %s", nombre, exc, exc_info=True)
    resumen = "SELF-TEST OK" if fallas == 0 else f"SELF-TEST FALLA ({fallas} pasos)"
    _print(resumen)
    return 0 if fallas == 0 else 1


def _run_app() -> int:
    try:
        from excel_tool.ui.app import run

        run()
        return 0
    except Exception:
        logger.exception("Error no manejado")
        try:
            from tkinter import messagebox

            messagebox.showerror(
                APP_NAME, "Ocurrió un problema inesperado y la aplicación se va a cerrar."
            )
        except Exception:
            logger.exception("No se pudo mostrar el mensaje de error")
        return 1


def main(argv: list[str] | None = None) -> int:
    _setup_logging()
    args = sys.argv[1:] if argv is None else argv
    if not args:
        return _run_app()
    if args == ["--version"]:
        logger.info("%s %s", APP_NAME, __version__)
        _print(f"{APP_NAME} {__version__}")
        return 0
    if args == ["--self-test"]:
        return _self_test()
    logger.warning("Argumento desconocido: %s", args)
    _print("Uso: excel-tool [--version | --self-test]", error=True)
    return 2
