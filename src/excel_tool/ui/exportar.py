"""Puentes hacia las funciones de exportar del core (import perezoso, reemplazables en tests)."""

from __future__ import annotations

import importlib
import logging
from collections.abc import Callable
from pathlib import Path
from typing import Any

from excel_tool.core.errors import ExcelToolError

log = logging.getLogger(__name__)

TEXTO_NO_DISPONIBLE = "Esta opción todavía no está disponible en esta versión."
MODULOS = ("excel_tool.core.profile", "excel_tool.core.explore", "excel_tool.core.export")


def _buscar(nombre: str) -> Callable[..., Any]:
    for modulo in MODULOS:
        try:
            funcion = getattr(importlib.import_module(modulo), nombre, None)
        except ImportError:
            continue
        if funcion is not None:
            return funcion
    log.warning("%s todavía no existe en el core.", nombre)
    raise ExcelToolError(TEXTO_NO_DISPONIBLE)


def guardar_informe(perfil: Any, distintos: list[Any], ruta: Path) -> Any:
    """Llama a ``export_profile`` del core."""
    return _buscar("export_profile")(perfil, distintos, ruta)


def guardar_distintos(resultado: Any, ruta: Path, formato: str) -> Any:
    """Llama a ``export_distincts`` del core."""
    return _buscar("export_distincts")(resultado, ruta, formato)
