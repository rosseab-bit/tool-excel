"""Errores del core hacia la UI. El mensaje va en castellano, listo para mostrar."""

from __future__ import annotations


class ExcelToolError(Exception):
    """Base. `mensaje` está en castellano, sin jerga, listo para mostrar al usuario."""

    def __init__(self, mensaje: str, detalle: str | None = None) -> None:
        super().__init__(mensaje)
        self.mensaje = mensaje
        self.detalle = detalle


class ArchivoNoEncontrado(ExcelToolError):
    """El archivo no existe (o se movió)."""


class ArchivoBloqueado(ExcelToolError):
    """PermissionError: "Cerrá el archivo en Excel y volvé a intentar."."""


class FormatoNoSoportado(ExcelToolError):
    """Extensión fuera de lo soportado (PLAN §1.1)."""


class ArchivoIlegible(ExcelToolError):
    """Dañado, vacío, encoding imposible, hoja inexistente."""


class Cancelado(ExcelToolError):
    """El usuario canceló una operación larga."""

    def __init__(self, mensaje: str = "Operación cancelada.", detalle: str | None = None) -> None:
        super().__init__(mensaje, detalle)
