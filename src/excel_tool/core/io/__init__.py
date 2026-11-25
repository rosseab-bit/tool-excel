"""Lectura y escritura de archivos (CSV y Excel). Contrato: docs/contratos.md § Fase 1.

`read_table`, `list_sheets`, `write_table` y `write_sheets` se completan en F1-02/F1-03.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Literal

import pandas as pd

if TYPE_CHECKING:
    from excel_tool.core.types import ColumnType

FileFormat = Literal["xlsx", "xlsm", "xlsb", "xls", "ods", "csv"]  # .txt se trata como "csv"
WriteFormat = Literal["xlsx", "csv_excel", "csv_google", "csv_standard"]


@dataclass(frozen=True)
class CsvOptions:
    separator: str  # ";", ",", "\t" o "|"
    encoding: str  # "utf-8-sig", "utf-8" o "cp1252"


@dataclass(frozen=True)
class ReadOptions:
    sheet: str | None = None  # Excel: None = primera hoja. CSV: se ignora
    header_row: int | None = None  # None = detectar. Las filas de arriba se descartan
    csv: CsvOptions | None = None  # None = detectar
    nrows: int | None = None  # None = todas (para preview: 10 o 50)


@dataclass(frozen=True)
class SheetSummary:
    name: str
    rows: int  # filas de datos sin contar el encabezado detectado
    columns: int


@dataclass
class ReadResult:
    df: pd.DataFrame  # todas las columnas dtype object con str; "" = vacío
    path: Path
    format: FileFormat
    sheet: str | None  # hoja leída (Excel)
    sheets: list[str]  # todas las hojas (Excel); [] en CSV
    csv_options: CsvOptions | None  # las usadas (detectadas o pedidas)
    header_row: int  # fila del encabezado usada (0 = primera fila)
    renamed_columns: dict[str, str]  # nombre final -> original, solo las renombradas
    warnings: list[str]  # en castellano, para mostrar tal cual


@dataclass(frozen=True)
class WriteOptions:
    format: WriteFormat
    sheet_name: str = "Datos"  # xlsx
    column_types: dict[str, ColumnType] | None = None
    # None -> todo como texto, tal cual. dict -> ver docs/contratos.md (reglas de escritura).


@dataclass
class WriteResult:
    path: Path
    rows: int
    unconverted: dict[str, list[tuple[int, str]]]  # columna -> [(fila, valor)] quedaron texto


from excel_tool.core.errors import (  # noqa: E402
    ArchivoNoEncontrado,
    FormatoNoSoportado,
)
from excel_tool.core.io.csv import detect_csv_options, read_csv  # noqa: E402
from excel_tool.core.io.excel import list_excel_sheets, read_excel  # noqa: E402
from excel_tool.core.io.writer import write_sheets_impl, write_table_impl  # noqa: E402

_EXT_EXCEL = {".xlsx", ".xlsm", ".xlsb", ".xls", ".ods"}
_EXT_CSV = {".csv", ".txt"}


def _verificar_existe(path: Path) -> None:
    if not path.exists():
        raise ArchivoNoEncontrado(
            "No se encontró el archivo. Fijate si se movió o se borró.", str(path)
        )


def _extension(path: Path) -> str:
    ext = path.suffix.lower()
    if ext not in _EXT_EXCEL and ext not in _EXT_CSV:
        raise FormatoNoSoportado(
            "Este tipo de archivo no se puede abrir. "
            "Usá un archivo de Excel (.xlsx, .xls, .xlsm, .xlsb, .ods) o de texto (.csv, .txt).",
            str(path),
        )
    return ext


def read_table(path: Path, options: ReadOptions | None = None) -> ReadResult:
    """Lee un archivo como texto, según su extensión."""
    path = Path(path)
    ext = _extension(path)
    _verificar_existe(path)
    if ext in _EXT_CSV:
        return read_csv(path, options)
    return read_excel(path, options)


def list_sheets(path: Path) -> list[SheetSummary]:
    """Hojas de un Excel con filas y columnas. CSV: una sola entrada con el nombre del archivo."""
    path = Path(path)
    ext = _extension(path)
    _verificar_existe(path)
    if ext in _EXT_EXCEL:
        return list_excel_sheets(path)
    resultado = read_csv(path)
    return [SheetSummary(name=path.name, rows=len(resultado.df), columns=len(resultado.df.columns))]


def write_table(df: pd.DataFrame, path: Path, options: WriteOptions) -> WriteResult:
    """Guarda una tabla como CSV (3 estilos) o xlsx. Sin límites (eso es de save_result)."""
    return write_table_impl(df, path, options)


def write_sheets(
    sheets: dict[str, pd.DataFrame],
    path: Path,
    column_types: dict[str, dict[str, ColumnType]] | None = None,
) -> list[WriteResult]:
    """Guarda varias hojas en un solo .xlsx."""
    return write_sheets_impl(sheets, path, column_types)


__all__ = [
    "CsvOptions",
    "FileFormat",
    "ReadOptions",
    "ReadResult",
    "SheetSummary",
    "WriteFormat",
    "WriteOptions",
    "WriteResult",
    "detect_csv_options",
    "list_sheets",
    "read_table",
    "write_sheets",
    "write_table",
]
