"""Lectura de Excel (xlsx, xlsm, xlsb, xls, ods) con calamine, todo como texto."""

from __future__ import annotations

import datetime as dt
import logging
import math
from decimal import Decimal
from pathlib import Path
from typing import Any

from python_calamine import CalamineError, CalamineWorkbook

from excel_tool.core.errors import ArchivoBloqueado, ArchivoIlegible, ArchivoNoEncontrado
from excel_tool.core.io import FileFormat, ReadOptions, ReadResult, SheetSummary
from excel_tool.core.io.header import (
    MAX_FILAS_ENCABEZADO,
    detect_header_row,
    normalize_column_names,
)
from excel_tool.core.progress import CancelToken, ProgressCallback, ProgressEvent

log = logging.getLogger(__name__)

_MAX_ENTERO_EXACTO = 1e15


def _float_a_texto(v: float) -> str:
    if math.isnan(v) or math.isinf(v):
        return ""
    if v == int(v) and abs(v) < _MAX_ENTERO_EXACTO:
        return str(int(v))
    texto = repr(v)
    if "e" in texto or "E" in texto:
        texto = format(Decimal(texto), "f")
        if "." in texto:
            texto = texto.rstrip("0").rstrip(".")
    return texto


def _fecha_a_texto(v: dt.date) -> str:
    return f"{v.day:02d}/{v.month:02d}/{v.year:04d}"


def _hora_a_texto(v: dt.time) -> str:
    return f"{v.hour:02d}:{v.minute:02d}:{v.second:02d}"


def _celda_a_texto(v: Any) -> str:
    """Convierte una celda de calamine a texto según el contrato (docs/contratos.md)."""
    if isinstance(v, str):
        return v
    if isinstance(v, bool):  # antes que int: bool es subclase de int
        return "VERDADERO" if v else "FALSO"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return _float_a_texto(v)
    if isinstance(v, dt.datetime):
        if v.hour == 0 and v.minute == 0 and v.second == 0 and v.microsecond == 0:
            return _fecha_a_texto(v)
        return f"{_fecha_a_texto(v)} {_hora_a_texto(v)}"
    if isinstance(v, dt.date):
        return _fecha_a_texto(v)
    if isinstance(v, dt.time):
        return _hora_a_texto(v)
    if isinstance(v, dt.timedelta):
        total = int(v.total_seconds())
        h, resto = divmod(abs(total), 3600)
        m, s = divmod(resto, 60)
        return f"{'-' if total < 0 else ''}{h:02d}:{m:02d}:{s:02d}"
    if v is None:
        return ""
    return str(v)


def _fila_a_texto(fila: list[Any]) -> list[str]:
    return [c if c.__class__ is str else _celda_a_texto(c) for c in fila]


def _fila_vacia(fila: list[Any]) -> bool:
    return all(c == "" or c is None for c in fila)


def _formato(path: Path) -> FileFormat:
    return path.suffix.lower().lstrip(".")  # type: ignore[return-value]


def _abrir(path: Path) -> CalamineWorkbook:
    try:
        return CalamineWorkbook.from_path(str(path))
    except FileNotFoundError as e:
        raise ArchivoNoEncontrado(
            "No se encontró el archivo. Fijate si se movió o se borró.", str(e)
        ) from e
    except PermissionError as e:
        raise ArchivoBloqueado("Cerrá el archivo en Excel y volvé a intentar.", str(e)) from e
    except CalamineError as e:
        if "assword" in f"{type(e).__name__} {e}":
            raise ArchivoIlegible(
                "El archivo tiene contraseña. Sacale la contraseña en Excel y probá de nuevo.",
                str(e),
            ) from e
        raise ArchivoIlegible("El archivo está dañado o no es un archivo de Excel.", str(e)) from e
    except OSError as e:
        raise ArchivoIlegible("No se pudo abrir el archivo.", str(e)) from e
    except Exception as e:  # calamine puede levantar otros tipos con archivos rotos
        raise ArchivoIlegible("El archivo está dañado o no es un archivo de Excel.", str(e)) from e


def _filas_crudas(wb: CalamineWorkbook, hoja: str) -> list[list[Any]]:
    """Todas las filas de la hoja desde A1, sin las filas vacías del final ni las
    columnas vacías del final."""
    try:
        filas = wb.get_sheet_by_name(hoja).to_python(skip_empty_area=False)
    except Exception as e:
        raise ArchivoIlegible("No se pudo leer la hoja del archivo.", f"{hoja}: {e}") from e
    while filas and _fila_vacia(filas[-1]):
        filas.pop()
    if not filas:
        return filas
    ancho = max(len(f) for f in filas)
    while ancho > 0 and all(
        len(f) < ancho or f[ancho - 1] == "" or f[ancho - 1] is None for f in filas
    ):
        ancho -= 1
    return [f[:ancho] for f in filas]


def _encabezado(filas: list[list[Any]], pedido: int | None) -> int:
    if pedido is not None:
        return pedido
    return detect_header_row([_fila_a_texto(f) for f in filas[:MAX_FILAS_ENCABEZADO]])


def read_excel(
    path: Path,
    options: ReadOptions | None = None,
    *,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> ReadResult:
    """Lee una hoja de un archivo Excel/ODS como texto. Lo usa `read_table`.

    calamine lee la hoja de una vez y no se puede cortar: se emite un evento indeterminado
    y se mira la cancelación antes y después de leer.
    """
    path = Path(path)
    options = options or ReadOptions()
    if cancel is not None:
        cancel.check()
    if progress is not None:
        progress(ProgressEvent("Leyendo la hoja…", 0, None))
    wb = _abrir(path)
    hojas = list(wb.sheet_names)
    if not hojas:
        raise ArchivoIlegible("El archivo no tiene hojas.", str(path))
    hoja = options.sheet if options.sheet is not None else hojas[0]
    if hoja not in hojas:
        raise ArchivoIlegible(
            f"No existe la hoja «{hoja}». Las hojas del archivo son: {', '.join(hojas)}.",
            str(path),
        )

    filas = _filas_crudas(wb, hoja)
    if cancel is not None:
        cancel.check()
    if not filas:
        raise ArchivoIlegible("La hoja está vacía.", f"{path}: {hoja}")
    fila_enc = _encabezado(filas, options.header_row)
    if fila_enc >= len(filas):
        raise ArchivoIlegible(
            "La hoja no tiene datos en la fila de encabezado indicada.",
            f"{path}: {hoja}, {len(filas)} filas",
        )

    ancho = max(len(f) for f in filas)
    datos = filas[fila_enc + 1 :]
    if options.nrows is not None:
        datos = datos[: options.nrows]
    nombres, renamed, avisos = normalize_column_names(
        (_fila_a_texto(filas[fila_enc]) + [""] * ancho)[:ancho]
    )
    import pandas as pd

    df = pd.DataFrame(
        [(_fila_a_texto(f) + [""] * ancho)[:ancho] for f in datos],
        columns=nombres,
        dtype=str,
    )
    return ReadResult(
        df=df,
        path=path,
        format=_formato(path),
        sheet=hoja,
        sheets=hojas,
        csv_options=None,
        header_row=fila_enc,
        renamed_columns=renamed,
        warnings=avisos,
    )


def list_excel_sheets(
    path: Path,
    *,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> list[SheetSummary]:
    """Nombre, filas de datos (sin el encabezado detectado) y columnas de cada hoja."""
    if cancel is not None:
        cancel.check()
    wb = _abrir(Path(path))
    resumen: list[SheetSummary] = []
    nombres = list(wb.sheet_names)
    for i, nombre in enumerate(nombres):
        if cancel is not None:
            cancel.check()
        if progress is not None:
            progress(ProgressEvent("Leyendo la hoja…", i, None))
        filas = _filas_crudas(wb, nombre)
        if not filas:
            resumen.append(SheetSummary(name=nombre, rows=0, columns=0))
            continue
        fila_enc = _encabezado(filas, None)
        resumen.append(
            SheetSummary(
                name=nombre,
                rows=max(len(filas) - fila_enc - 1, 0),
                columns=max(len(f) for f in filas),
            )
        )
    return resumen


__all__ = ["list_excel_sheets", "read_excel"]
