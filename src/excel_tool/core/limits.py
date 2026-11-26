"""Límites de Excel y Google Sheets y control de tamaño.

Contrato: docs/contratos.md § Fase 3 (ADR-032).

Los límites viven SOLO acá (valores en `limites_default.json`, se pisan con `limites.json`).
"""

from __future__ import annotations

import json
import logging
import math
import zlib
from dataclasses import dataclass
from importlib import resources
from pathlib import Path
from typing import Any, Literal

import numpy as np
import pandas as pd

from excel_tool import platform_utils
from excel_tool.core import explore
from excel_tool.core.io import WriteFormat, csv_bytes
from excel_tool.core.progress import CancelToken, ProgressCallback, ProgressEvent, ThrottledProgress

log = logging.getLogger(__name__)

Destination = Literal["excel", "google_sheets", "solo_guardar"]

_MB = 1024 * 1024
_NOMBRE_ARCHIVO = "limites.json"


@dataclass(frozen=True)
class Limits:
    excel_max_rows: int  # por hoja, CONTANDO el encabezado
    excel_max_columns: int
    excel_comfortable_rows: int
    excel_comfortable_mb: float
    gs_max_cells: int  # por planilla, contando el encabezado
    gs_max_import_mb: float
    gs_comfortable_cells: int
    gs_comfortable_mb: float


@dataclass(frozen=True)
class SizeEstimate:
    rows: int  # filas de datos (sin encabezado)
    columns: int
    cells: int  # (rows + 1) x columns
    approx_mb_xlsx: float
    approx_mb_csv: float


@dataclass(frozen=True)
class ColumnSuggestion:
    column: str
    parts: int  # valores distintos = cantidad de archivos
    largest_part_rows: int
    fits: bool  # la parte más grande entra en el límite cómodo del destino


@dataclass(frozen=True)
class SizeCheck:
    level: Literal["ok", "recomendado", "obligatorio"]
    reasons: list[str]
    suggested_parts: int
    suggested_rows_per_part: int
    suggested_columns: list[ColumnSuggestion]
    # Máximo de filas que puede tener una parte en ese destino (F5-A). Al dividir sin cortar un
    # grupo hay que usar este tope, no `suggested_rows_per_part` (que no deja margen).
    # None si no hay tope (CSV en "solo guardar").
    max_rows_per_part: int | None = None


# ---------------------------------------------------------------------------- carga

# (sección del json, clave del json, campo de Limits, tipo)
_CAMPOS: tuple[tuple[str, str, str, type], ...] = (
    ("excel", "max_rows", "excel_max_rows", int),
    ("excel", "max_columns", "excel_max_columns", int),
    ("excel", "comfortable_rows", "excel_comfortable_rows", int),
    ("excel", "comfortable_mb", "excel_comfortable_mb", float),
    ("google_sheets", "max_cells", "gs_max_cells", int),
    ("google_sheets", "max_import_mb", "gs_max_import_mb", float),
    ("google_sheets", "comfortable_cells", "gs_comfortable_cells", int),
    ("google_sheets", "comfortable_mb", "gs_comfortable_mb", float),
)


def _valor_valido(valor: Any, tipo: type) -> bool:
    if isinstance(valor, bool) or not isinstance(valor, int | float):
        return False
    if not math.isfinite(valor) or valor <= 0:
        return False
    return tipo is float or float(valor).is_integer()


def _leer_json(ruta: Path) -> dict | None:
    try:
        datos = json.loads(ruta.read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        return None
    except (OSError, ValueError) as exc:  # ValueError incluye JSON y Unicode dañados
        log.warning("No se pudo leer %s (se ignora): %s", ruta, exc)
        return None
    if not isinstance(datos, dict):
        log.warning("%s no tiene el formato esperado (se ignora)", ruta)
        return None
    return datos


def _aplicar(valores: dict[str, Any], datos: dict, origen: str) -> None:
    for seccion, clave, campo, tipo in _CAMPOS:
        bloque = datos.get(seccion)
        if not isinstance(bloque, dict) or clave not in bloque:
            continue
        valor = bloque[clave]
        if _valor_valido(valor, tipo):
            valores[campo] = int(valor) if tipo is int else float(valor)
        else:
            log.warning(
                "%s: valor inválido en %s.%s (%r); se usa el anterior",
                origen,
                seccion,
                clave,
                valor,
            )


def load_limits() -> Limits:
    """Default, pisado por `limites.json` junto al ejecutable y por el de la carpeta de datos."""
    ruta_default = resources.files("excel_tool") / "limites_default.json"
    valores: dict[str, Any] = {}
    _aplicar(valores, json.loads(ruta_default.read_text(encoding="utf-8")), "limites_default.json")
    faltan = [c for _, _, c, _ in _CAMPOS if c not in valores]
    if faltan:  # error de empaquetado, no del usuario
        raise ValueError(f"limites_default.json incompleto: faltan {faltan}")

    carpetas: list[Path] = []
    junto_exe = platform_utils.app_dir()
    if junto_exe is not None:
        carpetas.append(junto_exe)
    try:
        carpetas.append(platform_utils.user_data_dir())
    except OSError as exc:
        log.warning("No se pudo abrir la carpeta de datos del usuario: %s", exc)
    for carpeta in carpetas:
        ruta = carpeta / _NOMBRE_ARCHIVO
        datos = _leer_json(ruta)
        if datos is not None:
            _aplicar(valores, datos, str(ruta))

    # Coherencia: lo cómodo nunca puede superar lo máximo.
    tope_filas = max(valores["excel_max_rows"] - 1, 1)
    if valores["excel_comfortable_rows"] > tope_filas:
        log.warning("excel.comfortable_rows supera el máximo; se ajusta a %d", tope_filas)
        valores["excel_comfortable_rows"] = tope_filas
    if valores["gs_comfortable_cells"] > valores["gs_max_cells"]:
        log.warning("google_sheets.comfortable_cells supera el máximo; se ajusta")
        valores["gs_comfortable_cells"] = valores["gs_max_cells"]
    if valores["gs_comfortable_mb"] > valores["gs_max_import_mb"]:
        log.warning("google_sheets.comfortable_mb supera el máximo; se ajusta")
        valores["gs_comfortable_mb"] = valores["gs_max_import_mb"]
    return Limits(**valores)


# ---------------------------------------------------------------------------- estimación

# Calibración del tamaño del .xlsx (F3-01). Se escribieron con `write_table` archivos reales
# (rendimiento_2M.csv, clientes_grande.csv, columnas de números al azar y de 3 categorías) y se
# comparó bytes por celda del .xlsx con los bytes por celda del texto de la muestra comprimido con
# zlib. Ajuste lineal por mínimos cuadrados:
# xlsx_por_celda = _XLSX_BASE + _XLSX_FACTOR * zlib_por_celda.
# El .xlsx es XML comprimido: cada celda cuesta unos 3 bytes fijos (referencia, etiqueta) y el
# contenido comprime parecido al del texto plano.
_XLSX_BASE = 2.92
_XLSX_FACTOR = 1.31


def _muestra(df: pd.DataFrame, filas: int) -> pd.DataFrame:
    """Hasta `filas` filas repartidas parejo por toda la tabla (no copia la tabla entera)."""
    if len(df) <= filas:
        return df
    posiciones = np.linspace(0, len(df) - 1, filas).astype(np.int64)
    return df.iloc[posiciones]


def estimate(df: pd.DataFrame, *, sample_rows: int = 10_000) -> SizeEstimate:
    """Tamaño aproximado del resultado, a partir de una muestra de filas."""
    filas, columnas = len(df), len(df.columns)
    celdas = (filas + 1) * columnas
    if filas == 0 or columnas == 0:
        return SizeEstimate(filas, columnas, celdas, 0.0, 0.0)
    muestra = _muestra(df, max(sample_rows, 1))
    # Como el CSV para Excel (`;`, CRLF, comillas cuando hacen falta), medido en memoria.
    crudo = csv_bytes(muestra)
    celdas_muestra = muestra.size
    csv_por_celda = len(crudo) / celdas_muestra
    zlib_por_celda = len(zlib.compress(crudo, 6)) / celdas_muestra
    xlsx_por_celda = _XLSX_BASE + _XLSX_FACTOR * zlib_por_celda
    return SizeEstimate(
        rows=filas,
        columns=columnas,
        cells=celdas,
        approx_mb_xlsx=xlsx_por_celda * celdas / _MB,
        approx_mb_csv=csv_por_celda * celdas / _MB,
    )


# ---------------------------------------------------------------------------- control


def _n(valor: int) -> str:
    """Entero con punto de miles: 1.048.575."""
    return f"{valor:,}".replace(",", ".")


def _mb(valor: float) -> str:
    """MB con coma decimal y punto de miles: 1.234,5."""
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _partes(total: float, tope: float) -> int:
    return max(1, math.ceil(total / tope)) if tope > 0 else 1


def _mb_del_formato(est: SizeEstimate, format: WriteFormat) -> float:
    return est.approx_mb_xlsx if format == "xlsx" else est.approx_mb_csv


def check(
    est: SizeEstimate,
    destination: Destination,
    format: WriteFormat,
    *,
    limits: Limits | None = None,
    suggestions: list[ColumnSuggestion] | None = None,
) -> SizeCheck:
    """Semáforo de tamaño para el destino elegido, con motivos y partes sugeridas."""
    lim = limits if limits is not None else load_limits()
    duros: list[str] = []
    blandos: list[str] = []
    partes = 1
    mb = _mb_del_formato(est, format)
    max_datos = lim.excel_max_rows - 1
    tope_parte: int | None

    if destination == "google_sheets":
        tope_parte = max(lim.gs_comfortable_cells // max(est.columns, 1) - 1, 1)
        if est.cells > lim.gs_max_cells:
            duros.append(
                f"Google Sheets acepta hasta {_n(lim.gs_max_cells)} celdas por planilla; "
                f"el resultado tiene {_n(est.cells)}."
            )
        if mb > lim.gs_max_import_mb:
            duros.append(
                f"Google Sheets no importa archivos de más de {_mb(lim.gs_max_import_mb)} MB; "
                f"el resultado pesaría unos {_mb(mb)} MB."
            )
        if est.cells > lim.gs_comfortable_cells:
            blandos.append(
                "Para trabajar cómodo en Google Sheets conviene no pasar de "
                f"{_n(lim.gs_comfortable_cells)} celdas; el resultado tiene {_n(est.cells)}."
            )
            partes = max(partes, _partes(est.cells, lim.gs_comfortable_cells))
        if mb > lim.gs_comfortable_mb:
            blandos.append(
                f"Para importar cómodo a Google Sheets conviene que cada archivo pese menos de "
                f"{_mb(lim.gs_comfortable_mb)} MB; el resultado pesaría unos {_mb(mb)} MB."
            )
            partes = max(partes, _partes(mb, lim.gs_comfortable_mb))
    else:
        # Excel, y "solo guardar" con xlsx: el archivo no puede pasar el límite de la hoja.
        if destination == "solo_guardar" and format != "xlsx":
            return SizeCheck("ok", [], 1, max(est.rows, 1), list(suggestions or []), None)
        tope_parte = lim.excel_comfortable_rows if destination == "excel" else max_datos
        if est.rows + 1 > lim.excel_max_rows:
            duros.append(
                f"Excel acepta hasta {_n(max_datos)} filas de datos por hoja (más el encabezado); "
                f"el resultado tiene {_n(est.rows)}."
            )
            partes = max(partes, _partes(est.rows, max_datos))
        if est.columns > lim.excel_max_columns:
            duros.append(
                f"Excel acepta hasta {_n(lim.excel_max_columns)} columnas por hoja; "
                f"el resultado tiene {_n(est.columns)}."
            )
        if destination == "excel":
            if est.rows > lim.excel_comfortable_rows:
                blandos.append(
                    "Para trabajar cómodo en Excel conviene no pasar de "
                    f"{_n(lim.excel_comfortable_rows)} filas; el resultado tiene {_n(est.rows)}."
                )
                partes = max(partes, _partes(est.rows, lim.excel_comfortable_rows))
            if mb > lim.excel_comfortable_mb:
                blandos.append(
                    f"Para abrirlo cómodo en Excel conviene que cada archivo pese menos de "
                    f"{_mb(lim.excel_comfortable_mb)} MB; el resultado pesaría unos {_mb(mb)} MB."
                )
                partes = max(partes, _partes(mb, lim.excel_comfortable_mb))

    if duros:
        nivel: Literal["ok", "recomendado", "obligatorio"] = "obligatorio"
    elif blandos:
        nivel = "recomendado"
    else:
        nivel = "ok"
    if nivel == "ok":
        partes = 1
    filas_por_parte = max(math.ceil(est.rows / partes), 1)
    return SizeCheck(
        nivel, duros + blandos, partes, filas_por_parte, list(suggestions or []), tope_parte
    )


# ---------------------------------------------------------------------------- columnas para dividir

_FILAS_PRUEBA = 20_000


def _limite_filas_parte(
    est: SizeEstimate, destination: Destination, format: WriteFormat, lim: Limits
) -> int:
    """Filas que puede tener como mucho una parte para quedar dentro de lo cómodo."""
    mb_por_fila = _mb_del_formato(est, format) / max(est.rows, 1)
    if destination == "google_sheets":
        por_celdas = lim.gs_comfortable_cells // max(est.columns, 1) - 1
        mb_comodo = lim.gs_comfortable_mb
    elif destination == "excel":
        por_celdas = lim.excel_comfortable_rows
        mb_comodo = lim.excel_comfortable_mb
    elif format == "xlsx":  # solo guardar: solo cuenta el límite duro de la hoja
        por_celdas = lim.excel_max_rows - 1
        mb_comodo = math.inf
    else:
        return max(est.rows, 1)
    por_mb = mb_comodo / mb_por_fila if mb_por_fila > 0 else math.inf
    return int(max(min(por_celdas, por_mb), 1))


def suggest_split_columns(
    df: pd.DataFrame,
    destination: Destination,
    format: WriteFormat,
    *,
    limits: Limits | None = None,
    max_parts: int = 50,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> list[ColumnSuggestion]:
    """Columnas con entre 2 y `max_parts` valores distintos que sirven para dividir por valor.

    Primero descarta rápido las columnas con muchos valores mirando solo una muestra del principio
    (si la muestra ya tiene más de `max_parts` valores, la columna entera también), y recién después
    cuenta con `explore.value_counts` las que quedan.
    """
    lim = limits if limits is not None else load_limits()
    reporte = ThrottledProgress(progress)
    columnas = list(df.columns)
    total = len(columnas)
    if len(df) == 0 or total == 0:
        return []
    est = estimate(df)
    tope_parte = _limite_filas_parte(est, destination, format, lim)
    prueba = df.iloc[:_FILAS_PRUEBA] if len(df) > _FILAS_PRUEBA else df
    resultado: list[ColumnSuggestion] = []
    for i, nombre in enumerate(columnas):
        if cancel is not None:
            cancel.check()
        reporte(ProgressEvent(f"Mirando la columna {nombre}…", i, total))
        if prueba[nombre].nunique(dropna=False) > max_parts:
            continue
        cuentas = explore.value_counts(
            df,
            [nombre],
            explore.DistinctOptions(limit=max_parts + 1),
            cancel=cancel,
        )
        distintos = cuentas.distinct_total
        if distintos < 2 or distintos > max_parts:
            continue
        mayor = int(cuentas.table[cuentas.count_column].max())
        resultado.append(ColumnSuggestion(str(nombre), distintos, mayor, mayor <= tope_parte))
    reporte.force(ProgressEvent("Listo", total, total))
    resultado.sort(key=lambda s: (not s.fits, s.parts, s.largest_part_rows))
    return resultado
