"""Escritura de CSV y xlsx. Contrato: docs/contratos.md § Fase 1 (reglas de escritura).

Sin límites de tamaño ni división: eso es de `output.save_result` (Fase 3).
Los archivos se escriben en un temporal y se renombran al final (cancelar no deja restos).
"""

from __future__ import annotations

import contextlib
import io
import logging
import os
import re
import tempfile
from collections.abc import Callable, Iterable, Mapping
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any

import numpy as np
import pandas as pd
import xlsxwriter
import xlsxwriter.exceptions

from excel_tool.core.errors import ArchivoBloqueado, ExcelToolError
from excel_tool.core.progress import CancelToken, ProgressCallback, ProgressEvent, ThrottledProgress
from excel_tool.core.types import to_date, to_number

if TYPE_CHECKING:
    from excel_tool.core.io import WriteOptions, WriteResult
    from excel_tool.core.types import ColumnType

logger = logging.getLogger(__name__)

# formato -> (separador, encoding, decimal)
_CSV_FORMATOS = {
    "csv_excel": (";", "utf-8-sig", ","),
    "csv_google": (",", "utf-8", "."),
    "csv_standard": (",", "utf-8-sig", "."),
}
_NUMERICOS = {"entero", "decimal"}
_FECHAS = {"fecha", "fecha_hora"}
_FMT_FECHA = "dd/mm/yyyy"
_FMT_FECHA_HORA = "dd/mm/yyyy hh:mm:ss"
_HOJA_INVALIDA = re.compile(r"[\[\]:*?/\\]")
_MAX_HOJA = 31
_TANDA_CSV = 50_000  # filas por tanda al escribir CSV (cancelar / avance)
_MAX_CELDA = 32_767  # caracteres por celda en Excel
_ANIO_MINIMO = 1900  # Excel no muestra fechas anteriores
_CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")  # lo que Excel no admite
_CADA_FILAS = 5_000  # xlsx: cada cuántas filas se mira cancelar y se informa


def nombre_de_hoja(nombre: str, usados: set[str] | None = None) -> str:
    """Nombre válido para Excel: sin `[]:*?/\\`, máx. 31 caracteres, sin repetir (sin importar
    mayúsculas). Si se ajusta, lo avisa en el log. Registra el resultado en `usados`."""
    limpio = _HOJA_INVALIDA.sub("_", str(nombre)).strip().strip("'").strip()
    limpio = limpio[:_MAX_HOJA].strip().strip("'") or "Datos"
    final = limpio
    if usados is not None:
        n = 2
        while final.lower() in usados:
            sufijo = f" ({n})"
            final = limpio[: _MAX_HOJA - len(sufijo)] + sufijo
            n += 1
        usados.add(final.lower())
    if final != nombre:
        logger.warning("Nombre de hoja ajustado para Excel: %r -> %r", nombre, final)
    return final


def _columnas_convertidas(
    df: pd.DataFrame, column_types: dict[str, ColumnType] | None
) -> tuple[list[pd.Series], list[str], dict[str, list[tuple[int, str]]]]:
    """Por columna: serie lista para escribir y su clase ("texto", "entero", "decimal",
    "fecha", "fecha_hora"). Las no convertibles quedan como texto y se informan."""
    series: list[pd.Series] = []
    clases: list[str] = []
    fallidas: dict[str, list[tuple[int, str]]] = {}
    for i, nombre in enumerate(df.columns):
        col = df.iloc[:, i]
        tipo = (column_types or {}).get(str(nombre), "texto")
        if tipo in _NUMERICOS or tipo in _FECHAS:
            conv = to_number(col) if tipo in _NUMERICOS else to_date(col)
            if conv.failed:
                fallidas[str(nombre)] = conv.failed
            valores = conv.values.copy()
            if conv.failed:
                # lo no convertible se escribe tal cual: la columna pasa a ser mixta (objeto)
                valores = valores.astype(object)
                for pos, original in conv.failed:
                    valores.iloc[pos] = original
            series.append(valores)
            clases.append(tipo)
        else:
            series.append(col.fillna("").astype(str))
            clases.append("texto")
    return series, clases, fallidas


def _texto_decimal(serie: pd.Series, decimal: str, entero: bool) -> pd.Series:
    """Número -> texto sin exponente, con el decimal pedido. Vacío donde no hay valor."""
    numeros = pd.to_numeric(serie, errors="coerce")
    nulo = numeros.isna()
    if entero:
        texto = numeros.round().astype("Int64").astype(str)
    else:
        texto = numeros.astype(str)
        con_exp = texto.str.contains("e", regex=False) & ~nulo
        if con_exp.any():
            texto[con_exp] = numeros[con_exp].map(lambda x: np.format_float_positional(x, trim="-"))
        texto = texto.str.replace(r"\.0$", "", regex=True)
        if decimal != ".":
            texto = texto.str.replace(".", decimal, regex=False)
    return texto.where(~nulo, "")


def _a_texto_csv(serie: pd.Series, clase: str, decimal: str) -> pd.Series:
    """Serie convertida -> texto para el CSV. Los valores no convertidos (str) se dejan tal cual."""
    if clase == "texto":
        return serie
    es_texto = serie.map(lambda v: isinstance(v, str)) if serie.dtype == object else None
    if clase in _NUMERICOS:
        base = serie.where(~es_texto, np.nan) if es_texto is not None else serie
        salida = _texto_decimal(base, decimal, clase == "entero")
    else:
        base = pd.to_datetime(serie.where(~es_texto, pd.NaT) if es_texto is not None else serie)
        formato = "%d/%m/%Y" if clase == "fecha" else "%d/%m/%Y %H:%M:%S"
        salida = base.dt.strftime(formato).fillna("")
    if es_texto is not None and es_texto.any():
        salida = salida.where(~es_texto, serie)
    return salida


def _envolver(path: Path, accion) -> None:
    try:
        accion()
    except PermissionError as exc:
        raise ArchivoBloqueado(
            "No se pudo guardar el archivo. Si está abierto en Excel, cerralo y probá de nuevo.",
            str(exc),
        ) from exc
    except FileNotFoundError as exc:
        raise ExcelToolError(
            "No se pudo guardar el archivo: la carpeta no existe.", str(exc)
        ) from exc
    except OSError as exc:
        raise ExcelToolError(
            "No se pudo guardar el archivo. Revisá que haya espacio en el disco "
            "y que el nombre sea válido.",
            str(exc),
        ) from exc


def _escribir_atomico(path: Path, escribir: Callable[[Path], None]) -> None:
    """Escribe en un temporal junto al destino y recién al final lo pone en su lugar.

    Si se cancela o falla a mitad de camino, el destino queda como estaba y no sobra nada.
    """
    tmp = path.with_name(path.name + ".tmp")

    def accion() -> None:
        try:
            escribir(tmp)
            os.replace(tmp, path)
        except BaseException:
            with contextlib.suppress(OSError):
                tmp.unlink(missing_ok=True)
            raise

    _envolver(path, accion)


def _tabla_csv(
    df: pd.DataFrame, formato: str, column_types: dict[str, ColumnType] | None
) -> tuple[pd.DataFrame, dict[str, list[tuple[int, str]]]]:
    """Tabla de texto lista para `to_csv` y las celdas que no se pudieron convertir."""
    decimal = _CSV_FORMATOS[formato][2]
    series, clases, fallidas = _columnas_convertidas(df, column_types)
    salida = pd.concat(
        [
            _a_texto_csv(s, c, decimal).reset_index(drop=True)
            for s, c in zip(series, clases, strict=True)
        ],
        axis=1,
    )
    salida.columns = list(df.columns)
    return salida, fallidas


def csv_bytes_impl(df: pd.DataFrame, options: WriteOptions | None, header: bool) -> bytes:
    formato = options.format if options is not None else "csv_excel"
    if formato not in _CSV_FORMATOS:
        formato = "csv_excel"
    sep = _CSV_FORMATOS[formato][0]
    tipos = options.column_types if options is not None else None
    salida, _ = _tabla_csv(df, formato, tipos)
    buffer = io.StringIO()
    salida.to_csv(buffer, sep=sep, index=False, header=header, lineterminator="\r\n")
    return buffer.getvalue().encode("utf-8")


def _escribir_csv(
    df: pd.DataFrame,
    path: Path,
    formato: str,
    options: WriteOptions,
    progress: ProgressCallback | None,
    cancel: CancelToken | None,
) -> WriteResult:
    from excel_tool.core.io import WriteResult

    sep, encoding, _ = _CSV_FORMATOS[formato]
    salida, fallidas = _tabla_csv(df, formato, options.column_types)
    total = len(salida)
    report = ThrottledProgress(progress)

    def escribir(tmp: Path) -> None:
        with tmp.open("w", encoding=encoding, newline="") as f:
            # Por tandas: se puede cancelar y mostrar avance (con 0 filas sale solo el encabezado).
            for ini in range(0, max(total, 1), _TANDA_CSV):
                if cancel is not None:
                    cancel.check()
                salida.iloc[ini : ini + _TANDA_CSV].to_csv(
                    f, sep=sep, index=False, header=ini == 0, lineterminator="\r\n"
                )
                report(ProgressEvent("Guardando el archivo…", min(ini + _TANDA_CSV, total), total))

    _escribir_atomico(path, escribir)
    return WriteResult(path=path, rows=len(df), unconverted=fallidas)


def _limpiar_texto(serie: pd.Series) -> pd.Series:
    """Saca los caracteres de control que Excel no admite."""
    if serie.str.contains(_CONTROL_RE.pattern, regex=True).any():
        return serie.str.replace(_CONTROL_RE.pattern, "", regex=True)
    return serie


def _sin_control(texto: str) -> str:
    return _CONTROL_RE.sub("", texto)


class _Avisos:
    """Cuenta los casos raros de una hoja para armar los avisos al final."""

    def __init__(self) -> None:
        self.recortadas = 0
        self.fechas_viejas = 0

    def textos(self, nombre_hoja: str) -> list[str]:
        out = []
        if self.recortadas:
            out.append(
                f'En la hoja "{nombre_hoja}" hay {self.recortadas} celdas con más de '
                f"{_MAX_CELDA:,} caracteres: se recortaron porque Excel no admite más.".replace(
                    ",", "."
                )
            )
        if self.fechas_viejas:
            out.append(
                f'En la hoja "{nombre_hoja}" hay {self.fechas_viejas} fechas anteriores a 1900: '
                "se guardaron como texto porque Excel no las puede mostrar como fecha."
            )
        return out


def _recortar(valor: str, avisos: _Avisos) -> str:
    if len(valor) > _MAX_CELDA:
        avisos.recortadas += 1
        return valor[:_MAX_CELDA]
    return valor


def _valores_de_columna(serie: pd.Series, clase: str, avisos: _Avisos) -> list[Any]:
    """Valores listos para escribir. Tipos: None = vacío, str = texto, int/float = número,
    datetime = fecha. Las fechas anteriores a 1900 pasan a texto (y se cuentan)."""
    if clase == "texto":
        serie = _limpiar_texto(serie)
        largos = serie.str.len()
        if len(largos) and largos.max() > _MAX_CELDA:
            avisos.recortadas += int((largos > _MAX_CELDA).sum())
            serie = serie.str.slice(0, _MAX_CELDA)
        return [v or None for v in serie.tolist()]
    if clase in _NUMERICOS:
        cast = int if clase == "entero" else float
        return [
            _recortar(_sin_control(v), avisos)
            if isinstance(v, str)
            else (None if v is None or v != v else cast(v))
            for v in serie.tolist()
        ]
    salida: list[Any] = []
    formato = "%d/%m/%Y" if clase == "fecha" else "%d/%m/%Y %H:%M:%S"
    for v in serie.tolist():
        if isinstance(v, str):
            salida.append(_recortar(_sin_control(v), avisos))
        elif v is None or v is pd.NaT or v != v:
            salida.append(None)
        else:
            d = pd.Timestamp(v).to_pydatetime()
            if d.year < _ANIO_MINIMO:
                avisos.fechas_viejas += 1
                salida.append(d.strftime(formato))
            else:
                salida.append(d)
    return salida


def _llenar_hoja(
    ws,
    formatos: dict[str, Any],
    df: pd.DataFrame,
    column_types: dict[str, ColumnType] | None,
    freeze_header: bool,
    autofilter: bool,
    report: ThrottledProgress,
    cancel: CancelToken | None,
) -> tuple[dict[str, list[tuple[int, str]]], _Avisos]:
    if cancel is not None:
        cancel.check()
    avisos = _Avisos()
    series, clases, fallidas = _columnas_convertidas(df, column_types)
    nombres = [str(c) for c in df.columns]
    total = len(df)
    if freeze_header:
        ws.freeze_panes(1, 0)
    for j, nombre in enumerate(nombres):
        if nombre:
            ws.write_string(0, j, _recortar(_sin_control(nombre), avisos))
    columnas = [_valores_de_columna(s, c, avisos) for s, c in zip(series, clases, strict=True)]
    # Formato por columna; lo que quedó como str (texto o no convertible) va con write_string.
    fmts = [
        formatos["fecha"] if c == "fecha" else formatos["fecha_hora"] if c == "fecha_hora" else None
        for c in clases
    ]
    write_number, write_string, write_datetime = ws.write_number, ws.write_string, ws.write_datetime
    for i in range(total):
        if i % _CADA_FILAS == 0:
            if cancel is not None:
                cancel.check()
            report(ProgressEvent("Guardando el archivo…", i, total))
        fila = i + 1
        for j, col in enumerate(columnas):
            v = col[i]
            if v is None:
                continue
            if isinstance(v, str):
                write_string(fila, j, v)
            elif isinstance(v, datetime):
                write_datetime(fila, j, v, fmts[j])
            else:
                write_number(fila, j, v)
    report(ProgressEvent("Guardando el archivo…", total, total))
    if nombres and autofilter:
        ws.autofilter(0, 0, total, len(nombres) - 1)
    return fallidas, avisos


def _guardar_xlsx(
    hojas: Iterable[tuple[str, pd.DataFrame, dict[str, ColumnType] | None]],
    path: Path,
    *,
    freeze_header: bool = True,
    autofilter: bool = True,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> list[tuple[dict[str, list[tuple[int, str]]], int, list[str]]]:
    """Escribe las hojas una por una (`hojas` puede ser un generador: no se juntan en memoria).
    Devuelve por hoja: celdas sin convertir, filas y avisos."""
    usados: set[str] = set()
    resultados: list[tuple[dict[str, list[tuple[int, str]]], int, list[str]]] = []
    report = ThrottledProgress(progress)

    def escribir(tmp: Path) -> None:
        # Los temporales de constant_memory van a una carpeta propia que se borra siempre.
        with tempfile.TemporaryDirectory(
            prefix="exceltool_", ignore_cleanup_errors=True
        ) as carpeta:
            wb = xlsxwriter.Workbook(
                str(tmp),
                {
                    "constant_memory": True,
                    "tmpdir": carpeta,
                    "strings_to_formulas": False,
                    "strings_to_numbers": False,
                    "strings_to_urls": False,
                    "nan_inf_to_errors": False,
                },
            )
            try:
                formatos = {
                    "fecha": wb.add_format({"num_format": _FMT_FECHA}),
                    "fecha_hora": wb.add_format({"num_format": _FMT_FECHA_HORA}),
                }
                for nombre, df, tipos in hojas:
                    final = nombre_de_hoja(nombre, usados)
                    avisos = []
                    if final != str(nombre):
                        avisos.append(
                            f'El nombre de hoja "{nombre}" se cambió a "{final}" '
                            "porque Excel no admite esos caracteres o ese largo."
                        )
                    # constant_memory: cada hoja se escribe entera antes de crear la siguiente.
                    ws = wb.add_worksheet(final)
                    fallidas, extra = _llenar_hoja(
                        ws, formatos, df, tipos, freeze_header, autofilter, report, cancel
                    )
                    avisos += extra.textos(final)
                    resultados.append((fallidas, len(df), avisos))
                if not resultados:
                    raise ExcelToolError("No hay hojas para guardar.")
                wb.close()
            except BaseException:
                for ws in wb.worksheets():
                    with contextlib.suppress(Exception):
                        ws._opt_close()
                raise

    try:
        _escribir_atomico(path, escribir)
    except xlsxwriter.exceptions.FileCreateError as exc:
        # xlsxwriter envuelve los errores de disco: se vuelve al error original.
        causa = exc.__cause__ or exc.__context__
        if isinstance(causa, PermissionError):
            raise ArchivoBloqueado(
                "No se pudo guardar el archivo. Si está abierto en Excel, "
                "cerralo y probá de nuevo.",
                str(causa),
            ) from exc
        raise ExcelToolError(
            "No se pudo guardar el archivo. Revisá que haya espacio en el disco "
            "y que el nombre sea válido.",
            str(exc),
        ) from exc
    return resultados


def write_table_impl(
    df: pd.DataFrame,
    path: Path,
    options: WriteOptions,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> WriteResult:
    from excel_tool.core.io import WriteResult

    path = Path(path)
    if options.format == "xlsx":
        fallidas, filas, avisos = _guardar_xlsx(
            [(options.sheet_name, df, options.column_types)],
            path,
            freeze_header=options.freeze_header,
            autofilter=options.autofilter,
            progress=progress,
            cancel=cancel,
        )[0]
        return WriteResult(path=path, rows=filas, unconverted=fallidas, warnings=avisos)
    if options.format in _CSV_FORMATOS:
        return _escribir_csv(df, path, options.format, options, progress, cancel)
    raise ExcelToolError("Formato de salida no reconocido.", str(options.format))


def write_sheets_impl(
    sheets: Mapping[str, pd.DataFrame] | Iterable[tuple[str, pd.DataFrame]],
    path: Path,
    column_types: dict[str, dict[str, ColumnType]] | None = None,
    *,
    common_types: dict[str, ColumnType] | None = None,
    freeze_header: bool = True,
    autofilter: bool = True,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> list[WriteResult]:
    from excel_tool.core.io import WriteResult

    path = Path(path)
    pares = sheets.items() if isinstance(sheets, Mapping) else sheets

    def hojas():
        for nombre, df in pares:
            yield nombre, df, (column_types or {}).get(nombre, common_types)

    res = _guardar_xlsx(
        hojas(),
        path,
        freeze_header=freeze_header,
        autofilter=autofilter,
        progress=progress,
        cancel=cancel,
    )
    return [
        WriteResult(path=path, rows=filas, unconverted=f, warnings=avisos)
        for f, filas, avisos in res
    ]
