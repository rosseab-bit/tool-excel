"""Escritura de CSV y xlsx. Contrato: docs/contratos.md § Fase 1 (reglas de escritura).

Sin límites de tamaño ni división: eso es de `output.save_result` (Fase 3).
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import TYPE_CHECKING, Any

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter

from excel_tool.core.errors import ArchivoBloqueado, ExcelToolError
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


def _escribir_csv(df: pd.DataFrame, path: Path, formato: str, options: WriteOptions) -> WriteResult:
    from excel_tool.core.io import WriteResult

    sep, encoding, decimal = _CSV_FORMATOS[formato]
    series, clases, fallidas = _columnas_convertidas(df, options.column_types)
    salida = pd.concat(
        [
            _a_texto_csv(s, c, decimal).reset_index(drop=True)
            for s, c in zip(series, clases, strict=True)
        ],
        axis=1,
    )
    salida.columns = list(df.columns)

    def escribir() -> None:
        with path.open("w", encoding=encoding, newline="") as f:
            salida.to_csv(f, sep=sep, index=False, lineterminator="\r\n")

    _envolver(path, escribir)
    return WriteResult(path=path, rows=len(df), unconverted=fallidas)


def _limpiar_texto(serie: pd.Series) -> pd.Series:
    """Saca los caracteres de control que Excel no admite."""
    if serie.str.contains(ILLEGAL_CHARACTERS_RE.pattern, regex=True).any():
        return serie.str.replace(ILLEGAL_CHARACTERS_RE.pattern, "", regex=True)
    return serie


def _celdas_de_columna(serie: pd.Series, clase: str, ws) -> list[Any]:
    """Valores de una columna listos para `ws.append` (None = celda vacía)."""
    if clase == "texto":
        serie = _limpiar_texto(serie)
        valores = serie.tolist()
        return [_celda_texto(v, ws) if v else None for v in valores]
    if clase in _NUMERICOS:
        out: list[Any] = []
        for v in serie.tolist():
            if isinstance(v, str):
                out.append(_celda_texto(v, ws))
            elif v is None or v != v:
                out.append(None)
            elif clase == "entero":
                out.append(int(v))
            else:
                out.append(float(v))
        return out
    formato = _FMT_FECHA if clase == "fecha" else _FMT_FECHA_HORA
    out = []
    for v in serie.tolist():
        if isinstance(v, str):
            out.append(_celda_texto(v, ws))
        elif v is None or v is pd.NaT or v != v:
            out.append(None)
        else:
            celda = WriteOnlyCell(ws, value=pd.Timestamp(v).to_pydatetime())
            celda.number_format = formato
            out.append(celda)
    return out


def _celda_texto(valor: str, ws):
    """Texto que empieza con "=" se guarda como texto, nunca como fórmula."""
    if valor.startswith("="):
        celda = WriteOnlyCell(ws, value=valor)
        celda.data_type = "s"
        return celda
    return valor


def _llenar_hoja(ws, df: pd.DataFrame, column_types: dict[str, ColumnType] | None):
    series, clases, fallidas = _columnas_convertidas(df, column_types)
    nombres = [str(c) for c in df.columns]
    # En write_only la vista se escribe con la primera fila: congelar ANTES de agregar filas.
    ws.freeze_panes = "A2"
    ws.append([_celda_texto(_sin_control(n), ws) if n else None for n in nombres])
    columnas = [_celdas_de_columna(s, c, ws) for s, c in zip(series, clases, strict=True)]
    for fila in zip(*columnas, strict=True) if columnas else ():
        ws.append(list(fila))
    if nombres:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(nombres))}{len(df) + 1}"
    return fallidas


def _sin_control(texto: str) -> str:
    return ILLEGAL_CHARACTERS_RE.sub("", texto)


def _guardar_xlsx(hojas: list[tuple[str, pd.DataFrame, dict[str, ColumnType] | None]], path: Path):
    wb = Workbook(write_only=True)
    resultados = []
    for nombre, df, tipos in hojas:
        ws = wb.create_sheet(title=nombre)
        resultados.append(_llenar_hoja(ws, df, tipos))

    def escribir() -> None:
        try:
            wb.save(path)
        except BaseException:
            # Cierra los generadores pendientes de write_only (si no, avisa al recolectarlos).
            for ws in wb.worksheets:
                filas = getattr(ws, "_rows", None)
                if filas is not None and hasattr(filas, "close"):
                    filas.close()
            raise

    _envolver(path, escribir)
    return resultados


def write_table_impl(df: pd.DataFrame, path: Path, options: WriteOptions) -> WriteResult:
    from excel_tool.core.io import WriteResult

    path = Path(path)
    if options.format == "xlsx":
        nombre = nombre_de_hoja(options.sheet_name, set())
        fallidas = _guardar_xlsx([(nombre, df, options.column_types)], path)[0]
        return WriteResult(path=path, rows=len(df), unconverted=fallidas)
    if options.format in _CSV_FORMATOS:
        return _escribir_csv(df, path, options.format, options)
    raise ExcelToolError("Formato de salida no reconocido.", str(options.format))


def write_sheets_impl(
    sheets: dict[str, pd.DataFrame],
    path: Path,
    column_types: dict[str, dict[str, ColumnType]] | None = None,
) -> list[WriteResult]:
    from excel_tool.core.io import WriteResult

    path = Path(path)
    usados: set[str] = set()
    hojas = []
    for nombre, df in sheets.items():
        hojas.append((nombre_de_hoja(nombre, usados), df, (column_types or {}).get(nombre)))
    if not hojas:
        raise ExcelToolError("No hay hojas para guardar.")
    fallidas = _guardar_xlsx(hojas, path)
    return [
        WriteResult(path=path, rows=len(df), unconverted=f)
        for (_, df, _), f in zip(hojas, fallidas, strict=True)
    ]
