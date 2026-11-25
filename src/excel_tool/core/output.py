"""Guardar un resultado: un archivo, varias hojas o partes en una carpeta.

Contrato: docs/contratos.md § Fase 3 (`core/output.py`).

Decisiones de diseño:
- Antes de escribir se controla el tamaño (`limits.check`) y se tira un `ExcelToolError` si el
  resultado no entra y no se pidió dividir.
- Cada parte se escribe apenas se genera (`split.split_dataframe` es un iterador).
  En varias hojas, el iterador va a `io.write_sheets`, que escribe hoja por hoja.
- El límite duro por parte (filas de Excel o celdas de Google Sheets) se pasa a `split` como
  `max_rows_per_part`: ninguna parte lo supera.
- Cancelación o error: se borra solo lo que creó esta llamada (archivos de la carpeta y la carpeta
  si quedó vacía). Un archivo suelto que ya existía no se toca hasta el final (escritura atómica
  en `core/io`).
- Las piezas `escribir_partes` y `escribir_indice` están separadas para que `split_file` (F3-05)
  las reutilice.

`split_file` (CSV -> partes, sin cargar el archivo entero):
- Se lee por bloques (`io.CsvBlockReader`) y cada parte se escribe a medida que llega
  (`io.CsvPartWriter`). La memoria queda acotada por el tamaño de un bloque (100.000 filas).
- Como el total de partes no se conoce hasta el final, las partes se escriben con un nombre
  temporal dentro de la carpeta y recién al terminar se renombran al nombre final (que puede
  llevar "parte 3 de 7"). Con salida xlsx cada parte se escribe primero como CSV temporal y se
  convierte a xlsx al cerrarla (una parte pesa, como mucho, el límite de Excel).
- `rows`: una pasada. `by_value`: una pasada con un archivo abierto por valor (tope
  `MAX_ARCHIVOS_ABIERTOS`; si hay más valores, error claro). `rows_keep_groups`: DOS pasadas
  (la primera cuenta las filas de cada grupo y arma las partes; la segunda escribe). Las filas
  de cada parte conservan el orden original del archivo (no se reordenan por grupo como en
  `split.split_dataframe`). Igual que `by_value`, tiene el tope de archivos abiertos.
- Los valores se copian como texto, tal cual: `OutputSpec.column_types` solo se usa con xlsx.
- `sheets` no aplica (el resultado de este modo es una carpeta de archivos CSV/xlsx).
"""

from __future__ import annotations

import contextlib
import dataclasses
import logging
import os
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from excel_tool.core import limits as limits_mod
from excel_tool.core.errors import ArchivoIlegible, ArchivoNoEncontrado, ExcelToolError
from excel_tool.core.io import (
    CsvBlockReader,
    CsvOptions,
    CsvPartWriter,
    ReadOptions,
    WriteFormat,
    WriteOptions,
    detect_csv_options,
    read_csv_exact,
    write_sheets,
    write_table,
)
from excel_tool.core.io.writer import nombre_de_hoja
from excel_tool.core.limits import Destination, Limits
from excel_tool.core.progress import (
    CancelToken,
    ProgressCallback,
    ProgressEvent,
    ThrottledProgress,
)
from excel_tool.core.split import (
    PartPlan,
    SplitSpec,
    _limite,
    _Nombres,
    plan_split,
    split_dataframe,
)
from excel_tool.core.types import ColumnType

logger = logging.getLogger(__name__)

NOMBRE_INDICE = "_indice.xlsx"
COLUMNAS_INDICE = ["Archivo", "Hoja", "Filas", "Desde fila", "Hasta fila", "Valor"]
_TIPOS_INDICE: dict[str, ColumnType] = {
    "Filas": "entero",
    "Desde fila": "entero",
    "Hasta fila": "entero",
}
_MAX_RUTA = 250  # Windows corta en 260; se deja margen
_MB = 1024 * 1024
_STR = pd.StringDtype("pyarrow")
MAX_ARCHIVOS_ABIERTOS = 50  # split_file: partes abiertas a la vez (by_value / rows_keep_groups)
_EXT_CSV = {".csv", ".txt"}
_TEMPORAL = "_parte_tmp_{n:05d}.part"
_CSV_TEMPORAL = "csv_google"  # estilo de las partes intermedias cuando la salida es xlsx


@dataclass(frozen=True)
class OutputSpec:
    path: Path
    format: WriteFormat
    destination: Destination
    split: SplitSpec | None = None
    write_index: bool = True
    sheet_name: str = "Datos"
    column_types: dict[str, ColumnType] | None = None


@dataclass(frozen=True)
class PartResult:
    path: Path
    sheet: str | None
    rows: int
    value: tuple[str, ...] | None
    first_row: int
    last_row: int


@dataclass
class SaveResult:
    files: list[Path]
    folder: Path | None
    parts: list[PartResult]
    rows_total: int
    warnings: list[str]
    unconverted: dict[str, int]


# ------------------------------------------------------------------------------ utilidades


def _extension(formato: WriteFormat) -> str:
    return ".xlsx" if formato == "xlsx" else ".csv"


def _n(valor: int) -> str:
    return f"{valor:,}".replace(",", ".")


def max_filas_por_parte(
    n_columnas: int, destination: Destination, formato: WriteFormat, lim: Limits
) -> int | None:
    """Máximo de filas de datos que puede tener una parte según el destino (None = sin tope)."""
    if destination == "google_sheets":
        tope = lim.gs_max_cells // max(n_columnas, 1) - 1
        if tope < 1:
            raise ExcelToolError(
                "Esta tabla tiene demasiadas columnas para Google Sheets: "
                "ni una sola fila entra en el límite de celdas."
            )
        return tope
    if destination == "excel" or formato == "xlsx":
        return lim.excel_max_rows - 1
    return None


def _controlar(df: pd.DataFrame, spec: OutputSpec, lim: Limits) -> int | None:
    """Controles previos a escribir. Devuelve el tope duro de filas por parte."""
    split = spec.split
    if split is not None and split.mode == "sheets":
        if spec.destination == "google_sheets":
            raise ExcelToolError(
                "Google Sheets no acepta un archivo con varias hojas para importar. "
                "Elegí dividir en varios archivos."
            )
        if spec.format != "xlsx":
            raise ExcelToolError("Dividir en varias hojas solo se puede con archivos de Excel.")
    est = limits_mod.estimate(df)
    chk = limits_mod.check(est, spec.destination, spec.format, limits=lim)
    if chk.level == "obligatorio":
        en_excel = spec.destination == "excel" or (
            spec.destination == "solo_guardar" and spec.format == "xlsx"
        )
        columnas_de_mas = en_excel and est.columns > lim.excel_max_columns
        if split is None or columnas_de_mas:
            raise ExcelToolError(
                "El resultado no entra en un solo archivo. " + " ".join(chk.reasons),
                "; ".join(chk.reasons),
            )
    return max_filas_por_parte(est.columns, spec.destination, spec.format, lim)


class _Avance:
    """Convierte el avance por filas de cada parte en uno global (por parte y por filas)."""

    def __init__(self, progress: ProgressCallback | None, total_filas: int, total_partes: int):
        self._report = ThrottledProgress(progress)
        self.total = total_filas
        self.partes = total_partes
        self.hechas = 0  # filas de las partes ya terminadas
        self.parte = 0

    def empezar_parte(self, n: int) -> None:
        self.parte = n
        self._report(self._evento(0))

    def _evento(self, locales: int) -> ProgressEvent:
        if self.partes > 1:
            mensaje = f"Guardando la parte {self.parte} de {self.partes}…"
        else:
            mensaje = "Guardando el archivo…"
        return ProgressEvent(mensaje, min(self.hechas + locales, self.total), self.total)

    def __call__(self, ev: ProgressEvent) -> None:
        self._report(self._evento(ev.done))

    def terminar_parte(self, filas: int) -> None:
        self.hechas += filas
        self._report.force(self._evento(0))


def _sumar_no_convertidas(acum: dict[str, int], fallidas: dict[str, list]) -> None:
    for col, lista in fallidas.items():
        acum[col] = acum.get(col, 0) + len(lista)


def _agregar_avisos(destino: list[str], nuevos: Iterable[str]) -> None:
    for aviso in nuevos:
        if aviso not in destino:
            destino.append(aviso)


def _carpeta_nueva(base: Path) -> Path:
    """Crea `<base>`; si ya existe, `<base> (2)`, `(3)`... Nunca pisa nada."""
    n = 1
    while True:
        carpeta = base if n == 1 else base.with_name(f"{base.name} ({n})")
        try:
            carpeta.mkdir()
        except FileExistsError:
            n += 1
            continue
        except FileNotFoundError as exc:
            raise ExcelToolError(
                "No se pudo crear la carpeta para las partes: la carpeta de destino no existe.",
                str(exc),
            ) from exc
        except OSError as exc:
            raise ExcelToolError(
                "No se pudo crear la carpeta para las partes. Revisá el nombre y los permisos.",
                str(exc),
            ) from exc
        return carpeta


def _limpiar(creados: list[Path], carpeta: Path | None) -> None:
    for archivo in creados:
        with contextlib.suppress(OSError):
            archivo.unlink(missing_ok=True)
    if carpeta is not None:
        with contextlib.suppress(OSError):
            carpeta.rmdir()  # solo si quedó vacía


def _nombre_de_archivo(carpeta: Path, nombre: str, ext: str, n: int, usados: set[str]) -> str:
    """Nombre de archivo (con extensión) que mantiene la ruta completa dentro de `_MAX_RUTA`."""
    completo = nombre + ext
    if len(str(carpeta / completo)) <= _MAX_RUTA:
        usados.add(completo.lower())
        return completo
    permitido = _MAX_RUTA - len(str(carpeta)) - 1 - len(ext)
    etiqueta = f"_{n:03d}"
    if permitido < len(etiqueta) + 8:
        raise ExcelToolError(
            "La ruta de la carpeta elegida es demasiado larga para guardar los archivos. "
            "Elegí una carpeta con una ruta más corta.",
            str(carpeta),
        )
    corto = nombre[: permitido - len(etiqueta)].rstrip(" .") + etiqueta + ext
    k = 2
    while corto.lower() in usados:
        sufijo = f"{etiqueta}_{k}"
        corto = nombre[: permitido - len(sufijo)].rstrip(" .") + sufijo + ext
        k += 1
    usados.add(corto.lower())
    return corto


# ------------------------------------------------------------------------- partes y índice


def escribir_partes(
    partes: Iterable[tuple[PartPlan, pd.DataFrame]],
    carpeta: Path,
    *,
    options: WriteOptions,
    avance: _Avance,
    creados: list[Path],
    warnings: list[str],
    unconverted: dict[str, int],
    cancel: CancelToken | None,
    gs_max_mb: float | None = None,
) -> list[PartResult]:
    """Escribe cada parte en `carpeta` apenas llega (sin juntarlas). Anota cada archivo en
    `creados` antes de escribirlo, para poder borrarlo si algo falla."""
    ext = _extension(options.format)
    usados: set[str] = set()
    resultados: list[PartResult] = []
    acortados = False
    for n, (plan, df) in enumerate(partes, start=1):
        if cancel is not None:
            cancel.check()
        archivo = _nombre_de_archivo(carpeta, plan.name, ext, n, usados)
        acortados = acortados or archivo != plan.name + ext
        destino = carpeta / archivo
        avance.empezar_parte(n)
        creados.append(destino)
        res = write_table(df, destino, options, progress=avance, cancel=cancel)
        _agregar_avisos(warnings, res.warnings)
        _sumar_no_convertidas(unconverted, res.unconverted)
        if gs_max_mb is not None and destino.stat().st_size > gs_max_mb * _MB:
            peso = f"{destino.stat().st_size / _MB:,.0f}".replace(",", ".")
            warnings.append(
                f"La parte {n} ({archivo}) pesa {peso} MB: Google Sheets puede no importarla."
            )
        hoja = nombre_de_hoja(options.sheet_name) if options.format == "xlsx" else None
        resultados.append(
            PartResult(destino, hoja, res.rows, plan.value, plan.first_row, plan.last_row)
        )
        avance.terminar_parte(res.rows)
    if acortados:
        warnings.append(
            "Algunos nombres de archivo se acortaron porque la ruta completa pasaba "
            f"los {_MAX_RUTA} caracteres (Windows no los admite)."
        )
    return resultados


def escribir_indice(
    partes: list[PartResult], carpeta: Path, cancel: CancelToken | None = None
) -> Path:
    """`_indice.xlsx`: una fila por parte (Archivo, Hoja, Filas, Desde fila, Hasta fila, Valor)."""
    filas = [
        [
            p.path.name,
            p.sheet or "",
            str(p.rows),
            str(p.first_row),
            str(p.last_row),
            " - ".join(v if v != "" else "(vacío)" for v in p.value) if p.value else "",
        ]
        for p in partes
    ]
    tabla = pd.DataFrame(filas, columns=COLUMNAS_INDICE, dtype=object).astype(_STR)
    destino = carpeta / NOMBRE_INDICE
    write_table(
        tabla,
        destino,
        WriteOptions(format="xlsx", sheet_name="Índice", column_types=_TIPOS_INDICE),
        cancel=cancel,
    )
    return destino


def avisos_de_grupos(plans: list[PartPlan], split: SplitSpec, tope: int | None) -> list[str]:
    """`rows_keep_groups`: partes que superan lo pedido porque un mismo valor no se corta."""
    if split.mode != "rows_keep_groups":
        return []
    candidatos = [v for v in (split.rows_per_part, tope) if v is not None]
    if not candidatos:
        return []
    limite = min(candidatos)
    col = split.group_column or "esa columna"
    return [
        f"La parte {i} tiene {_n(p.rows)} filas, más que las {_n(limite)} pedidas: "
        f"un mismo valor de «{col}» no se puede cortar entre archivos."
        for i, p in enumerate(plans, start=1)
        if p.rows > limite
    ]


# ------------------------------------------------------------------------------- principal


def save_result(
    df: pd.DataFrame,
    spec: OutputSpec,
    *,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> SaveResult:
    """Guarda `df` según `spec`. Ver el contrato en docs/contratos.md."""
    lim = limits_mod.load_limits()
    tope = _controlar(df, spec, lim)
    path = Path(spec.path)
    total = len(df)
    opciones = WriteOptions(
        format=spec.format, sheet_name=spec.sheet_name, column_types=spec.column_types
    )
    warnings: list[str] = []
    unconverted: dict[str, int] = {}
    split = spec.split

    # Sin filas (o sin división): un solo archivo; con 0 filas queda solo el encabezado.
    if split is None or total == 0:
        avance = _Avance(progress, total, 1)
        avance.empezar_parte(1)
        res = write_table(df, path, opciones, progress=avance, cancel=cancel)
        _agregar_avisos(warnings, res.warnings)
        _sumar_no_convertidas(unconverted, res.unconverted)
        hoja = nombre_de_hoja(spec.sheet_name) if spec.format == "xlsx" else None
        parte = PartResult(path, hoja, res.rows, None, 1 if total else 0, total)
        return _cerrar([path], None, [parte], total, warnings, unconverted)

    if split.mode == "sheets":
        return _guardar_hojas(df, spec, tope, progress, cancel)

    # Partes en una carpeta.
    plans = plan_split(df, split, base_name=path.stem, max_rows_per_part=tope)
    _agregar_avisos(warnings, avisos_de_grupos(plans, split, tope))
    avance = _Avance(progress, total, len(plans))
    creados: list[Path] = []
    carpeta = _carpeta_nueva(path.parent / _nombre_carpeta(path))
    try:
        partes = escribir_partes(
            split_dataframe(df, split, base_name=path.stem, max_rows_per_part=tope),
            carpeta,
            options=opciones,
            avance=avance,
            creados=creados,
            warnings=warnings,
            unconverted=unconverted,
            cancel=cancel,
            gs_max_mb=lim.gs_max_import_mb if spec.destination == "google_sheets" else None,
        )
        _verificar(partes, total)
        if spec.write_index:
            indice = carpeta / NOMBRE_INDICE
            creados.append(indice)
            escribir_indice(partes, carpeta, cancel)
        return _cerrar(list(creados), carpeta, partes, total, warnings, unconverted)
    except BaseException:
        _limpiar(creados, carpeta)
        raise


def _nombre_carpeta(path: Path) -> str:
    """Nombre de la carpeta de partes: el del archivo elegido, sin extensión."""
    return path.stem or "resultado"


def _verificar(partes: list[PartResult], total: int) -> None:
    suma = sum(p.rows for p in partes)
    if suma != total:
        logger.error(
            "La suma de filas de las partes (%s) no coincide con el total (%s)", suma, total
        )
        raise ExcelToolError(
            "Algo salió mal al guardar: la cantidad de filas de los archivos no coincide "
            "con la del resultado. No se guardó nada.",
            f"suma de partes={suma}, total={total}",
        )


def _cerrar(
    archivos: list[Path],
    carpeta: Path | None,
    partes: list[PartResult],
    total: int,
    warnings: list[str],
    unconverted: dict[str, int],
) -> SaveResult:
    _verificar(partes, total)
    return SaveResult(
        files=archivos,
        folder=carpeta,
        parts=partes,
        rows_total=total,
        warnings=warnings,
        unconverted=unconverted,
    )


def _guardar_hojas(
    df: pd.DataFrame,
    spec: OutputSpec,
    tope: int | None,
    progress: ProgressCallback | None,
    cancel: CancelToken | None,
) -> SaveResult:
    """Un .xlsx con una hoja por parte. Las hojas se escriben una por una (iterador)."""
    split = spec.split
    assert split is not None
    path = Path(spec.path)
    total = len(df)
    plans = plan_split(df, split, base_name=path.stem, max_rows_per_part=tope)
    warnings: list[str] = []
    _agregar_avisos(warnings, avisos_de_grupos(plans, split, tope))
    avance = _Avance(progress, total, len(plans))
    vistos: list[PartPlan] = []

    def hojas() -> Iterator[tuple[str, pd.DataFrame]]:
        for n, (plan, parte) in enumerate(
            split_dataframe(df, split, base_name=path.stem, max_rows_per_part=tope), start=1
        ):
            if vistos:
                avance.terminar_parte(vistos[-1].rows)
            vistos.append(plan)
            avance.empezar_parte(n)
            yield plan.name, parte

    resultados = write_sheets(
        hojas(),
        path,
        common_types=spec.column_types,
        progress=avance,
        cancel=cancel,
    )
    unconverted: dict[str, int] = {}
    partes: list[PartResult] = []
    for plan, res in zip(vistos, resultados, strict=True):
        _agregar_avisos(warnings, res.warnings)
        _sumar_no_convertidas(unconverted, res.unconverted)
        partes.append(
            PartResult(path, plan.name, res.rows, plan.value, plan.first_row, plan.last_row)
        )
    return _cerrar([path], None, partes, total, warnings, unconverted)


# ------------------------------------------------------------------------------ split_file


class _Parte:
    """Una parte en construcción: su archivo temporal, el escritor abierto y los contadores."""

    def __init__(self, n: int, valor: tuple[str, ...] | None, tmp: Path, writer: CsvPartWriter):
        self.n = n
        self.valor = valor
        self.tmp = tmp
        self.writer = writer
        self.rows = 0
        self.first = 0
        self.last = 0

    def agregar(self, bloque: pd.DataFrame, posiciones_ini: int, posiciones_fin: int) -> None:
        """`posiciones_*`: filas del original (desde 1) más chica y más grande del trozo."""
        if self.rows == 0:
            self.first = posiciones_ini
        else:
            self.first = min(self.first, posiciones_ini)
        self.last = max(self.last, posiciones_fin)
        self.writer.write(bloque)
        self.rows += len(bloque)

    def cerrar(self) -> None:
        self.writer.close()


class _Dividir:
    """Estado de una corrida de `split_file`."""

    def __init__(
        self,
        reader: CsvBlockReader,
        spec: OutputSpec,
        carpeta: Path,
        limite: int | None,
        creados: list[Path],
        progress: ProgressCallback | None,
        cancel: CancelToken | None,
    ):
        self.reader = reader
        self.spec = spec
        self.carpeta = carpeta
        self.limite = limite
        self.creados = creados
        self.cancel = cancel
        self.report = ThrottledProgress(progress)
        self.columnas = reader.info.columns
        self.partes: list[_Parte] = []
        self.formato_temporal = spec.format if spec.format != "xlsx" else _CSV_TEMPORAL

    def nueva(self, valor: tuple[str, ...] | None) -> _Parte:
        tmp = self.carpeta / _TEMPORAL.format(n=len(self.partes) + 1)
        self.creados.append(tmp)
        parte = _Parte(
            len(self.partes) + 1,
            valor,
            tmp,
            CsvPartWriter(tmp, self.formato_temporal, self.columnas),
        )
        self.partes.append(parte)
        return parte

    def avanzar(self, mensaje: str) -> None:
        if self.cancel is not None:
            self.cancel.check()
        total = self.reader.info.total_bytes
        self.report(ProgressEvent(mensaje, self.reader.bytes_read, total))

    def cerrar_todo(self) -> None:
        for p in self.partes:
            p.cerrar()


def _escribir_en_partes(
    d: _Dividir,
    valor: tuple[str, ...] | None,
    actual: _Parte | None,
    trozo: pd.DataFrame,
    base: int,
) -> _Parte:
    """Escribe `trozo` en `actual` (abriendo partes nuevas si se llena el límite). `base`: número
    de fila del original (desde 1) de la primera fila del trozo, para un rango continuo.
    Devuelve la parte que queda abierta."""
    pos = 0
    n = len(trozo)
    while pos < n:
        if actual is None:
            actual = d.nueva(valor)
        espacio = n - pos if d.limite is None else d.limite - actual.rows
        pedazo = trozo.iloc[pos : pos + espacio]
        actual.agregar(pedazo, base + pos, base + pos + len(pedazo) - 1)
        pos += len(pedazo)
        if d.limite is not None and actual.rows >= d.limite:
            actual.cerrar()
            actual = None
    return actual  # type: ignore[return-value]


def _pasada_filas(d: _Dividir) -> None:
    actual: _Parte | None = None
    for bloque in d.reader:
        base = d.reader.rows - len(bloque) + 1
        actual = _escribir_en_partes(d, None, actual, bloque, base)
        d.avanzar("Dividiendo el archivo…")
    d.cerrar_todo()


def _clave_valor(valor: tuple[str, ...]) -> tuple:
    """Mismo orden que `split.by_value`: sin distinguir mayúsculas, el vacío al final."""
    return tuple((v == "", v.casefold(), v) for v in valor)


def _error_demasiados(que: str, hasta: int) -> ExcelToolError:
    return ExcelToolError(
        f"Hay demasiados {que} para dividir un archivo sin abrirlo entero (el máximo es "
        f"{hasta}). Elegí otra columna o dividí por cantidad de filas.",
        f"{que} > {hasta}",
    )


def _pasada_valores(d: _Dividir, columnas: list[str]) -> dict[tuple[str, ...], list[_Parte]]:
    por_valor: dict[tuple[str, ...], list[_Parte]] = {}
    abiertas: dict[tuple[str, ...], _Parte | None] = {}
    for bloque in d.reader:
        base = d.reader.rows - len(bloque) + 1  # fila (desde 1) de la primera del bloque
        for clave, sub in bloque.groupby(columnas, sort=False, dropna=False):
            valor = clave if isinstance(clave, tuple) else (clave,)
            valor = tuple(str(v) for v in valor)
            if valor not in por_valor:
                if len(por_valor) >= MAX_ARCHIVOS_ABIERTOS:
                    raise _error_demasiados("valores distintos", MAX_ARCHIVOS_ABIERTOS)
                por_valor[valor] = []
                abiertas[valor] = None
            filas = sub.index.to_numpy()
            actual = abiertas[valor]
            n = len(sub)
            pos = 0
            while pos < n:
                if actual is None:
                    actual = d.nueva(valor)
                    por_valor[valor].append(actual)
                espacio = n - pos if d.limite is None else d.limite - actual.rows
                pedazo = sub.iloc[pos : pos + espacio]
                f = filas[pos : pos + len(pedazo)]
                actual.agregar(pedazo, base + int(f.min()), base + int(f.max()))
                pos += len(pedazo)
                if d.limite is not None and actual.rows >= d.limite:
                    actual.cerrar()
                    actual = None
            abiertas[valor] = actual
        d.avanzar("Dividiendo el archivo…")
    d.cerrar_todo()
    return por_valor


def _contar_grupos(d: _Dividir, columna: str) -> dict[str, int]:
    """Primera pasada de `rows_keep_groups`: filas de cada valor, en orden de aparición."""
    cuentas: dict[str, int] = {}
    for bloque in d.reader:
        serie = bloque[columna].fillna("")
        vc = serie.value_counts(sort=False)
        for valor in pd.unique(serie):
            cuentas[valor] = cuentas.get(valor, 0) + int(vc[valor])
        d.avanzar("Contando las filas de cada grupo…")
    return cuentas


def _armar_grupos(cuentas: dict[str, int], limite: int) -> tuple[dict[str, int], list[int]]:
    """Junta grupos enteros en partes de hasta `limite` filas. Un grupo más grande va solo."""
    parte_de: dict[str, int] = {}
    filas_de: list[int] = []
    actual = -1
    for valor, c in cuentas.items():
        if actual < 0 or filas_de[actual] + c > limite:
            filas_de.append(0)
            actual += 1
        parte_de[valor] = actual
        filas_de[actual] += c
    return parte_de, filas_de


def _pasada_grupos(d: _Dividir, columna: str, parte_de: dict[str, int], n_partes: int) -> None:
    abiertas: dict[int, _Parte] = {}
    for bloque in d.reader:
        base = d.reader.rows - len(bloque) + 1
        indices = bloque[columna].fillna("").map(parte_de)
        if indices.isna().any():
            raise ExcelToolError(
                "El archivo cambió mientras se dividía. Probá de nuevo.", "grupo sin parte"
            )
        indices = indices.astype(int)
        for k, sub in bloque.groupby(indices, sort=True):
            k = int(k)  # noqa: PLW2901
            parte = abiertas.get(k)
            if parte is None:
                parte = d.nueva(None)
                abiertas[k] = parte
            filas = sub.index.to_numpy()
            parte.agregar(sub, base + int(filas.min()), base + int(filas.max()))
        d.avanzar("Dividiendo el archivo…")
    d.cerrar_todo()
    # `d.partes` quedó en orden de creación; se reordena por número de parte de grupo.
    orden = {id(p): k for k, p in abiertas.items()}
    d.partes.sort(key=lambda p: orden[id(p)])


def _validar_columnas(columnas: list[str], pedidas: list[str] | None, que: str) -> list[str]:
    if not pedidas:
        raise ExcelToolError(f"Elegí {que}.")
    faltan = [c for c in pedidas if c not in columnas]
    if faltan:
        raise ExcelToolError(
            "No encuentro estas columnas en el archivo: " + ", ".join(faltan) + ".", str(faltan)
        )
    return list(pedidas)


def _verificar_origen(path: Path) -> None:
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xlsb", ".xls", ".ods"):
        raise ExcelToolError(
            "Por ahora solo se pueden dividir archivos CSV sin abrirlos enteros; "
            "para Excel, abrilo con 'Ver información' y usá Guardar.",
            str(path),
        )
    if ext not in _EXT_CSV:
        from excel_tool.core.errors import FormatoNoSoportado

        raise FormatoNoSoportado(
            "Este tipo de archivo no se puede abrir. Usá un archivo de texto (.csv, .txt).",
            str(path),
        )
    if not path.exists():
        raise ArchivoNoEncontrado(
            "No se encontró el archivo. Fijate si se movió o se borró.", str(path)
        )


def split_file(
    path: Path,
    read: ReadOptions,
    spec: OutputSpec,
    *,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> SaveResult:
    """Divide un CSV en partes sin cargarlo entero en memoria. Ver el módulo y el contrato.

    `spec.split` es obligatorio (`rows`, `rows_keep_groups` o `by_value`). Las partes van a una
    carpeta nueva (nunca se pisa nada), con `_indice.xlsx` si `spec.write_index`.
    """
    path = Path(path)
    _verificar_origen(path)
    split = spec.split
    if split is None:
        raise ExcelToolError("Indicá cómo dividir el archivo.")
    if split.mode == "sheets":
        raise ExcelToolError("Dividir en hojas no se puede con un archivo CSV sin abrirlo entero.")
    if split.mode not in ("rows", "rows_keep_groups", "by_value"):
        raise ExcelToolError("No conozco esa forma de dividir.", str(split.mode))

    if read.csv is None:
        read = dataclasses.replace(read, csv=detect_csv_options(path))
    try:
        return _dividir(path, read, spec, progress, cancel)
    except UnicodeDecodeError as exc:
        assert read.csv is not None
        if read.csv.encoding == "cp1252":
            raise ArchivoIlegible(
                "No se pudo leer el texto del archivo (caracteres no reconocidos).", str(exc)
            ) from exc
        logger.info("Reintentando %s con cp1252: %s", path, exc)
        read = dataclasses.replace(
            read, csv=CsvOptions(separator=read.csv.separator, encoding="cp1252")
        )
        return _dividir(path, read, spec, progress, cancel)


def _dividir(
    path: Path,
    read: ReadOptions,
    spec: OutputSpec,
    progress: ProgressCallback | None,
    cancel: CancelToken | None,
) -> SaveResult:
    split = spec.split
    assert split is not None
    lim = limits_mod.load_limits()
    out = Path(spec.path)
    warnings: list[str] = []
    creados: list[Path] = []
    if cancel is not None:
        cancel.check()
    with CsvBlockReader(path, read) as reader:
        info = reader.info
        _agregar_avisos(warnings, info.warnings)
        n_cols = len(info.columns)
        if (spec.destination == "excel" or spec.format == "xlsx") and (
            n_cols > lim.excel_max_columns
        ):
            raise ExcelToolError(
                f"El archivo tiene {_n(n_cols)} columnas y Excel acepta hasta "
                f"{_n(lim.excel_max_columns)}. Elegí guardar en CSV."
            )
        tope = max_filas_por_parte(n_cols, spec.destination, spec.format, lim)
        if split.mode == "by_value":
            columnas = _validar_columnas(
                info.columns, split.value_columns, "la columna por la que dividir"
            )
            candidatos = [v for v in (split.rows_per_part, tope) if v is not None]
            limite = int(min(candidatos)) if candidatos else None
            if limite is not None and limite < 1:
                raise ExcelToolError("El máximo de filas por parte tiene que ser mayor que cero.")
        else:
            limite = _limite(split, tope)
            columnas = (
                _validar_columnas(
                    info.columns,
                    [split.group_column] if split.group_column else None,
                    "la columna de grupos",
                )
                if split.mode == "rows_keep_groups"
                else []
            )
        carpeta = _carpeta_nueva(out.parent / _nombre_carpeta(out))
        try:
            d = _Dividir(reader, spec, carpeta, limite, creados, progress, cancel)
            if split.mode == "rows":
                _pasada_filas(d)
                orden = [(p, None, None) for p in d.partes]
            elif split.mode == "by_value":
                por_valor = _pasada_valores(d, columnas)
                orden = []
                for valor in sorted(por_valor, key=_clave_valor):
                    ps = por_valor[valor]
                    orden += [
                        (p, valor, (k + 1, len(ps)) if len(ps) > 1 else None)
                        for k, p in enumerate(ps)
                    ]
            else:
                assert limite is not None
                cuentas = _contar_grupos(d, columnas[0])
                parte_de, filas_de = _armar_grupos(cuentas, limite)
                if len(filas_de) > MAX_ARCHIVOS_ABIERTOS:
                    raise _error_demasiados("archivos de salida a la vez", MAX_ARCHIVOS_ABIERTOS)
                # Segunda pasada: se vuelve a leer desde el principio.
                d.reader.close()
                with CsvBlockReader(path, read) as segunda:
                    d.reader = segunda
                    _pasada_grupos(d, columnas[0], parte_de, len(filas_de))
                orden = [(p, None, None) for p in d.partes]
            total = d.reader.rows
            if total == 0:
                raise _SinFilas
            partes = _finalizar(d, orden, out, lim, warnings, cancel, progress)
            if split.mode == "rows_keep_groups":
                plans = [
                    PartPlan(p.path.name, p.rows, None, p.first_row, p.last_row) for p in partes
                ]
                _agregar_avisos(warnings, avisos_de_grupos(plans, split, tope))
            _verificar(partes, total)
            archivos = [p.path for p in partes]
            if spec.write_index:
                indice = carpeta / NOMBRE_INDICE
                creados.append(indice)
                escribir_indice(partes, carpeta, cancel)
                archivos.append(indice)
            return _cerrar(archivos, carpeta, partes, total, warnings, {})
        except _SinFilas:
            _limpiar(creados, carpeta)
        except BaseException:
            _limpiar(creados, carpeta)
            raise
    return _archivo_vacio(info.columns, spec, out, warnings, cancel)


class _SinFilas(Exception):
    """El archivo no tiene filas de datos: se guarda solo el encabezado en un archivo."""


def _archivo_vacio(
    columnas: list[str],
    spec: OutputSpec,
    out: Path,
    warnings: list[str],
    cancel: CancelToken | None,
) -> SaveResult:
    vacio = pd.DataFrame({c: pd.Series([], dtype=_STR) for c in columnas})
    opciones = WriteOptions(
        format=spec.format, sheet_name=spec.sheet_name, column_types=spec.column_types
    )
    res = write_table(vacio, out, opciones, cancel=cancel)
    _agregar_avisos(warnings, res.warnings)
    hoja = nombre_de_hoja(spec.sheet_name) if spec.format == "xlsx" else None
    parte = PartResult(out, hoja, 0, None, 0, 0)
    return _cerrar([out], None, [parte], 0, warnings, {})


def _finalizar(
    d: _Dividir,
    orden: list[tuple[_Parte, tuple[str, ...] | None, tuple[int, int] | None]],
    out: Path,
    lim: Limits,
    warnings: list[str],
    cancel: CancelToken | None,
    progress: ProgressCallback | None,
) -> list[PartResult]:
    """Pone nombre final a cada parte (renombra el temporal, o lo convierte a xlsx)."""
    split = d.spec.split
    assert split is not None
    ext = _extension(d.spec.format)
    total = len(orden)
    nombres = _Nombres(out.stem, split, hojas=False)
    usados: set[str] = set()
    resultados: list[PartResult] = []
    acortados = False
    reporte = ThrottledProgress(progress)
    gs_mb = lim.gs_max_import_mb if d.spec.destination == "google_sheets" else None
    for n, (parte, valor, sub) in enumerate(orden, start=1):
        if cancel is not None:
            cancel.check()
        plan_nombre = nombres.nombre(n, total, valor, sub)
        archivo = _nombre_de_archivo(d.carpeta, plan_nombre, ext, n, usados)
        acortados = acortados or archivo != plan_nombre + ext
        destino = d.carpeta / archivo
        d.creados.append(destino)
        if d.spec.format == "xlsx":
            if total > 1:
                reporte.force(ProgressEvent(f"Guardando la parte {n} de {total}…", n - 1, total))
            df = read_csv_exact(parte.tmp, CsvOptions(*_estilo_temporal()))
            res = write_table(
                df,
                destino,
                WriteOptions(
                    format="xlsx",
                    sheet_name=d.spec.sheet_name,
                    column_types=d.spec.column_types,
                ),
                cancel=cancel,
            )
            del df
            _agregar_avisos(warnings, res.warnings)
            parte.tmp.unlink(missing_ok=True)
            hoja = nombre_de_hoja(d.spec.sheet_name)
        else:
            try:
                os.replace(parte.tmp, destino)
            except OSError as exc:
                raise ExcelToolError(
                    "No se pudo guardar una de las partes. Revisá el nombre y los permisos.",
                    str(exc),
                ) from exc
            hoja = None
        if gs_mb is not None and destino.stat().st_size > gs_mb * _MB:
            peso = f"{destino.stat().st_size / _MB:,.0f}".replace(",", ".")
            warnings.append(
                f"La parte {n} ({archivo}) pesa {peso} MB: Google Sheets puede no importarla."
            )
        resultados.append(PartResult(destino, hoja, parte.rows, valor, parte.first, parte.last))
    if acortados:
        warnings.append(
            "Algunos nombres de archivo se acortaron porque la ruta completa pasaba "
            f"los {_MAX_RUTA} caracteres (Windows no los admite)."
        )
    return resultados


def _estilo_temporal() -> tuple[str, str]:
    from excel_tool.core.io.csv import CSV_ESTILOS

    return CSV_ESTILOS[_CSV_TEMPORAL]
