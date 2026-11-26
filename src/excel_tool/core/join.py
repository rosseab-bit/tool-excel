"""Unir un archivo principal con N archivos secundarios. Contrato: docs/contratos.md § Fase 4.

Cómo funciona (para quien mantenga esto):

- Las claves se normalizan de a una columna con `normalize.normalize_keys` y se convierten en
  códigos enteros (diccionario de pyarrow). Una clave vacía queda en -1 y NO empareja.
- Cada secundario se reduce a una "tabla" con sus códigos de clave:
    * detalle: todas sus filas;
    * resumen: el resultado de `aggregate.summarize` (agrupando por la clave NORMALIZADA, así
      " 12 " y "12" caen en el mismo grupo) más las columnas extra elegidas;
    * primera fila: una fila por clave.
- La unión no copia columnas hasta el final: se arman vectores de posiciones (fila del principal
  y fila de cada tabla, -1 = sin pareja) y recién al terminar se toman las columnas que salen.
- El diagnóstico cuenta filas con los mismos vectores de tamaños por clave, sin unir.

Decisiones de nivel 1 (detalle interno del core):
- `SourceSpec.columns` se toma literal: salen esas columnas (más las claves del principal, que
  salen siempre una sola vez). Una lista vacía en un secundario no trae columnas.
- `prefix` se antepone a TODAS las columnas de ese secundario; si igual coincide con otra, se
  agrega " (<archivo sin extensión>)".
- Las claves que se muestran como ejemplo son los valores originales, recortados.
- Control de filas: aplica cuando el plan no puede multiplicar ni quitar filas (todos los
  secundarios en `left`, sin detalle, resúmenes solo por la clave) y no hay resumen final.
  Un resumen por cliente y artículo, o una unión `inner`, cambian las filas a propósito.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Literal

import numpy as np
import pandas as pd
import pyarrow as pa
import pyarrow.compute as pc

from excel_tool.core import limits, memory
from excel_tool.core.aggregate import (
    Aggregation,
    AggregationSpec,
    GroupBy,
    result_columns,
    summarize,
)
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io import ReadOptions
from excel_tool.core.limits import SizeEstimate
from excel_tool.core.memory import MemoryCheck
from excel_tool.core.normalize import KeyOptions, normalize_keys
from excel_tool.core.output import OutputSpec
from excel_tool.core.progress import CancelToken, ProgressCallback, ProgressEvent
from excel_tool.core.types import ColumnType, detect_column_type

log = logging.getLogger(__name__)

_STR = pd.StringDtype("pyarrow")
_SEP_CLAVE = "\x1f"  # separa las partes de una clave de varias columnas
_MUESTRA_CLAVES = 10  # claves de ejemplo en el diagnóstico
_FILAS_MUESTRA = 2_000  # filas del principal para estimar el tamaño en disco
_TOPE_FILAS_SECUNDARIO_MUESTRA = 200_000
_BYTES_POSICION = 8  # int64 por cada vector de posiciones


# ---------------------------------------------------------------------- contrato


@dataclass(frozen=True)
class SourceSpec:
    name: str
    path: Path
    read_options: ReadOptions
    columns: list[str]  # columnas a conservar


@dataclass(frozen=True)
class SecondarySpec(SourceSpec):
    key_pairs: list[tuple[str, str]]  # (columna del principal, columna de este archivo)
    how: Literal["left", "inner"] = "left"
    mode: Literal["detail", "summary", "first"] = "detail"
    summary: AggregationSpec | None = None  # mode == "summary"; la clave va siempre en el grupo
    prefix: str | None = None


@dataclass(frozen=True)
class JoinPlan:
    principal: SourceSpec
    key_options: KeyOptions
    secondaries: list[SecondarySpec]
    output_column_order: list[str] | None = None
    final_summary: AggregationSpec | None = None
    output: OutputSpec | None = None


@dataclass(frozen=True)
class SecondaryDiagnosis:
    name: str
    keys_principal: int
    keys_matched: int
    keys_unmatched_sample: list[str]
    secondary_orphans: int  # claves del secundario que no están en el principal
    secondary_orphans_sample: list[str]
    keys_with_many_rows: int
    example_many: tuple[str, int] | None
    rows_after: int  # filas del resultado tras unir este archivo (exacto, sin unir)
    keys_matched_without_zeros: int | None = None  # solo si sacar los ceros suma coincidencias


@dataclass(frozen=True)
class JoinDiagnosis:
    secondaries: list[SecondaryDiagnosis]
    rows_result: int
    columns_result: int
    size: SizeEstimate
    memory: MemoryCheck
    many_details: bool
    warnings: list[str]


@dataclass(frozen=True)
class JoinResult:
    df: pd.DataFrame
    column_types: dict[str, ColumnType]
    unmatched: dict[str, pd.DataFrame]
    diagnosis: JoinDiagnosis
    warnings: list[str]


# ---------------------------------------------------------------------- utilidades


def _n(x: int) -> str:
    """Entero con punto de miles."""
    return f"{x:,}".replace(",", ".")


def _arrow(s: pd.Series) -> pa.ChunkedArray:
    arr = s.array
    pa_arr = getattr(arr, "_pa_array", None)
    if isinstance(pa_arr, pa.ChunkedArray):
        if pa.types.is_string(pa_arr.type):
            return pa_arr
        if pa.types.is_large_string(pa_arr.type):
            return pa_arr.cast(pa.string())
    return pa.chunked_array([pa.array(s.tolist(), type=pa.string(), from_pandas=True)])


def _plano(x: pa.Array | pa.ChunkedArray) -> pa.Array:
    if isinstance(x, pa.ChunkedArray):
        if x.num_chunks == 1:
            return x.chunk(0)
        return pa.concat_arrays(x.chunks) if x.num_chunks else pa.array([], type=x.type)
    return x


def _como_serie(x: pa.Array | pa.ChunkedArray, index=None) -> pd.Series:
    if isinstance(x, pa.Array):
        x = pa.chunked_array([x])
    return pd.Series(pd.arrays.ArrowStringArray(x), index=index, dtype=_STR)


def _clave(
    df: pd.DataFrame, columnas: list[str], opciones: KeyOptions
) -> tuple[pa.Array, np.ndarray]:
    """Clave normalizada (una sola columna de texto) y máscara de claves vacías.

    Normaliza de a una columna para no gastar memoria. Con varias columnas, la clave es vacía si
    cualquiera de las partes está vacía. Las claves vacías quedan como nulos.
    """
    partes: list[pa.Array] = []
    vacio: np.ndarray | None = None
    for c in columnas:
        arr = _plano(_arrow(normalize_keys(df[c], opciones)))
        v = pc.equal(arr, "").fill_null(True).to_numpy(zero_copy_only=False).astype(bool)
        vacio = v if vacio is None else (vacio | v)
        partes.append(arr)
    comb = (
        partes[0]
        if len(partes) == 1
        else _plano(pc.binary_join_element_wise(*partes, pa.scalar(_SEP_CLAVE)))
    )
    assert vacio is not None
    if vacio.any():
        comb = _plano(pc.if_else(pa.array(vacio), pa.scalar(None, pa.string()), comb))
    return comb, vacio


def _codificar(clave: pa.Array) -> tuple[np.ndarray, pa.Array]:
    """(código de cada fila, -1 si es vacía; diccionario de claves en orden de aparición)."""
    d = clave.dictionary_encode()
    codigos = d.indices.fill_null(-1).to_numpy(zero_copy_only=False).astype(np.int64, copy=False)
    return codigos, d.dictionary


def _buscar_en(valores: pa.Array, conjunto: pa.Array) -> np.ndarray:
    """Posición de cada valor dentro de `conjunto` (-1 si no está)."""
    if len(valores) == 0:
        return np.zeros(0, dtype=np.int64)
    if len(conjunto) == 0:
        return np.full(len(valores), -1, dtype=np.int64)
    pos = pc.index_in(valores, value_set=conjunto)
    return pos.fill_null(-1).to_numpy(zero_copy_only=False).astype(np.int64, copy=False)


def _primeras(codigos: np.ndarray, n: int) -> np.ndarray:
    """Posición de la primera fila de cada código."""
    primera = np.full(n, -1, dtype=np.int64)
    validos = np.flatnonzero(codigos >= 0)
    primera[codigos[validos][::-1]] = validos[::-1]
    return primera


def _mostrar(df: pd.DataFrame, columnas: list[str], fila: int) -> str:
    return " / ".join(str(df[c].iloc[fila]).strip() for c in columnas)


def _tomar(serie: pd.Series, pos: np.ndarray | None) -> pd.Series:
    """Toma las filas `pos` (-1 = sin pareja, queda vacío). `pos=None` = la columna tal cual."""
    if pos is None:
        return serie.reset_index(drop=True)
    arr = _arrow(serie)
    if len(arr) == 0:
        return _como_serie(pa.array([""] * len(pos), type=pa.string()))
    faltan = pos < 0
    if faltan.any():
        idx = pa.array(np.where(faltan, 0, pos), mask=faltan)
        return _como_serie(pc.fill_null(arr.take(idx), ""))
    return _como_serie(arr.take(pa.array(pos)))


def _columna_de(df: pd.DataFrame, nombre: str, archivo: str) -> pd.Series:
    posiciones = np.flatnonzero(df.columns == nombre)
    if len(posiciones) == 0:
        raise ExcelToolError(f"No encuentro la columna «{nombre}» en «{archivo}».")
    return df.iloc[:, int(posiciones[0])]


# ---------------------------------------------------------------------- nombres


@dataclass
class _Nombres:
    """Qué columnas salen de cada archivo y cómo se llaman en el resultado."""

    principal: list[tuple[str, str]]  # (columna de origen, nombre final)
    secundarios: list[list[tuple[str, str]]]
    aux: str  # nombre de la columna auxiliar de clave en los resúmenes
    specs: list[AggregationSpec | None]  # resumen del secundario con la clave auxiliar


def _clave_aux(df: pd.DataFrame) -> str:
    nombre = "__clave_unida__"
    while nombre in set(map(str, df.columns)):
        nombre = "_" + nombre
    return nombre


def _spec_con_clave(spec: AggregationSpec, claves: list[str], aux: str) -> AggregationSpec:
    """Cambia las columnas clave del grupo por una sola columna auxiliar (la clave normalizada)."""
    grupos: list[GroupBy] = []
    puesta = False
    for g in spec.group_by:
        if g.period is None and g.column in claves:
            if not puesta:
                grupos.append(GroupBy(aux))
                puesta = True
        else:
            grupos.append(g)
    if not puesta:
        grupos.insert(0, GroupBy(aux))
    return AggregationSpec(grupos, list(spec.aggregations))


def _armar_nombres(plan: JoinPlan, tables: dict[str, pd.DataFrame]) -> _Nombres:
    principal = plan.principal
    dfp = tables[principal.name]
    claves_p = [p for s in plan.secondaries for p, _ in s.key_pairs]
    for c in claves_p + list(principal.columns):
        _columna_de(dfp, c, principal.name)
    pedidas = set(principal.columns) | set(claves_p)
    cols_p = [str(c) for c in dfp.columns if str(c) in pedidas]
    usados = set(cols_p)
    nombres_p = [(c, c) for c in cols_p]
    aux = _clave_aux(dfp)
    for s in plan.secondaries:
        aux = max(aux, _clave_aux(tables[s.name]), key=len)

    secundarios: list[list[tuple[str, str]]] = []
    specs: list[AggregationSpec | None] = []
    for s in plan.secondaries:
        dfs = tables[s.name]
        claves_s = [q for _, q in s.key_pairs]
        for c in claves_s + list(s.columns):
            _columna_de(dfs, c, s.name)
        if s.mode == "summary":
            if s.summary is None:
                raise ExcelToolError(
                    f"Falta definir el resumen de «{s.name}» (o elegí traer el detalle)."
                )
            spec = _spec_con_clave(s.summary, claves_s, aux)
            for g in spec.group_by:
                if g.column != aux:
                    _columna_de(dfs, g.column, s.name)
            for a in spec.aggregations:
                if a.func != "count":
                    _columna_de(dfs, a.column, s.name)
            origen = [c for c in result_columns(spec) if c != aux]
            specs.append(spec)
        else:
            vistas: set[str] = set()
            origen = []
            for c in s.columns:
                if c not in claves_s and c not in vistas:
                    vistas.add(c)
                    origen.append(c)
            specs.append(None)
        stem = Path(s.name).stem
        lista = []
        for c in origen:
            base = f"{s.prefix}{c}" if s.prefix else c
            final = base if base not in usados else f"{base} ({stem})"
            k = 2
            while final in usados:
                final = f"{base} ({stem}) ({k})"
                k += 1
            usados.add(final)
            lista.append((c, final))
        secundarios.append(lista)
    return _Nombres(nombres_p, secundarios, aux, specs)


def _ordenar(nombres: list[str], orden: list[str] | None) -> list[str]:
    """Las columnas de `orden` primero (las que existan), el resto detrás en su orden."""
    if not orden:
        return list(nombres)
    existentes = set(nombres)
    elegidas: list[str] = []
    for c in orden:
        if c in existentes and c not in elegidas:
            elegidas.append(c)
    return elegidas + [c for c in nombres if c not in set(elegidas)]


def _validar(plan: JoinPlan, tables: dict[str, pd.DataFrame]) -> None:
    nombres = [plan.principal.name] + [s.name for s in plan.secondaries]
    if len(set(nombres)) != len(nombres):
        raise ExcelToolError("Dos de los archivos tienen el mismo nombre. Renombrá uno.")
    for n in nombres:
        if n not in tables:
            raise ExcelToolError(f"Falta cargar el archivo «{n}».")
    for s in plan.secondaries:
        if not s.key_pairs:
            raise ExcelToolError(f"Elegí por qué columna se une «{s.name}» con el archivo base.")
        if s.mode not in ("detail", "summary", "first"):
            raise ExcelToolError(f"Elegí cómo traer los datos de «{s.name}».")
        if s.how not in ("left", "inner"):
            raise ExcelToolError(f"Tipo de unión desconocido para «{s.name}».")


# ---------------------------------------------------------------------- cruce de claves


@dataclass
class _Cruce:
    """Resultado de emparejar las claves del principal con las de un secundario."""

    spec: SecondarySpec
    df: pd.DataFrame
    codigos_s: np.ndarray  # código de cada fila del secundario (-1 = clave vacía)
    dic_s: pa.Array
    pcodigo: np.ndarray  # código (en dic_s) de cada fila del principal; -1 = sin pareja
    sizes: np.ndarray  # filas del secundario por código
    vacio_s: np.ndarray
    cod_p: np.ndarray
    dic_p: pa.Array
    mapa_p: np.ndarray  # clave distinta del principal -> código en dic_s (o -1)
    aux: pa.Array  # clave normalizada del secundario (nulos donde está vacía)


def _cruzar(
    plan: JoinPlan, spec: SecondarySpec, dfp: pd.DataFrame, dfs: pd.DataFrame, opciones: KeyOptions
) -> _Cruce:
    cols_p = [p for p, _ in spec.key_pairs]
    cols_s = [q for _, q in spec.key_pairs]
    clave_p, _ = _clave(dfp, cols_p, opciones)
    cod_p, dic_p = _codificar(clave_p)
    del clave_p
    clave_s, vacio_s = _clave(dfs, cols_s, opciones)
    codigos_s, dic_s = _codificar(clave_s)
    mapa_p = _buscar_en(dic_p, dic_s)
    pcodigo = (
        np.where(cod_p >= 0, mapa_p[np.maximum(cod_p, 0)], -1)
        if len(mapa_p)
        else np.full(len(cod_p), -1, dtype=np.int64)
    )
    sizes = np.bincount(codigos_s[codigos_s >= 0], minlength=len(dic_s)).astype(np.int64)
    return _Cruce(
        spec, dfs, codigos_s, dic_s, pcodigo.astype(np.int64, copy=False), sizes, vacio_s,
        cod_p, dic_p, mapa_p, clave_s,
    )  # fmt: skip


def _coincidencias(
    dfp: pd.DataFrame, dfs: pd.DataFrame, spec: SecondarySpec, opc: KeyOptions
) -> int:
    """Cuántas claves distintas del principal tienen pareja con las opciones dadas."""
    clave_p, _ = _clave(dfp, [p for p, _ in spec.key_pairs], opc)
    _, dic_p = _codificar(clave_p)
    clave_s, _ = _clave(dfs, [q for _, q in spec.key_pairs], opc)
    _, dic_s = _codificar(clave_s)
    return int((_buscar_en(dic_p, dic_s) >= 0).sum())


# ---------------------------------------------------------------------- tablas


@dataclass
class _Tabla:
    """Lo que aporta un secundario a la unión: filas con su código de clave."""

    frame: pd.DataFrame | None
    cols: list[str]  # columnas de `frame` que salen
    codigos: np.ndarray  # código de cada fila de la tabla
    sizes: np.ndarray  # filas de la tabla por código de clave
    filas: np.ndarray | None = None  # posición en `frame` de cada fila de la tabla (primera fila)
    tipos: dict[str, ColumnType] = field(default_factory=dict)
    avisos: list[str] = field(default_factory=list)


def _tabla(
    c: _Cruce, nombres: _Nombres, k: int, completo: bool, cancel: CancelToken | None
) -> _Tabla:
    spec = c.spec
    n_codigos = len(c.dic_s)
    if spec.mode == "detail":
        return _Tabla(c.df, [o for o, _ in nombres.secundarios[k]], c.codigos_s, c.sizes)
    if spec.mode == "first":
        _, primeras = np.unique(c.codigos_s, return_index=True)
        if len(primeras) and c.codigos_s[primeras[0]] < 0:
            primeras = primeras[1:]
        primeras = np.sort(primeras)
        return _Tabla(
            c.df,
            [o for o, _ in nombres.secundarios[k]],
            c.codigos_s[primeras],
            np.ones(n_codigos, dtype=np.int64),
            primeras,
        )
    # resumen
    spec_aux = nombres.specs[k]
    assert spec_aux is not None
    aux = nombres.aux
    solo_clave = len(spec_aux.group_by) == 1
    if not completo and solo_clave:
        return _Tabla(None, [], np.arange(n_codigos, dtype=np.int64), np.ones(n_codigos, np.int64))
    usadas = [g.column for g in spec_aux.group_by if g.column != aux]
    usadas += [a.column for a in spec_aux.aggregations if a.func != "count"]
    sub = pd.DataFrame(
        {col: _columna_de(c.df, col, spec.name) for col in dict.fromkeys(usadas)},
        index=c.df.index,
    )
    sub[aux] = _como_serie(c.aux, index=c.df.index)
    if c.vacio_s.any():
        sub = sub.loc[~c.vacio_s]
    corrida = (
        spec_aux if completo else AggregationSpec(spec_aux.group_by, [Aggregation("", "count")])
    )
    res = summarize(sub, corrida, cancel=cancel)
    del sub
    codigos = _buscar_en(_plano(_arrow(res.df[aux])), c.dic_s)
    sizes = np.bincount(codigos[codigos >= 0], minlength=n_codigos).astype(np.int64)
    if completo:
        cols = [o for o, _ in nombres.secundarios[k]]
        return _Tabla(res.df, cols, codigos, sizes, None, dict(res.column_types), res.warnings)
    todas = result_columns(spec_aux)
    cols = [x for x in todas[: len(spec_aux.group_by)] if x != aux]
    return _Tabla(res.df, cols, codigos, sizes)


# ---------------------------------------------------------------------- unión por posiciones


def _factor(sizes_por_fila: np.ndarray, how: str) -> np.ndarray:
    """Filas que genera cada fila actual: coincidencias; con `left`, al menos 1."""
    return np.maximum(sizes_por_fila, 1) if how == "left" else sizes_por_fila


def _coincidencias_por_fila(pcodigo: np.ndarray, sizes: np.ndarray) -> np.ndarray:
    if len(sizes) == 0:
        return np.zeros(len(pcodigo), dtype=np.int64)
    return np.where(pcodigo >= 0, sizes[np.maximum(pcodigo, 0)], 0)


@dataclass
class _Posiciones:
    """Para cada fila del resultado: fila del principal y fila de cada tabla (-1 = sin pareja)."""

    pidx: np.ndarray | None  # None = identidad (todas las filas del principal, en orden)
    sidx: list[np.ndarray]
    n: int


def _elegir(
    orden: np.ndarray,
    inicios: np.ndarray,
    codigo: np.ndarray,
    dentro: np.ndarray | int,
    m: np.ndarray,
) -> np.ndarray:
    """Fila de la tabla que le toca a cada fila del resultado (-1 = sin pareja)."""
    if len(orden) == 0 or len(inicios) == 0:
        return np.full(len(codigo), -1, dtype=np.int64)
    base = inicios[np.maximum(codigo, 0)] + dentro
    return np.where(m > 0, orden[np.clip(base, 0, len(orden) - 1)], -1).astype(np.int64)


def _unir(pos: _Posiciones, c: _Cruce, t: _Tabla, cancel: CancelToken | None) -> _Posiciones:
    """Suma una tabla a las posiciones ya armadas."""
    actual = c.pcodigo if pos.pidx is None else c.pcodigo[pos.pidx]
    m = _coincidencias_por_fila(actual, t.sizes)
    reps = _factor(m, c.spec.how)
    orden = np.argsort(t.codigos, kind="stable")
    negativos = int((t.codigos < 0).sum())
    inicios = negativos + np.cumsum(t.sizes) - t.sizes
    if cancel:
        cancel.check()
    if len(reps) == 0 or (reps == 1).all():
        # Una sola pareja (o ninguna) por fila: no se repite ni se quita ninguna fila.
        nuevo = _elegir(orden, inicios, actual, 0, m)
        return _Posiciones(pos.pidx, [*pos.sidx, nuevo], pos.n)
    total = int(reps.sum())
    origen = np.repeat(np.arange(pos.n, dtype=np.int64), reps)
    primero = np.cumsum(reps) - reps
    dentro = np.arange(total, dtype=np.int64) - np.repeat(primero, reps)
    nuevo = _elegir(orden, inicios, actual[origen], dentro, m[origen])
    base_p = np.arange(pos.n, dtype=np.int64) if pos.pidx is None else pos.pidx
    return _Posiciones(base_p[origen], [*(s[origen] for s in pos.sidx), nuevo], total)


def _posiciones_tabla(t: _Tabla, sidx: np.ndarray) -> np.ndarray:
    """Pasa de fila de la tabla a fila del `frame` de origen (solo cambia con 'primera fila')."""
    if t.filas is None:
        return sidx
    return np.where(sidx >= 0, t.filas[np.maximum(sidx, 0)], -1)


# ---------------------------------------------------------------------- diagnóstico


@dataclass
class _Analisis:
    plan: JoinPlan
    tables: dict[str, pd.DataFrame]
    nombres: _Nombres
    cruces: list[_Cruce]
    por_archivo: list[SecondaryDiagnosis]
    rows_join: int  # filas tras unir todos (antes del resumen final)
    avisos: list[str]


class _Avance:
    def __init__(self, progress: ProgressCallback | None, cancel: CancelToken | None, total: int):
        self._progress = progress
        self._cancel = cancel
        self._total = total
        self._hechos = 0

    def paso(self, mensaje: str) -> None:
        if self._cancel:
            self._cancel.check()
        if self._progress:
            self._progress(ProgressEvent(mensaje, self._hechos, self._total))
        self._hechos += 1

    def fin(self, mensaje: str) -> None:
        if self._progress:
            self._progress(ProgressEvent(mensaje, self._total, self._total))


def _analizar(
    plan: JoinPlan,
    tables: dict[str, pd.DataFrame],
    avance: _Avance,
    cancel: CancelToken | None,
    *,
    con_ceros: bool,
) -> _Analisis:
    """Cruza las claves de cada secundario y cuenta filas, sin unir."""
    _validar(plan, tables)
    nombres = _armar_nombres(plan, tables)
    dfp = tables[plan.principal.name]
    opciones = plan.key_options
    filas_por_principal = np.ones(len(dfp), dtype=np.float64)
    cruces: list[_Cruce] = []
    por_archivo: list[SecondaryDiagnosis] = []
    avisos: list[str] = []
    detalles = [s.name for s in plan.secondaries if s.mode == "detail"]
    if len(detalles) > 1:
        avisos.append(
            f"Pediste el detalle de {len(detalles)} archivos a la vez "
            f"({', '.join('«' + d + '»' for d in detalles)}). Si traés el detalle de dos "
            "archivos a la vez, las filas se combinan todas con todas y los totales quedan "
            "inflados. Te recomendamos ver el detalle de uno solo y el resumen de los otros."
        )
    for k, spec in enumerate(plan.secondaries):
        avance.paso(f"Comparando las claves de «{spec.name}»")
        dfs = tables[spec.name]
        c = _cruzar(plan, spec, dfp, dfs, opciones)
        avance.paso(f"Contando las filas de «{spec.name}»")
        t = _tabla(c, nombres, k, completo=False, cancel=cancel)
        m = _coincidencias_por_fila(c.pcodigo, t.sizes)
        filas_por_principal = filas_por_principal * _factor(m, spec.how).astype(np.float64)
        rows_after = int(filas_por_principal.sum())
        d, extra = _diagnosticar(c, dfp, dfs, rows_after, opciones, con_ceros)
        por_archivo.append(d)
        avisos.extend(extra)
        del t
        cruces.append(c)
    rows_join = int(filas_por_principal.sum())
    return _Analisis(plan, tables, nombres, cruces, por_archivo, rows_join, avisos)


def _diagnosticar(
    c: _Cruce,
    dfp: pd.DataFrame,
    dfs: pd.DataFrame,
    rows_after: int,
    opciones: KeyOptions,
    con_ceros: bool,
) -> tuple[SecondaryDiagnosis, list[str]]:
    spec = c.spec
    cols_p = [p for p, _ in spec.key_pairs]
    cols_s = [q for _, q in spec.key_pairs]
    avisos: list[str] = []
    keys_principal = len(c.dic_p)
    encontrada = c.mapa_p >= 0
    keys_matched = int(encontrada.sum())

    primeras_p = _primeras(c.cod_p, keys_principal)
    sin = np.flatnonzero(~encontrada)[:_MUESTRA_CLAVES]
    muestra_sin = [_mostrar(dfp, cols_p, int(primeras_p[i])) for i in sin]

    n_s = len(c.dic_s)
    cubierta = np.zeros(n_s, dtype=bool)
    cubierta[c.mapa_p[encontrada]] = True
    primeras_s = _primeras(c.codigos_s, n_s)
    huerfanas = np.flatnonzero(~cubierta)
    muestra_h = [_mostrar(dfs, cols_s, int(primeras_s[i])) for i in huerfanas[:_MUESTRA_CLAVES]]

    # (solo claves que tienen pareja: son las que repiten filas en el resultado)
    varias = np.flatnonzero((c.sizes > 1) & cubierta)
    ejemplo = None
    if len(varias):
        mejor = int(varias[np.argmax(c.sizes[varias])])
        ejemplo = (_mostrar(dfs, cols_s, int(primeras_s[mejor])), int(c.sizes[mejor]))

    # claves repetidas en el principal (solo importan las que tienen pareja)
    reps_p = np.bincount(c.cod_p[c.cod_p >= 0], minlength=keys_principal)
    repetidas = np.flatnonzero((reps_p > 1) & encontrada)
    if len(repetidas):
        i = int(repetidas[0])
        efecto = (
            "las filas del resultado se multiplican"
            if spec.mode == "detail"
            else "los valores se repiten en cada una y los totales pueden contar de más"
        )
        avisos.append(
            f"En el archivo base hay {_n(len(repetidas))} clave(s) repetida(s) que también "
            f"están en «{spec.name}» (por ejemplo «{_mostrar(dfp, cols_p, int(primeras_p[i]))}» "
            f"en {_n(int(reps_p[i]))} filas): {efecto}."
        )

    sin_ceros: int | None = None
    if con_ceros and not opciones.strip_leading_zeros and keys_principal:
        hay_digitos = bool(pc.any(pc.match_substring_regex(c.dic_p, r"^[0-9]+$")).as_py())
        if hay_digitos:
            alt = _coincidencias(dfp, dfs, spec, replace(opciones, strip_leading_zeros=True))
            if alt > keys_matched:
                sin_ceros = alt
                avisos.append(
                    f"En «{spec.name}»: si sacás los ceros a la izquierda, coinciden "
                    f"{_n(alt)} claves en vez de {_n(keys_matched)}."
                )

    return (
        SecondaryDiagnosis(
            name=spec.name,
            keys_principal=keys_principal,
            keys_matched=keys_matched,
            keys_unmatched_sample=muestra_sin,
            secondary_orphans=len(huerfanas),
            secondary_orphans_sample=muestra_h,
            keys_with_many_rows=len(varias),
            example_many=ejemplo,
            rows_after=rows_after,
            keys_matched_without_zeros=sin_ceros,
        ),
        avisos,
    )


def _nombres_finales(a: _Analisis) -> list[str]:
    plan = a.plan
    unidas = [n for _, n in a.nombres.principal] + [
        n for lista in a.nombres.secundarios for _, n in lista
    ]
    if plan.final_summary is not None:
        unidas = result_columns(plan.final_summary)
    return _ordenar(unidas, plan.output_column_order)


def _cerrar_diagnostico(
    a: _Analisis, avance: _Avance | None, cancel: CancelToken | None
) -> JoinDiagnosis:
    plan = a.plan
    columnas_join = len(a.nombres.principal) + sum(len(x) for x in a.nombres.secundarios)
    rows_result = a.rows_join
    columnas = len(_nombres_finales(a))
    if plan.final_summary is not None:
        if avance:
            avance.paso("Calculando cuántas filas tiene el resumen final")
        rows_result = _filas_resumen_final(a, cancel)
    n_sec = len(a.cruces) if a.cruces else len(plan.secondaries)
    extra = a.rows_join * (1 + n_sec) * _BYTES_POSICION * 3
    mem = memory.check_result_memory(a.rows_join, max(columnas_join, 1), extra_bytes=extra)
    if avance:
        avance.paso("Estimando el tamaño del resultado")
    size = _estimar_tamano(a, rows_result, columnas)
    many = sum(1 for s in plan.secondaries if s.mode == "detail") > 1
    return JoinDiagnosis(
        secondaries=a.por_archivo,
        rows_result=rows_result,
        columns_result=columnas,
        size=size,
        memory=mem,
        many_details=many,
        warnings=list(a.avisos),
    )


def _filas_resumen_final(a: _Analisis, cancel: CancelToken | None) -> int:
    """Filas del resumen final: se une solo con las columnas por las que agrupa y se cuentan."""
    final = a.plan.final_summary
    assert final is not None
    if not a.rows_join:
        return 0
    grupos = [g.column for g in final.group_by]
    if not grupos:
        return 1
    cruces = a.cruces or _recruzar(a)
    tablas = [
        _tabla(c, a.nombres, k, completo=_necesita_completo(a, k, grupos), cancel=cancel)
        for k, c in enumerate(cruces)
    ]
    pos = _Posiciones(None, [], len(a.tables[a.plan.principal.name]))
    for c, t in zip(cruces, tablas, strict=True):
        pos = _unir(pos, c, t, cancel)
    df = _juntar(a, tablas, pos, set(grupos))
    cuenta = summarize(
        df, AggregationSpec(final.group_by, [Aggregation("", "count")]), cancel=cancel
    )
    return len(cuenta.df)


def _necesita_completo(a: _Analisis, k: int, grupos: list[str]) -> bool:
    spec_aux = a.nombres.specs[k]
    if spec_aux is None:
        return False
    esqueleto = set(result_columns(spec_aux)[: len(spec_aux.group_by)])
    propios = {n for o, n in a.nombres.secundarios[k] if o not in esqueleto}
    return any(g in propios for g in grupos)


def _recruzar(a: _Analisis) -> list[_Cruce]:
    dfp = a.tables[a.plan.principal.name]
    return [
        _cruzar(a.plan, s, dfp, a.tables[s.name], a.plan.key_options) for s in a.plan.secondaries
    ]


def _columnas_con_cero(spec: AggregationSpec | None) -> set[str]:
    """Columnas (nombre de origen) de un resumen que valen 0 en las filas sin pareja."""
    if spec is None:
        return set()
    nombres = result_columns(spec)[len(spec.group_by) :]
    return {
        n
        for n, ag in zip(nombres, spec.aggregations, strict=True)
        if ag.func in ("sum", "count", "count_distinct")
    }


def _juntar(
    a: _Analisis, tablas: list[_Tabla], pos: _Posiciones, solo: set[str] | None
) -> pd.DataFrame:
    """Toma las columnas del resultado (todas, o las de `solo`) con los vectores de posiciones."""
    dfp = a.tables[a.plan.principal.name]
    columnas: dict[str, pd.Series] = {}
    for origen, final in a.nombres.principal:
        if solo is None or final in solo:
            columnas[final] = _tomar(_columna_de(dfp, origen, a.plan.principal.name), pos.pidx)
    for k, (t, lista) in enumerate(zip(tablas, a.nombres.secundarios, strict=True)):
        usar = [(o, f) for o, f in lista if (solo is None or f in solo) and o in t.cols]
        if not usar:
            continue
        sidx = pos.sidx[k]
        assert t.frame is not None
        origen_pos = _posiciones_tabla(t, sidx)
        con_cero = _columnas_con_cero(a.nombres.specs[k])
        for o, f in usar:
            col = _tomar(_columna_de(t.frame, o, a.plan.secondaries[k].name), origen_pos)
            if o in con_cero and (origen_pos < 0).any():
                # ADR-037: sin pareja, contar y sumar dan 0 (promedio, mínimo, etc. quedan vacíos)
                col = col.where(origen_pos >= 0, "0")
            columnas[f] = col
    return pd.DataFrame(columnas, copy=False)


# ---------------------------------------------------------------------- tamaño (muestra)


def _estimar_tamano(a: _Analisis, rows_result: int, columnas: int) -> SizeEstimate:
    """Tamaño aproximado: filas y columnas exactas; bytes por celda de una muestra unida."""
    celdas = (rows_result + 1) * columnas
    if rows_result == 0 or columnas == 0:
        return SizeEstimate(rows_result, columnas, celdas, 0.0, 0.0)
    try:
        muestra = _unir_muestra(a)
    except ExcelToolError:
        raise
    except Exception:  # la muestra es solo una aproximación: no tiene que romper el diagnóstico
        log.exception("No se pudo armar la muestra para estimar el tamaño")
        muestra = None
    if muestra is None or muestra.empty or muestra.shape[1] == 0:
        return SizeEstimate(rows_result, columnas, celdas, 0.0, 0.0)
    est = limits.estimate(muestra)
    escala = celdas / max(est.cells, 1)
    return SizeEstimate(
        rows=rows_result,
        columns=columnas,
        cells=celdas,
        approx_mb_xlsx=est.approx_mb_xlsx * escala,
        approx_mb_csv=est.approx_mb_csv * escala,
    )


def _unir_muestra(a: _Analisis) -> pd.DataFrame | None:
    """Une de verdad un subconjunto chico (filas repartidas del principal y lo que les
    corresponde en cada secundario) para medir cuánto ocupa una celda."""
    plan = a.plan
    dfp = a.tables[plan.principal.name]
    if not len(dfp):
        return None
    if len(dfp) > _FILAS_MUESTRA:
        pos = np.linspace(0, len(dfp) - 1, _FILAS_MUESTRA).astype(np.int64)
    else:
        pos = np.arange(len(dfp), dtype=np.int64)
    cruces = a.cruces or _recruzar(a)
    tablas = dict(a.tables)
    tablas[plan.principal.name] = dfp.iloc[pos].reset_index(drop=True)
    for s, c in zip(plan.secondaries, cruces, strict=True):
        codigos = np.unique(c.pcodigo[pos])
        codigos = codigos[codigos >= 0]
        quedan = np.flatnonzero(np.isin(c.codigos_s, codigos))
        if len(quedan) > _TOPE_FILAS_SECUNDARIO_MUESTRA:
            quedan = quedan[:: -(-len(quedan) // _TOPE_FILAS_SECUNDARIO_MUESTRA)]
        tablas[s.name] = c.df.iloc[quedan].reset_index(drop=True)
    sub = _analizar(plan, tablas, _Avance(None, None, 1), None, con_ceros=False)
    return _construir(sub, None, None, con_tipos=False).df


# ---------------------------------------------------------------------- armado


@dataclass
class _Armado:
    df: pd.DataFrame
    tipos: dict[str, ColumnType]
    unmatched: dict[str, pd.DataFrame]
    avisos: list[str]


def _detectar_tipos(df: pd.DataFrame, columnas: list[tuple[str, str]], archivo: str) -> dict:
    tipos: dict[str, ColumnType] = {}
    for origen, final in columnas:
        serie = _columna_de(df, origen, archivo)
        tipos[final] = detect_column_type(serie).detected
    return tipos


def _sin_pareja(c: _Cruce) -> pd.DataFrame:
    """Filas del secundario cuya clave no está en el principal (o está vacía)."""
    cubierta = np.zeros(len(c.dic_s), dtype=bool)
    ok = c.mapa_p[c.mapa_p >= 0]
    cubierta[ok] = True
    huerfana = (c.codigos_s < 0) | ~cubierta[np.maximum(c.codigos_s, 0)]
    return c.df.loc[huerfana].reset_index(drop=True)


def _control_de_filas(plan: JoinPlan) -> bool:
    """¿Este plan tiene que dar exactamente las filas del principal?"""
    if plan.final_summary is not None:
        return False
    for s in plan.secondaries:
        if s.how != "left" or s.mode == "detail":
            return False
        if s.mode == "summary" and s.summary is not None:
            claves = {q for _, q in s.key_pairs}
            if any(g.column not in claves or g.period for g in s.summary.group_by):
                return False
    return True


def _construir(
    a: _Analisis, avance: _Avance | None, cancel: CancelToken | None, *, con_tipos: bool = True
) -> _Armado:
    plan = a.plan
    dfp = a.tables[plan.principal.name]
    cruces = a.cruces
    avisos: list[str] = []
    tablas: list[_Tabla] = []
    pos = _Posiciones(None, [], len(dfp))
    unmatched: dict[str, pd.DataFrame] = {}
    for k, c in enumerate(cruces):
        spec = c.spec
        if avance:
            avance.paso(f"Preparando «{spec.name}»")
        t = _tabla(c, a.nombres, k, completo=True, cancel=cancel)
        avisos.extend(t.avisos)
        tablas.append(t)
        unmatched[spec.name] = _sin_pareja(c)
        if avance:
            avance.paso(f"Uniendo «{spec.name}»")
        pos = _unir(pos, c, t, cancel)
    if pos.n != a.rows_join:
        raise ExcelToolError(
            "La cantidad de filas del resultado no coincide con la estimada. Avisá al soporte.",
            f"filas={pos.n} estimadas={a.rows_join}",
        )
    if _control_de_filas(plan) and pos.n != len(dfp):
        raise ExcelToolError(
            "El resultado no tiene las mismas filas que el archivo base y no tendría que "
            "pasar. No se guardó nada.",
            f"filas={pos.n} base={len(dfp)}",
        )
    if avance:
        avance.paso("Armando el resultado")
    df = _juntar(a, tablas, pos, None)
    tipos_resumen = [t.tipos for t in tablas]
    del tablas, pos

    tipos: dict[str, ColumnType] = {}
    if con_tipos:
        tipos = _detectar_tipos(dfp, a.nombres.principal, plan.principal.name)
        for k, c in enumerate(cruces):
            lista = a.nombres.secundarios[k]
            if plan.secondaries[k].mode == "summary":
                for origen, final in lista:
                    tipos[final] = tipos_resumen[k].get(origen, "texto")
            else:
                tipos.update(_detectar_tipos(c.df, lista, plan.secondaries[k].name))

    if plan.final_summary is not None:
        if avance:
            avance.paso("Resumiendo el resultado")
        res = summarize(df, plan.final_summary, column_types=tipos or None, cancel=cancel)
        df, tipos = res.df, dict(res.column_types)
        avisos.extend(res.warnings)
    orden = _nombres_finales(a)
    df = df[orden]
    tipos = {n: tipos[n] for n in orden if n in tipos}
    return _Armado(df.reset_index(drop=True), tipos, unmatched, avisos)


# ---------------------------------------------------------------------- API pública


def diagnose_join(
    plan: JoinPlan,
    tables: dict[str, pd.DataFrame],
    *,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> JoinDiagnosis:
    """Diagnóstico del cruce sin unir: coincidencias, filas exactas, tamaño y memoria."""
    avance = _Avance(progress, cancel, 2 * len(plan.secondaries) + 2)
    a = _analizar(plan, tables, avance, cancel, con_ceros=True)
    diag = _cerrar_diagnostico(a, avance, cancel)
    avance.fin("Diagnóstico listo")
    return diag


def run_join(
    plan: JoinPlan,
    tables: dict[str, pd.DataFrame],
    *,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> JoinResult:
    """Une los archivos según el plan. Antes de armar nada controla la memoria."""
    k = len(plan.secondaries)
    avance = _Avance(progress, cancel, 4 * k + 4)
    a = _analizar(plan, tables, avance, cancel, con_ceros=True)
    diag = _cerrar_diagnostico(a, avance, cancel)
    if not diag.memory.ok:
        raise ExcelToolError(diag.memory.message)
    armado = _construir(a, avance, cancel)
    if len(armado.df) != diag.rows_result:
        raise ExcelToolError(
            "La cantidad de filas del resultado no coincide con la estimada. No se guardó nada.",
            f"filas={len(armado.df)} estimadas={diag.rows_result}",
        )
    avance.fin("Unión lista")
    return JoinResult(
        df=armado.df,
        column_types=armado.tipos,
        unmatched=armado.unmatched,
        diagnosis=diag,
        warnings=[*diag.warnings, *armado.avisos],
    )
