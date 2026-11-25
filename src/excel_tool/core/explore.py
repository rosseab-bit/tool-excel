"""Valores distintos, filtros, orden y páginas. Contrato: docs/contratos.md § Fase 2 (ADR-024).

Todo trabaja con posiciones de fila (0 = primera fila de datos) y nunca copia la tabla: cada
columna se factoriza (códigos + valores únicos) y las cuentas, comparaciones y conversiones se
hacen sobre los valores únicos, que son pocos comparados con las filas.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np
import pandas as pd
import pyarrow.compute as pc

from excel_tool.core import types as tipos
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.progress import CancelToken, ProgressCallback, ProgressEvent, ThrottledProgress
from excel_tool.core.types import ColumnType

VACIO = "(vacío)"
COL_CANTIDAD = "Cantidad"
COL_PORCENTAJE = "Porcentaje"
COL_FILA = "N.º fila"
_STR = pd.StringDtype("pyarrow")

FilterOp = Literal["igual", "distinto", "contiene", "vacio", "no_vacio", "mayor", "menor"]


@dataclass(frozen=True)
class DistinctOptions:
    sort: Literal["cantidad", "alfabetico"] = "cantidad"
    include_empty: bool = True  # fila "(vacío)"
    normalize: bool = False  # ignorar mayúsculas y espacios de los bordes
    search: str | None = None  # valores que contienen el texto (sin importar mayúsculas)
    limit: int = 100_000  # máximo de filas devueltas


@dataclass
class DistinctResult:
    columns: list[str]
    table: pd.DataFrame  # columnas elegidas + "Cantidad" (int) + "Porcentaje" (float 0-100)
    distinct_total: int  # antes de search y limit
    rows_total: int
    truncated: bool
    variant_groups: list[list[str]]  # solo con normalize=False


@dataclass(frozen=True)
class FilterCondition:
    column: str
    op: FilterOp
    value: str = ""


# --------------------------------------------------------------------------- helpers


def _columna(df: pd.DataFrame, nombre: str) -> pd.Series:
    posiciones = np.flatnonzero(df.columns == nombre)
    if len(posiciones) == 0:
        raise ExcelToolError(f"No existe la columna «{nombre}».")
    return df.iloc[:, int(posiciones[0])]


def _texto(valores) -> pd.Series:
    """Serie de texto con almacenamiento pyarrow (compacta), índice 0..n-1."""
    return pd.Series(valores, dtype=_STR, copy=False).reset_index(drop=True)


def _factorizar(serie: pd.Series) -> tuple[np.ndarray, pd.Series]:
    """(códigos int64 sin negativos, valores únicos como Serie de texto pyarrow).

    Los nulos se tratan como "" y se juntan con los "" que ya había."""
    codigos, unicos = pd.factorize(serie, use_na_sentinel=True)
    valores = _texto(unicos.array)
    nulos = codigos < 0
    if nulos.any():
        donde = np.flatnonzero((valores == "").to_numpy(dtype=bool, na_value=False))
        if len(donde):
            destino = int(donde[0])
        else:
            destino = len(valores)
            valores = _texto(pd.concat([valores, _texto([""])], ignore_index=True))
        codigos[nulos] = destino
    return codigos.astype(np.int64, copy=False), valores


def _bool(serie: pd.Series) -> np.ndarray:
    return serie.to_numpy(dtype=bool, na_value=False)


def _normalizar(valores: pd.Series) -> pd.Series:
    return valores.str.strip().str.lower()


def _vacios(valores: pd.Series) -> np.ndarray:
    return _bool(valores.str.strip() == "")


def _etiqueta(valor: str) -> str:
    return VACIO if valor == "" else valor


# --------------------------------------------------------------------------- value_counts


def _agrupar_columna(
    codigos: np.ndarray, valores: pd.Series, normalize: bool
) -> tuple[np.ndarray, pd.Series, list[list[str]]]:
    """Códigos de grupo por fila, texto que muestra cada grupo y grupos de variantes.

    Sin normalizar, cada valor único es un grupo y se informan las variantes que se
    juntarían al normalizar. Normalizando, el grupo se muestra con su variante más
    frecuente (sin espacios en los bordes)."""
    normales = _normalizar(valores)
    grupo_de, _ = pd.factorize(normales)
    n_grupos = int(grupo_de.max()) + 1 if len(grupo_de) else 0
    if not normalize:
        variantes: list[list[str]] = []
        if n_grupos < len(valores):
            tam = np.bincount(grupo_de, minlength=n_grupos)
            multiples = np.flatnonzero(tam > 1)
            ordenados = np.argsort(grupo_de, kind="stable")
            inicio = np.concatenate(([0], np.cumsum(tam)))
            for g in multiples:
                miembros = ordenados[inicio[g] : inicio[g + 1]]
                variantes.append(valores.iloc[miembros].tolist())
        return codigos, valores, variantes
    cuenta = np.bincount(codigos, minlength=len(valores))
    # variante más frecuente de cada grupo (a igualdad, la primera que apareció)
    orden = np.lexsort((np.arange(len(valores)), -cuenta, grupo_de))
    primero = np.ones(len(orden), dtype=bool)
    primero[1:] = grupo_de[orden][1:] != grupo_de[orden][:-1]
    elegido = np.empty(n_grupos, dtype=np.int64)
    elegido[grupo_de[orden][primero]] = orden[primero]
    return grupo_de[codigos], valores.iloc[elegido].str.strip().reset_index(drop=True), []


def _rango_alfabetico(valores: pd.Series) -> np.ndarray:
    """Posición de cada valor en el orden alfabético sin importar mayúsculas; "" al final."""
    if len(valores) == 0:
        return np.empty(0, dtype=np.int64)
    bajos = valores.str.lower().array.__arrow_array__()
    rango = pc.rank(bajos, sort_keys="ascending", tiebreaker="dense").to_numpy().astype(np.int64)
    rango[_bool(valores == "")] = len(valores) + 1
    return rango


def value_counts(
    df: pd.DataFrame,
    columns: list[str],
    options: DistinctOptions = DistinctOptions(),
    *,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> DistinctResult:
    """Cuenta cuántas filas hay de cada valor (o combinación de valores) de las columnas."""
    if not columns:
        raise ExcelToolError("Elegí al menos una columna.")
    reporte = ThrottledProgress(progress)
    filas = len(df)
    total_pasos = len(columns) + 1
    por_columna: list[tuple[np.ndarray, pd.Series]] = []  # (códigos de grupo, textos de grupo)
    variantes: list[list[str]] = []
    for paso, nombre in enumerate(columns):
        if cancel is not None:
            cancel.check()
        reporte.force(ProgressEvent(f"Analizando la columna {nombre}…", paso, total_pasos))
        codigos, valores = _factorizar(_columna(df, nombre))
        cod_g, textos, var = _agrupar_columna(codigos, valores, options.normalize)
        por_columna.append((cod_g, textos))
        variantes.extend(var)
        del codigos

    if cancel is not None:
        cancel.check()
    reporte.force(ProgressEvent("Contando los valores…", len(columns), total_pasos))

    # Clave combinada: se re-factoriza cada vez que el producto podría desbordar.
    clave = por_columna[0][0]
    tam_clave = len(por_columna[0][1])
    if len(columns) > 1:
        for cod_g, textos in por_columna[1:]:
            if cancel is not None:
                cancel.check()
            if tam_clave * len(textos) >= 2**62:
                clave, uni = pd.factorize(clave)
                clave = clave.astype(np.int64)
                tam_clave = len(uni)
            clave = clave * len(textos) + cod_g
            tam_clave *= len(textos)
        clave, _ = pd.factorize(clave)
        clave = clave.astype(np.int64)
    # `clave` ya son ids consecutivos 0..G-1 (en una columna, los grupos tal cual)
    n_grupos = int(clave.max()) + 1 if len(clave) else 0
    cantidades = np.bincount(clave, minlength=n_grupos).astype(np.int64)
    # primera fila de cada grupo: de ahí se saca el texto de cada columna
    if n_grupos:
        # asignando al revés, el último valor escrito es la primera aparición
        primera = np.empty(n_grupos, dtype=np.int64)
        primera[clave[::-1]] = np.arange(filas - 1, -1, -1, dtype=np.int64)
    else:
        primera = np.empty(0, dtype=np.int64)
    del clave
    textos_grupo = [
        textos.iloc[cod_g[primera]].reset_index(drop=True) for cod_g, textos in por_columna
    ]
    del por_columna

    distinct_total = n_grupos
    mantener = np.ones(n_grupos, dtype=bool)
    if not options.include_empty:
        vacio_alguno = np.zeros(n_grupos, dtype=bool)
        for t in textos_grupo:
            vacio_alguno |= _bool(t == "")
        mantener &= ~vacio_alguno
    if options.search:
        buscado = options.search.lower()
        alguno = np.zeros(n_grupos, dtype=bool)
        for t in textos_grupo:
            if n_grupos:
                alguno |= _bool(t.str.lower().str.contains(buscado, regex=False))
        mantener &= alguno
    ids = np.flatnonzero(mantener)

    if cancel is not None:
        cancel.check()
    if options.sort == "alfabetico":
        claves_orden = [_rango_alfabetico(t)[ids] for t in textos_grupo]
        orden = np.lexsort(tuple(reversed(claves_orden))) if len(ids) else ids
    else:
        claves_orden = [_rango_alfabetico(t)[ids] for t in textos_grupo]
        orden = np.lexsort((*reversed(claves_orden), -cantidades[ids])) if len(ids) else ids
    ids = ids[orden]
    truncado = len(ids) > options.limit
    ids = ids[: max(options.limit, 0)]

    tabla_cols: list[pd.Series] = []
    for t in textos_grupo:
        sel = t.iloc[ids].astype(object).reset_index(drop=True)
        tabla_cols.append(sel.where(sel != "", VACIO))
    cuentas = cantidades[ids]
    porc = cuentas * 100.0 / filas if filas else np.zeros(len(ids), dtype=np.float64)
    tabla = pd.concat(
        [*tabla_cols, pd.Series(cuentas, dtype="int64"), pd.Series(porc, dtype="float64")],
        axis=1,
    )
    tabla.columns = [*columns, COL_CANTIDAD, COL_PORCENTAJE]
    reporte.force(ProgressEvent("Listo", total_pasos, total_pasos))
    return DistinctResult(
        columns=list(columns),
        table=tabla,
        distinct_total=distinct_total,
        rows_total=filas,
        truncated=truncado,
        variant_groups=variantes,
    )


def rows_with_value(
    df: pd.DataFrame,
    columns: list[str],
    values: tuple[str, ...],
    normalize: bool = False,
) -> pd.Index:
    """Posiciones de las filas que tienen esos valores ("(vacío)" = "")."""
    if len(columns) != len(values):
        raise ExcelToolError("Tiene que haber un valor por cada columna.")
    mascara = np.ones(len(df), dtype=bool)
    for nombre, valor in zip(columns, values, strict=True):
        buscado = "" if valor == VACIO else valor
        codigos, unicos = _factorizar(_columna(df, nombre))
        if normalize:
            coincide = _bool(_normalizar(unicos) == buscado.strip().lower())
        else:
            coincide = _bool(unicos == buscado)
        mascara &= coincide[codigos]
    return pd.Index(np.flatnonzero(mascara))


# --------------------------------------------------------------------------- filtros


BLOQUE = 200_000


def _convertir_todo(valores: pd.Series, modo: str) -> tuple[np.ndarray, str | None]:
    """Convierte los valores únicos con `core/types`, de a bloques (para no duplicar toda la
    columna en memoria). El estilo decimal se decide una sola vez, mirando el primer bloque.
    Devuelve (convertidos con NaN/NaT donde no se pudo, estilo decimal)."""
    estilo = None
    if modo == "numero" and len(valores):
        estilo = tipos.detect_column_type(valores.iloc[:BLOQUE]).decimal_style
    partes = []
    for ini in range(0, len(valores), BLOQUE):
        bloque = valores.iloc[ini : ini + BLOQUE].reset_index(drop=True)
        if modo == "numero":
            conv = tipos.to_number(bloque, estilo).values.to_numpy(
                dtype=np.float64, na_value=np.nan
            )
        elif modo == "fecha":
            conv = tipos.to_date(bloque).values.to_numpy().astype("datetime64[us]")
        else:
            conv = tipos.to_bool(bloque).values.to_numpy(dtype=np.float64, na_value=np.nan)
        partes.append(conv)
    if not partes:
        vacio = np.empty(0, dtype="datetime64[us]" if modo == "fecha" else np.float64)
        return vacio, estilo
    return np.concatenate(partes), estilo


def _convertir_valor(valor: str, modo: str, estilo: str | None):
    """El valor del filtro convertido igual que los de la columna (NaN/NaT si no se puede)."""
    serie = _texto([valor])
    if modo == "numero":
        return tipos.to_number(serie, estilo).values.to_numpy(dtype=np.float64, na_value=np.nan)[0]
    return tipos.to_date(serie).values.to_numpy().astype("datetime64[us]")[0]


def _comparar(
    valores: pd.Series, op: str, valor: str, tipo: ColumnType, columna: str
) -> np.ndarray:
    """Máscara sobre los valores únicos para mayor/menor."""
    no_vacio = ~_vacios(valores)
    modo = None
    if tipo in ("entero", "decimal"):
        modo = "numero"
    elif tipo in ("fecha", "fecha_hora"):
        modo = "fecha"
    elif tipo == "mixto":
        for candidato in ("numero", "fecha"):
            estilo = (
                tipos.detect_column_type(valores.iloc[:BLOQUE]).decimal_style
                if candidato == "numero" and len(valores)
                else None
            )
            if not pd.isna(_convertir_valor(valor, candidato, estilo)):
                modo = candidato
                break
    if modo is None:
        base = valores.str.lower()
        ref = valor.lower()
        return _bool(base > ref if op == "mayor" else base < ref) & no_vacio
    conv, estilo = _convertir_todo(valores, modo)
    v = _convertir_valor(valor, modo, estilo)
    if pd.isna(v):
        quien = "un número" if modo == "numero" else "una fecha"
        raise ExcelToolError(f"«{valor}» no es {quien} válido para filtrar la columna «{columna}».")
    ok = ~pd.isna(conv)
    res = np.zeros(len(valores), dtype=bool)
    res[ok] = conv[ok] > v if op == "mayor" else conv[ok] < v
    return res


def _mascara_unicos(valores: pd.Series, cond: FilterCondition, tipo: ColumnType) -> np.ndarray:
    op = cond.op
    vacio = _vacios(valores)
    if op == "vacio":
        return vacio
    if op == "no_vacio":
        return ~vacio
    if op in ("igual", "distinto"):
        igual = _bool(valores.str.lower() == cond.value.lower())
        return igual if op == "igual" else ~igual
    if op == "contiene":
        if len(valores) == 0:
            return np.zeros(0, dtype=bool)
        return _bool(valores.str.lower().str.contains(cond.value.lower(), regex=False))
    if op in ("mayor", "menor"):
        return _comparar(valores, op, cond.value, tipo, cond.column)
    raise ExcelToolError(f"Operación de filtro desconocida: {op}.")


def filter_rows(
    df: pd.DataFrame,
    conditions: list[FilterCondition],
    column_types: dict[str, ColumnType],
) -> pd.Index:
    """Posiciones de las filas que cumplen TODAS las condiciones. Las celdas que no se pueden
    convertir al comparar con mayor/menor no pasan el filtro."""
    if not conditions:
        return pd.RangeIndex(len(df))
    mascara = np.ones(len(df), dtype=bool)
    for cond in conditions:
        codigos, unicos = _factorizar(_columna(df, cond.column))
        tipo = column_types.get(cond.column, "texto")
        mascara &= _mascara_unicos(unicos, cond, tipo)[codigos]
    return pd.Index(np.flatnonzero(mascara))


# --------------------------------------------------------------------------- orden


def _claves_orden(valores: pd.Series, tipo: ColumnType) -> tuple[np.ndarray, np.ndarray]:
    """(grupo, clave) por valor único: grupo 0 = convertible, 1 = no convertible, 2 = vacío."""
    vacio = _vacios(valores)
    modo = "texto"
    if tipo in ("entero", "decimal", "mixto"):
        modo = "numero"
    elif tipo in ("fecha", "fecha_hora"):
        modo = "fecha"
    elif tipo == "si_no":
        modo = "si_no"
    grupo = np.where(vacio, 2, 0).astype(np.int8)
    clave = np.zeros(len(valores), dtype=np.float64)
    if modo == "texto":
        return grupo, _rango_alfabetico(valores).astype(np.float64)
    conv, _ = _convertir_todo(valores, modo)
    if modo == "fecha":
        nulos = np.isnat(conv)
        conv = conv.astype(np.int64).astype(np.float64)
        conv[nulos] = np.nan
    malo = np.isnan(conv) & ~vacio
    grupo[malo] = 1
    if tipo == "mixto":
        # lo que no es número se ordena alfabéticamente, después de los números
        conv = np.where(malo, _rango_alfabetico(valores).astype(np.float64), conv)
    clave = np.where(np.isnan(conv), 0.0, conv)
    return grupo, clave


def sort_rows(
    df: pd.DataFrame,
    rows: pd.Index,
    column: str,
    ascending: bool,
    column_type: ColumnType,
) -> pd.Index:
    """Las mismas filas ordenadas por la columna (orden estable). Vacíos y valores que no se
    pueden convertir van siempre al final, también en orden descendente."""
    pos = np.asarray(rows, dtype=np.int64)
    if len(pos) == 0:
        return pd.Index(pos)
    sub = _columna(df, column).iloc[pos]
    codigos, unicos = _factorizar(sub)
    del sub
    grupo, clave = _claves_orden(unicos, column_type)
    g = grupo[codigos]
    k = clave[codigos]
    if not ascending:
        k = -k
    orden = np.lexsort((k, g))
    return pd.Index(pos[orden])


# --------------------------------------------------------------------------- páginas


def get_page(df: pd.DataFrame, rows: pd.Index, page: int, page_size: int = 1000) -> pd.DataFrame:
    """Una página (desde 0) de las filas `rows`, con la columna "N.º fila" a la izquierda
    (1 = primera fila de datos). Una página fuera de rango devuelve 0 filas."""
    if page_size <= 0:
        raise ValueError("page_size tiene que ser mayor que 0")
    inicio = max(page, 0) * page_size
    pos = np.asarray(rows[inicio : inicio + page_size] if page >= 0 else rows[:0], dtype=np.int64)
    pagina = df.iloc[pos].reset_index(drop=True)
    pagina.insert(0, COL_FILA, pd.Series(pos + 1, dtype="int64"))
    return pagina
