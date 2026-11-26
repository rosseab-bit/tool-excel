"""Resúmenes (agrupar y calcular). Contrato: docs/contratos.md § Fase 4 (ADR-034).

TODOS los resúmenes de la app pasan por `summarize` (también el resumen final del join).
La conversión a número y a fecha NO se hace acá: se usa `core/types`.

Cómo funciona (vectorizado, sin bucles por grupo):
- Cada clave de grupo se factoriza (códigos en orden de primera aparición); si hay varias se
  combinan los códigos y se vuelve a factorizar. Los grupos salen en orden de primera aparición.
- Sumas y promedios son EXACTOS: los valores se pasan a enteros escalados (10**decimales) y se
  suman como enteros, para que 0,1 + 0,2 dé 0,3 y no 0,30000000000000004.
- La salida es texto en formato argentino: coma decimal, sin separador de miles, enteros sin
  ".0", decimales de 2 a 6 cifras.

Reglas que tiene que saber quien use esto:
- Un valor vacío (o solo espacios) NO cuenta como inválido: no suma y se informa aparte en un aviso.
  `not_numeric[columna]` cuenta solo las celdas CON contenido que no se pudieron convertir.
- `not_numeric` y los avisos de números/fechas se informan una vez por columna de origen.
- Sin `group_by` (y con cálculos) devuelve UNA fila con el resultado de toda la tabla (tabla sin
  filas → 0 filas). Con `group_by` y sin cálculos devuelve las combinaciones distintas.
- Clave de grupo vacía → grupo "(vacío)". Fecha inválida al agrupar por mes/año → "(sin fecha)".
- Una columna con varios tipos ("mixto") se trata como número o como fecha si más de la mitad
  de sus valores con contenido lo son; si no, como texto.
- `first`/`last`: el primer/último valor CON contenido en el orden de la tabla. Si ese valor se
  puede leer como número o fecha se muestra en formato argentino; si no, tal cual.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np
import pandas as pd
import pyarrow as pa
import pyarrow.compute as pc

from excel_tool.core.errors import ExcelToolError
from excel_tool.core.progress import CancelToken, ProgressCallback, ProgressEvent
from excel_tool.core.types import ColumnType, detect_column_type, to_date, to_number

AggFunc = Literal[
    "sum", "count", "count_distinct", "mean", "min", "max", "first", "last", "join_text"
]

VACIO = "(vacío)"
SIN_FECHA = "(sin fecha)"
MAX_LARGO_LISTA = 1000
_STR = pd.StringDtype("pyarrow")
_MAX_ESCALADO = 9_000_000_000_000_000_000  # margen bajo el tope de int64
_MAX_DECIMALES = 6

_NOMBRES: dict[str, str] = {
    "sum": "Suma de",
    "count_distinct": "Distintos de",
    "mean": "Promedio de",
    "min": "Mínimo de",
    "max": "Máximo de",
    "first": "Primero de",
    "last": "Último de",
    "join_text": "Lista de",
}
_USA_TIPO = {"min", "max", "first", "last"}


@dataclass(frozen=True)
class GroupBy:
    column: str
    period: Literal["month", "year"] | None = None  # agrupar por mes/año de una fecha


@dataclass(frozen=True)
class Aggregation:
    column: str  # "" para count (contar filas)
    func: AggFunc
    name: str | None = None  # None → automático: "Suma de Importe", "Cantidad de filas", ...


@dataclass(frozen=True)
class AggregationSpec:
    group_by: list[GroupBy]
    aggregations: list[Aggregation]


@dataclass(frozen=True)
class SummaryResult:
    df: pd.DataFrame  # columnas: group_by (en orden) + aggregations; texto con formato argentino
    column_types: dict[str, ColumnType]
    not_numeric: dict[str, int]  # columna → celdas con contenido que no se pudieron convertir
    warnings: list[str]


# --------------------------------------------------------------------------- nombres


def default_name(agg: Aggregation) -> str:
    """Nombre automático de la columna resultado ("Suma de Importe", "Cantidad de filas")."""
    if agg.func == "count":
        return "Cantidad de filas"
    return f"{_NOMBRES[agg.func]} {agg.column}"


def group_column_name(group: GroupBy) -> str:
    """Nombre de la columna de un grupo ("Fecha (mes)" / "Fecha (año)" si agrupa por período)."""
    if group.period == "month":
        return f"{group.column} (mes)"
    if group.period == "year":
        return f"{group.column} (año)"
    return group.column


def result_columns(spec: AggregationSpec) -> list[str]:
    """Nombres finales de las columnas del resultado (grupos + cálculos), ya sin repetidos."""
    nombres: list[str] = []
    usados: set[str] = set()
    todos = [group_column_name(g) for g in spec.group_by] + [
        a.name if a.name else default_name(a) for a in spec.aggregations
    ]
    for nombre in todos:
        final = nombre
        n = 2
        while final in usados:
            final = f"{nombre} ({n})"
            n += 1
        usados.add(final)
        nombres.append(final)
    return nombres


def _plural(x: int, singular: str, plural: str) -> str:
    """«1 celda» / «3 celdas» (con punto de miles)."""
    return f"{_n(x)} {singular if x == 1 else plural}"


def _n(x: int) -> str:
    """Entero con punto de miles (para los avisos)."""
    return f"{x:,}".replace(",", ".")


# --------------------------------------------------------------------------- columnas


def _texto(valores) -> pd.Series:
    return pd.Series(valores, dtype=_STR, copy=False).reset_index(drop=True)


def _buscar(df: pd.DataFrame, nombre: str) -> pd.Series:
    posiciones = np.flatnonzero(df.columns == nombre)
    if len(posiciones) == 0:
        raise ExcelToolError(f"No encuentro la columna «{nombre}» en la tabla.")
    return _texto(df.iloc[:, int(posiciones[0])].fillna(""))


def _vacios(serie: pd.Series) -> np.ndarray:
    arr = pc.equal(pc.utf8_trim_whitespace(pa.Array.from_pandas(serie)), "")
    return arr.fill_null(False).to_numpy(zero_copy_only=False).astype(bool)


class _Columna:
    """Una columna de origen, con sus conversiones calculadas una sola vez."""

    def __init__(self, nombre: str, serie: pd.Series, tipo: ColumnType | None) -> None:
        self.nombre = nombre
        self.raw = serie
        self.vacio = _vacios(serie)
        self._tipo = tipo
        self._num: tuple[np.ndarray, int] | None = None
        self._fecha: tuple[np.ndarray, int] | None = None
        self._clase: str | None = None

    @property
    def tipo(self) -> ColumnType:
        if self._tipo is None:
            self._tipo = detect_column_type(self.raw).detected
        return self._tipo

    def numeros(self) -> tuple[np.ndarray, int]:
        """(valores float con NaN donde no hay número, cantidad de inválidos con contenido)."""
        if self._num is None:
            r = to_number(self.raw)
            self._num = (r.values.to_numpy(dtype="float64"), len(r.failed))
        return self._num

    def fechas(self) -> tuple[np.ndarray, int]:
        """(datetime64[ns] con NaT donde no hay fecha, cantidad de inválidas con contenido)."""
        if self._fecha is None:
            r = to_date(self.raw)
            self._fecha = (r.values.to_numpy(dtype="datetime64[ns]"), len(r.failed))
        return self._fecha

    def clase(self) -> str:
        """'num', 'fecha' o 'texto': cómo se comparan y se muestran los valores."""
        if self._clase is None:
            tipo = self.tipo
            if tipo in ("entero", "decimal"):
                self._clase = "num"
            elif tipo in ("fecha", "fecha_hora"):
                self._clase = "fecha"
            elif tipo == "mixto":
                con_contenido = int((~self.vacio).sum())
                valores, _ = self.numeros()
                n_num = int(np.isfinite(valores).sum())
                fechas, _ = self.fechas()
                n_fecha = int((~np.isnat(fechas)).sum())
                if con_contenido and n_num * 2 > con_contenido and n_num >= n_fecha:
                    self._clase = "num"
                elif con_contenido and n_fecha * 2 > con_contenido:
                    self._clase = "fecha"
                else:
                    self._clase = "texto"
            else:
                self._clase = "texto"
        return self._clase


# --------------------------------------------------------------------------- números


def _decimales(valores: np.ndarray) -> int:
    """Menor cantidad de decimales (0 a 6) con la que se representan todos los valores."""
    if not len(valores):
        return 0
    for d in range(_MAX_DECIMALES + 1):
        esc = valores * (10.0**d)
        tol = 1e-12 * np.maximum(1.0, np.abs(esc))
        if bool(np.all(np.abs(esc - np.rint(esc)) <= tol)):
            return d
    return _MAX_DECIMALES


def _escalar(valores: np.ndarray, d: int) -> np.ndarray:
    """int64 de valores * 10**d (los NaN quedan en 0; usar una máscara aparte)."""
    return np.rint(np.nan_to_num(valores, nan=0.0) * (10.0**d)).astype(np.int64)


def _formatear(escalado: np.ndarray, d: int, minimo: int) -> np.ndarray:
    """Texto argentino de enteros escalados por 10**d.

    d == 0 → entero sin decimales. d > 0 → decimales de `minimo` a d cifras (se sacan los
    ceros de la derecha hasta `minimo`)."""
    n = len(escalado)
    neg = escalado < 0
    a = np.abs(escalado)
    if d == 0:
        salida = a.astype(str).astype(object)
    else:
        if d < minimo:
            a = a * (10 ** (minimo - d))
            d = minimo
        base = 10**d
        ent = a // base
        frac = a % base
        decs = np.full(n, d)
        for k in range(d, minimo, -1):
            m = (decs == k) & (frac % 10 == 0)
            frac[m] //= 10
            decs[m] = k - 1
        salida = np.empty(n, dtype=object)
        for k in np.unique(decs):
            m = decs == k
            salida[m] = np.char.add(
                np.char.add(ent[m].astype(str), ","), np.char.zfill(frac[m].astype(str), int(k))
            )
    if neg.any():
        salida[neg] = np.char.add("-", salida[neg].astype(str))
    return salida


def _por_grupo(valores: np.ndarray, codigos: np.ndarray, n_grupos: int, func: str):
    """Suma/mín/máx de enteros o fechas por grupo (los grupos sin dato quedan vacíos)."""
    g = pd.Series(valores).groupby(codigos, sort=True)
    r = getattr(g, func)()
    return r.reindex(range(n_grupos))


def _sumar(valores: np.ndarray, codigos: np.ndarray, n_grupos: int):
    """(suma exacta escalada por grupo, decimales usados, cantidad de válidos por grupo)."""
    ok = np.isfinite(valores)
    d = _decimales(valores[ok])
    esc = _escalar(valores, d)
    esc[~ok] = 0
    # Si la suma pudiera pasarse de int64, se baja la precisión en lugar de desbordar.
    tope = int(np.abs(esc).max()) * int(ok.sum()) if ok.any() else 0
    while d > 0 and tope > _MAX_ESCALADO:
        d -= 1
        esc = _escalar(valores, d)
        esc[~ok] = 0
        tope = int(np.abs(esc).max()) * int(ok.sum())
    suma = _por_grupo(esc, codigos, n_grupos, "sum").fillna(0).to_numpy(dtype=np.int64)
    cuenta = np.bincount(codigos[ok], minlength=n_grupos)
    return suma, d, cuenta


# --------------------------------------------------------------------------- claves


def _primera_aparicion(codigos: np.ndarray, n_grupos: int) -> np.ndarray:
    """Posición de la primera fila de cada grupo (los códigos van en orden de aparición)."""
    n = len(codigos)
    primero = np.zeros(n_grupos, dtype=np.int64)
    primero[codigos[::-1]] = np.arange(n - 1, -1, -1, dtype=np.int64)
    return primero


def _clave_texto(col: _Columna) -> tuple[np.ndarray, list[str]]:
    serie = col.raw.mask(col.vacio, "")
    codigos, unicos = pd.factorize(serie)
    etiquetas = [VACIO if v == "" else str(v) for v in unicos]
    return codigos.astype(np.int64, copy=False), etiquetas


def _clave_periodo(col: _Columna, periodo: str) -> tuple[np.ndarray, list[str], int]:
    """(códigos, etiquetas, cantidad de fechas inválidas)."""
    fechas, _ = col.fechas()
    nat = np.isnat(fechas)
    invalidas = nat & ~col.vacio
    anios = fechas.astype("datetime64[Y]").astype(np.int64) + 1970
    if periodo == "month":
        meses = fechas.astype("datetime64[M]").astype(np.int64) % 12 + 1
        clave = anios * 100 + meses
    else:
        clave = anios
    clave = np.where(nat, 0, clave)
    clave = np.where(invalidas, -1, clave)
    clave = np.where(col.vacio, -2, clave)
    codigos, unicos = pd.factorize(clave)
    etiquetas = []
    for v in unicos:
        v = int(v)
        if v == -2:
            etiquetas.append(VACIO)
        elif v == -1:
            etiquetas.append(SIN_FECHA)
        elif periodo == "month":
            etiquetas.append(f"{v // 100:04d}-{v % 100:02d}")
        else:
            etiquetas.append(f"{v:04d}")
    return codigos.astype(np.int64, copy=False), etiquetas, int(invalidas.sum())


def _ordenar_periodos(
    codigos: np.ndarray,
    ng: int,
    claves: list[tuple[np.ndarray, list[str]]],
    group_by: list[GroupBy] | tuple[GroupBy, ...],
) -> np.ndarray:
    """Renumera los grupos para que los períodos salgan en orden cronológico.

    Se ordena (estable) solo por las columnas de período; «(vacío)» y «(sin fecha)» van al final.
    Los grupos que empatan conservan el orden de primera aparición.
    """
    primeros = _primera_aparicion(codigos, ng)
    claves_orden = []
    for (cod, etiquetas), g in zip(claves, group_by, strict=True):
        if not g.period:
            continue
        # las etiquetas de período con ceros a la izquierda ordenan bien como texto
        rango = np.empty(len(etiquetas), dtype=np.int64)
        normales = sorted(e for e in etiquetas if e not in (VACIO, SIN_FECHA))
        pos = {e: i for i, e in enumerate(normales)}
        for i, e in enumerate(etiquetas):
            if e == VACIO:
                rango[i] = len(normales)
            elif e == SIN_FECHA:
                rango[i] = len(normales) + 1
            else:
                rango[i] = pos[e]
        claves_orden.append(rango[cod[primeros]])
    orden = np.lexsort(tuple(reversed(claves_orden)))  # estable
    nuevo = np.empty(ng, dtype=np.int64)
    nuevo[orden] = np.arange(ng, dtype=np.int64)
    return nuevo[codigos]


# --------------------------------------------------------------------------- cálculos


@dataclass
class _Salida:
    valores: np.ndarray  # object, str
    tipo: ColumnType


def _tipo_numero(d: int) -> ColumnType:
    return "entero" if d == 0 else "decimal"


def _formato_fecha(fechas: pd.Series) -> tuple[np.ndarray, ColumnType]:
    """Texto dd/mm/aaaa (con hora si alguna fecha tiene hora) y el tipo de la columna."""
    con_hora = bool(((fechas != fechas.dt.normalize()) & fechas.notna()).any())
    formato = "%d/%m/%Y %H:%M:%S" if con_hora else "%d/%m/%Y"
    texto = fechas.dt.strftime(formato).fillna("").to_numpy(dtype=object)
    sin_dato = fechas.isna().all()
    tipo: ColumnType = "texto" if sin_dato else ("fecha_hora" if con_hora else "fecha")
    return texto, tipo


def _calc_suma(col: _Columna, codigos, ng) -> _Salida:
    valores, _ = col.numeros()
    suma, d, _ = _sumar(valores, codigos, ng)
    return _Salida(_formatear(suma, d, 2), _tipo_numero(d))


def _calc_promedio(col: _Columna, codigos, ng) -> _Salida:
    valores, _ = col.numeros()
    suma, d, cuenta = _sumar(valores, codigos, ng)
    con_dato = cuenta > 0
    medio = np.zeros(ng)
    medio[con_dato] = suma[con_dato] / cuenta[con_dato] / (10.0**d)
    dec = _MAX_DECIMALES
    while dec > 2 and np.abs(medio).max(initial=0.0) * (10.0**dec) > _MAX_ESCALADO:
        dec -= 1
    salida = _formatear(_escalar(medio, dec), dec, 2)
    salida[~con_dato] = ""
    return _Salida(salida, "decimal")


def _calc_min_max(col: _Columna, codigos, ng, func: str) -> _Salida:
    clase = col.clase()
    if clase == "num":
        valores, _ = col.numeros()
        ok = np.isfinite(valores)
        d = _decimales(valores[ok])
        esc = pd.Series(_escalar(valores, d)).where(ok)
        r = getattr(esc.groupby(codigos, sort=True), func)().reindex(range(ng))
        sin = r.isna().to_numpy()
        salida = _formatear(r.fillna(0).to_numpy(dtype=np.int64), d, 2)
        salida[sin] = ""
        return _Salida(salida, "texto" if sin.all() else _tipo_numero(d))
    if clase == "fecha":
        fechas, _ = col.fechas()
        r = getattr(pd.Series(fechas).groupby(codigos, sort=True), func)().reindex(range(ng))
        texto, tipo = _formato_fecha(r)
        return _Salida(texto, tipo)
    # texto: comparación alfabética de los valores con contenido
    s = col.raw.mask(col.vacio, pd.NA)
    r = getattr(s.groupby(codigos, sort=True), func)().reindex(range(ng))
    return _Salida(r.fillna("").to_numpy(dtype=object), "texto")


def _calc_primero_ultimo(col: _Columna, codigos, ng, ultimo: bool) -> _Salida:
    idx = np.flatnonzero(~col.vacio)
    pos = np.full(ng, -1, dtype=np.int64)
    if ultimo:
        pos[codigos[idx]] = idx
    else:
        pos[codigos[idx][::-1]] = idx[::-1]
    hay = pos >= 0
    crudo = np.full(ng, "", dtype=object)
    if hay.any():
        crudo[hay] = col.raw.iloc[pos[hay]].to_numpy(dtype=object)
    clase = col.clase()
    if clase == "num":
        valores, _ = col.numeros()
        v = np.full(ng, np.nan)
        v[hay] = valores[pos[hay]]
        ok = np.isfinite(v)
        if not ok.any():
            return _Salida(crudo, "texto")
        d = _decimales(v[ok])
        salida = _formatear(_escalar(v, d), d, 2)
        salida[~ok] = crudo[~ok]
        return _Salida(salida, _tipo_numero(d) if ok.sum() == hay.sum() else "texto")
    if clase == "fecha":
        fechas, _ = col.fechas()
        v = np.full(ng, np.datetime64("NaT", "ns"), dtype="datetime64[ns]")
        v[hay] = fechas[pos[hay]]
        ok = ~np.isnat(v)
        if not ok.any():
            return _Salida(crudo, "texto")
        texto, tipo = _formato_fecha(pd.Series(v))
        texto[~ok] = crudo[~ok]
        return _Salida(texto, tipo if ok.sum() == hay.sum() else "texto")
    return _Salida(crudo, "texto")


def _calc_distintos(col: _Columna, codigos, ng) -> _Salida:
    sel = np.flatnonzero(~col.vacio)
    if not len(sel):
        return _Salida(np.full(ng, "0", dtype=object), "entero")
    vcod, unicos = pd.factorize(col.raw.iloc[sel])
    nv = max(len(unicos), 1)
    par = np.unique(codigos[sel].astype(np.int64) * nv + vcod)
    cuenta = np.bincount(par // nv, minlength=ng)
    return _Salida(cuenta.astype(str).astype(object), "entero")


def _calc_lista(col: _Columna, codigos, ng, avisos: list[str], nombre: str) -> _Salida:
    sel = np.flatnonzero(~col.vacio)
    if not len(sel):
        return _Salida(np.full(ng, "", dtype=object), "texto")
    vcod, unicos = pd.factorize(col.raw.iloc[sel])
    nv = max(len(unicos), 1)
    par = codigos[sel].astype(np.int64) * nv + vcod
    primera = ~pd.Series(par).duplicated().to_numpy(dtype=bool)
    pos = sel[primera]  # valores sin repetir por grupo, en orden de aparición
    g = codigos[pos]
    orden = np.argsort(g, kind="stable")
    pos, g = pos[orden], g[orden]
    cuentas = np.bincount(g, minlength=ng)
    offsets = np.concatenate([[0], np.cumsum(cuentas)]).astype(np.int64)
    valores = pa.array(col.raw.iloc[pos]).cast(pa.string())
    listas = pa.ListArray.from_arrays(pa.array(offsets, type=pa.int32()), valores)
    unidos = pc.binary_join(listas, ", ")
    largos = pc.utf8_length(unidos).to_numpy(zero_copy_only=False)
    salida = np.array(unidos.to_pylist(), dtype=object)
    largos_idx = np.flatnonzero(largos > MAX_LARGO_LISTA)
    for i in largos_idx:  # solo los grupos que se pasan del tope
        corte = salida[i][: MAX_LARGO_LISTA - 3]
        sep = corte.rfind(", ")
        if sep > 0:
            corte = corte[:sep]
        salida[i] = corte + "…"
    if len(largos_idx):
        k = len(largos_idx)
        lista = _plural(k, "lista que pasaba", "listas que pasaban")
        corte_txt = "se cortó y termina" if k == 1 else "se cortaron y terminan"
        avisos.append(
            f"En «{nombre}» hay {lista} de {_n(MAX_LARGO_LISTA)} caracteres: {corte_txt} en «…»."
        )
    return _Salida(salida, "texto")


# --------------------------------------------------------------------------- summarize


def _tipos_dados(column_types: dict[str, ColumnType] | None, nombre: str) -> ColumnType | None:
    return column_types.get(nombre) if column_types else None


def summarize(
    df: pd.DataFrame,
    spec: AggregationSpec,
    *,
    column_types: dict[str, ColumnType] | None = None,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> SummaryResult:
    """Agrupa `df` según `spec` y calcula cada resumen. Ver el docstring del módulo."""
    if not spec.group_by and not spec.aggregations:
        raise ExcelToolError("Elegí por qué columnas agrupar o agregá al menos un cálculo.")
    for a in spec.aggregations:
        if a.func != "count" and not a.column:
            raise ExcelToolError("Falta elegir la columna de uno de los cálculos.")
        if a.func not in _NOMBRES and a.func != "count":
            raise ExcelToolError(f"Cálculo desconocido: {a.func}.")

    total_pasos = len(spec.group_by) + len(spec.aggregations) + 1
    hechos = 0

    def avanzar(mensaje: str) -> None:
        nonlocal hechos
        if cancel:
            cancel.check()
        if progress:
            progress(ProgressEvent(mensaje, hechos, total_pasos))
        hechos += 1

    columnas: dict[str, _Columna] = {}

    def columna(nombre: str) -> _Columna:
        if nombre not in columnas:
            columnas[nombre] = _Columna(
                nombre, _buscar(df, nombre), _tipos_dados(column_types, nombre)
            )
        return columnas[nombre]

    nombres = result_columns(spec)
    n = len(df)
    avisos: list[str] = []
    no_numericos: dict[str, int] = {}

    # --- grupos
    codigos = np.zeros(n, dtype=np.int64)
    claves: list[tuple[np.ndarray, list[str]]] = []
    tipos_claves: list[ColumnType] = []
    for g in spec.group_by:
        avanzar(f"Agrupando por «{g.column}»")
        col = columna(g.column)
        if g.period:
            cod, etiquetas, invalidas = _clave_periodo(col, g.period)
            if invalidas:
                fechas = _plural(
                    invalidas, "fecha que no se pudo leer", "fechas que no se pudieron leer"
                )
                avisos.append(f"En «{g.column}» hay {fechas}: quedaron en el grupo «{SIN_FECHA}».")
            claves.append((cod, etiquetas))
            limpio = VACIO not in etiquetas and SIN_FECHA not in etiquetas
            tipos_claves.append("entero" if g.period == "year" and limpio else "texto")
        else:
            cod, etiquetas = _clave_texto(col)
            claves.append((cod, etiquetas))
            tipo = col.tipo
            usable = tipo in ("entero", "decimal", "fecha", "fecha_hora") and VACIO not in etiquetas
            tipos_claves.append(tipo if usable else "texto")
        k = len(claves[-1][1])
        codigos = cod if len(claves) == 1 else codigos * k + cod
        if len(claves) > 1:
            codigos = pd.factorize(codigos)[0].astype(np.int64, copy=False)
    if spec.group_by and n:
        if len(claves) == 1:
            codigos = codigos.astype(np.int64, copy=False)
        ng = int(codigos.max()) + 1
    else:
        ng = 1 if n else 0
        codigos = np.zeros(n, dtype=np.int64)

    if ng > 1 and any(g.period for g in spec.group_by):
        codigos = _ordenar_periodos(codigos, ng, claves, spec.group_by)

    salida: dict[str, np.ndarray] = {}
    tipos: dict[str, ColumnType] = {}
    primeros = _primera_aparicion(codigos, ng) if ng else np.zeros(0, dtype=np.int64)
    for i, (cod, etiquetas) in enumerate(claves):
        arr = np.array(etiquetas + [""], dtype=object)
        salida[nombres[i]] = arr[cod[primeros]] if ng else np.zeros(0, dtype=object)
        tipos[nombres[i]] = tipos_claves[i]

    # --- cálculos
    avisados_num: set[str] = set()
    for j, agg in enumerate(spec.aggregations):
        nombre = nombres[len(spec.group_by) + j]
        avanzar(f"Calculando «{nombre}»")
        if ng == 0:
            salida[nombre] = np.zeros(0, dtype=object)
            tipos[nombre] = "entero" if agg.func in ("count", "count_distinct") else "texto"
            continue
        f = agg.func
        if f == "count":
            r = _Salida(np.bincount(codigos, minlength=ng).astype(str).astype(object), "entero")
        else:
            col = columna(agg.column)
            if f == "count_distinct":
                r = _calc_distintos(col, codigos, ng)
            elif f == "join_text":
                r = _calc_lista(col, codigos, ng, avisos, nombre)
            elif f in ("sum", "mean"):
                r = (_calc_suma if f == "sum" else _calc_promedio)(col, codigos, ng)
                _anotar_numeros(col, "num", no_numericos, avisos, avisados_num)
            elif f in ("min", "max"):
                r = _calc_min_max(col, codigos, ng, f)
                _anotar_numeros(col, col.clase(), no_numericos, avisos, avisados_num)
            else:
                r = _calc_primero_ultimo(col, codigos, ng, f == "last")
        salida[nombre] = r.valores
        tipos[nombre] = r.tipo

    avanzar("Armando el resultado")
    out = pd.DataFrame({k: pd.Series(v, dtype=_STR) for k, v in salida.items()})
    if not len(salida) or not ng:
        out = pd.DataFrame({k: pd.Series([], dtype=_STR) for k in nombres})
    if progress:
        progress(ProgressEvent("Resumen listo", total_pasos, total_pasos))
    return SummaryResult(df=out, column_types=tipos, not_numeric=no_numericos, warnings=avisos)


def _anotar_numeros(
    col: _Columna,
    clase: str,
    no_numericos: dict[str, int],
    avisos: list[str],
    avisados: set[str],
) -> None:
    """Registra (una vez por columna) las celdas con contenido que no se pudieron convertir."""
    if clase not in ("num", "fecha") or col.nombre in avisados:
        return
    avisados.add(col.nombre)
    _, invalidos = col.numeros() if clase == "num" else col.fechas()
    vacios = int(col.vacio.sum())
    que = "números" if clase == "num" else "fechas"
    sing, plu = (
        ("no se suma", "no se suman")
        if clase == "num"
        else ("no se tiene en cuenta", "no se tienen en cuenta")
    )
    accion_inv = sing if invalidos == 1 else plu
    accion_vac = sing if vacios == 1 else plu
    if invalidos:
        no_numericos[col.nombre] = invalidos
        celdas = _plural(
            invalidos,
            "celda con contenido que no se pudo leer",
            "celdas con contenido que no se pudieron leer",
        )
        avisos.append(f"En «{col.nombre}» hay {celdas} como {que}: {accion_inv}.")
    if vacios:
        celdas = _plural(vacios, "celda vacía", "celdas vacías")
        avisos.append(f"Hay {celdas} de {col.nombre}: {accion_vac}.")
