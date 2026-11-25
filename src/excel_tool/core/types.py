"""Detección de tipos y conversión de texto a número, fecha y sí/no.

Es el ÚNICO lugar donde se decide qué es un número o una fecha. Todo se lee como texto y
acá se interpreta, siempre sobre la serie completa (vectorizado) y sin perder en silencio
las celdas que no se pueden convertir.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np
import pandas as pd

from excel_tool.core.progress import CancelToken

ColumnType = Literal["texto", "entero", "decimal", "fecha", "fecha_hora", "si_no", "mixto", "vacia"]
ValueKind = Literal["texto", "entero", "decimal", "fecha", "fecha_hora", "si_no", "vacia"]
DecimalStyle = Literal[",", "."]

_TRUE = ("sí", "si", "verdadero", "true")
_FALSE = ("no", "falso", "false")
_SI_NO = frozenset(_TRUE + _FALSE)


@dataclass
class TypeDetection:
    detected: ColumnType
    breakdown: dict[ValueKind, int]  # cuenta por clase de valor, sin "vacia"
    empty: int  # vacíos ("" o solo espacios)
    decimal_style: Literal[",", "."] | None  # para entero/decimal/mixto con números


@dataclass
class ConversionResult:
    values: (
        pd.Series
    )  # float (número), datetime64 (fecha) o boolean; NaN/NaT/NA si falló o estaba vacío
    failed: list[
        tuple[int, str]
    ]  # (posición, valor original) de las celdas NO vacías no convertidas


def _patrones_numero(estilo: DecimalStyle) -> tuple[str, str]:
    """Regex (entero, decimal) para el estilo pedido."""
    miles = r"\." if estilo == "," else ","
    dec = "," if estilo == "," else r"\."
    base = rf"(?:0|[1-9]\d*|[1-9]\d{{0,2}}(?:{miles}\d{{3}})+)"
    return rf"-?{base}", rf"-?{base}{dec}\d+"


def _como_texto(values: pd.Series) -> np.ndarray:
    """Array de objetos str (los nulos pasan a "")."""
    arr = values.to_numpy(dtype=object, na_value="")
    return np.where(pd.isna(arr), "", arr).astype(object)


def _strip(arr: np.ndarray) -> pd.Series:
    return pd.Series(arr, dtype=object).str.strip()


def _match(serie: pd.Series, patron: str) -> np.ndarray:
    return serie.str.fullmatch(patron).fillna(False).to_numpy(dtype=bool)


def _parse_fechas(limpio: pd.Series) -> tuple[pd.Series, np.ndarray]:
    """(fechas, con_hora): fechas con NaT donde no hay una fecha válida (31/02, 25:00, etc.)."""
    n = len(limpio)
    fechas = pd.Series(pd.NaT, index=range(n), dtype="datetime64[ns]")
    con_hora = np.zeros(n, dtype=bool)
    # Filtros baratos (sin regex) para no probar los formatos sobre todo el texto.
    con_dos_puntos = limpio.str.contains(":", regex=False).to_numpy(dtype=bool)
    con_segundos = np.zeros(n, dtype=bool)
    if con_dos_puntos.any():
        con_segundos = limpio.str.count(":").to_numpy() >= 2
    candidato_barra = limpio.str.contains("/", regex=False).to_numpy(dtype=bool)
    candidato_guion = limpio.str.contains("-", regex=False).to_numpy(dtype=bool) & ~candidato_barra

    grupos = (
        (con_segundos, " %H:%M:%S"),
        (con_dos_puntos & ~con_segundos, " %H:%M"),
        (~con_dos_puntos, ""),
    )
    for sep, candidato in (("/", candidato_barra), ("-", candidato_guion)):
        for en_grupo, sufijo in grupos:
            idx = np.flatnonzero(candidato & en_grupo)
            if not len(idx):
                continue
            formato = f"%d{sep}%m{sep}%Y{sufijo}"
            parseadas = pd.to_datetime(limpio.iloc[idx], format=formato, errors="coerce")
            fechas.iloc[idx] = parseadas.to_numpy()
            con_hora[idx] = bool(sufijo)
    return fechas, con_hora


def _clasificar(limpio: pd.Series, estilo: DecimalStyle) -> np.ndarray:
    """Clase de cada valor (ya sin espacios en los bordes) como array de objetos."""
    n = len(limpio)
    clases = np.full(n, "texto", dtype=object)
    vacio = (limpio == "").to_numpy(dtype=bool)
    clases[vacio] = "vacia"
    pendiente = ~vacio

    def aplicar(patron: str, clase: str | None = None) -> np.ndarray:
        """Marca (sobre lo pendiente) lo que cumple el patrón; devuelve el hit global."""
        nonlocal pendiente
        idx = np.flatnonzero(pendiente)
        hit = np.zeros(n, dtype=bool)
        if len(idx):
            hit[idx] = _match(limpio.iloc[idx], patron)
        if clase:
            clases[hit] = clase
            pendiente = pendiente & ~hit
        return hit

    pat_ent, pat_dec = _patrones_numero(estilo)
    numero = aplicar(f"{pat_ent}|{pat_dec}")
    sep = "," if estilo == "," else "."
    es_dec = np.zeros(n, dtype=bool)
    idx_num = np.flatnonzero(numero)
    if len(idx_num):
        es_dec[idx_num] = limpio.iloc[idx_num].str.contains(sep, regex=False).to_numpy(dtype=bool)
    clases[numero & ~es_dec] = "entero"
    clases[numero & es_dec] = "decimal"
    pendiente = pendiente & ~numero

    idx = np.flatnonzero(pendiente)
    if len(idx):
        fechas, con_hora = _parse_fechas(limpio.iloc[idx].reset_index(drop=True))
        valida = fechas.notna().to_numpy(dtype=bool)
        sub = clases[idx]
        sub[valida & ~con_hora] = "fecha"
        sub[valida & con_hora] = "fecha_hora"
        clases[idx] = sub
        pendiente[idx[valida]] = False

    idx = np.flatnonzero(pendiente)
    if len(idx):
        si_no = limpio.iloc[idx].str.lower().isin(_SI_NO).to_numpy(dtype=bool)
        clases[idx[si_no]] = "si_no"
    return clases


def classify_value(value: str, decimal_style: DecimalStyle = ",") -> ValueKind:
    """Clase de un único valor de texto."""
    serie = pd.Series([value if isinstance(value, str) else ""], dtype=object)
    return _clasificar(serie.str.strip(), decimal_style)[0]  # type: ignore[return-value]


BLOQUE_UNICOS = 200_000


def _resolver_unicos(values: pd.Series):
    """Devuelve (codigos, unicos) para clasificar cada valor distinto una sola vez.

    `unicos` queda en su forma compacta (pyarrow si la columna lo es); se pasa a texto de
    Python por bloques (`_bloques`), para no armar millones de objetos de una sola vez.
    """
    if isinstance(values.dtype, pd.StringDtype):
        codigos, unicos = pd.factorize(values.fillna("").array)
        return codigos, unicos
    codigos, unicos = pd.factorize(_como_texto(values))
    return codigos, np.asarray(unicos, dtype=object)


def _bloques(unicos):
    """(inicio, valores sin espacios en los bordes) de a BLOQUE_UNICOS valores."""
    for ini in range(0, len(unicos), BLOQUE_UNICOS):
        trozo = np.asarray(unicos[ini : ini + BLOQUE_UNICOS], dtype=object)
        yield ini, _strip(trozo)


def _cuentas_estilo(limpio: pd.Series) -> tuple[int, int, int]:
    """(con separador, válidos como argentino, válidos como inglés).

    Solo importan las celdas con "." o ",": un número sin separadores vale igual en los dos.
    """
    con_sep = limpio[limpio.str.contains(r"[.,]", regex=True).to_numpy(dtype=bool)]
    if con_sep.empty:
        return 0, 0, 0
    cuentas = []
    for estilo in (",", "."):
        e, d = _patrones_numero(estilo)  # type: ignore[arg-type]
        cuentas.append(int(_match(con_sep, f"{e}|{d}").sum()))
    return len(con_sep), cuentas[0], cuentas[1]


def _elegir_estilo(con_sep: int, argentino: int, ingles: int) -> DecimalStyle:
    """Estilo con más valores numéricos; si empatan, el argentino (",")."""
    if con_sep == 0 or argentino == con_sep:
        return ","
    return "." if ingles > argentino else ","


def _estilo_de(limpio: pd.Series) -> DecimalStyle:
    return _elegir_estilo(*_cuentas_estilo(limpio))


def _tipo_de_columna(presentes: set[str]) -> ColumnType:
    if not presentes:
        return "vacia"
    if len(presentes) == 1:
        return next(iter(presentes))  # type: ignore[return-value]
    if presentes == {"entero", "decimal"}:
        return "decimal"
    if presentes == {"fecha", "fecha_hora"}:
        return "fecha_hora"
    return "mixto"


def detect_column_type(values: pd.Series, *, cancel: CancelToken | None = None) -> TypeDetection:
    """Tipo de una columna mirando todos sus valores.

    Con `cancel`, mira la bandera en cada bloque de valores distintos y lanza `Cancelado`.
    """
    codigos, unicos = _resolver_unicos(values)
    if cancel:
        cancel.check()
    repeticiones = np.bincount(codigos[codigos >= 0], minlength=len(unicos))
    del codigos
    totales = np.zeros(3, dtype=np.int64)
    for _, limpio in _bloques(unicos):
        if cancel:
            cancel.check()
        totales += _cuentas_estilo(limpio)
    estilo = _elegir_estilo(*(int(x) for x in totales))

    cuenta: dict[str, int] = {}
    for ini, limpio in _bloques(unicos):
        if cancel:
            cancel.check()
        clases = _clasificar(limpio, estilo)
        for clase, cant in zip(clases, repeticiones[ini : ini + len(clases)], strict=True):
            cuenta[clase] = cuenta.get(clase, 0) + int(cant)
    vacios = cuenta.pop("vacia", 0)
    desglose = {k: v for k, v in cuenta.items() if v > 0}
    tipo = _tipo_de_columna(set(desglose))
    con_numeros = "entero" in desglose or "decimal" in desglose
    return TypeDetection(
        detected=tipo,
        breakdown=desglose,  # type: ignore[arg-type]
        empty=vacios,
        decimal_style=estilo if con_numeros else None,
    )


def detect_types(
    df: pd.DataFrame, *, cancel: CancelToken | None = None
) -> dict[str, TypeDetection]:
    """Detección de cada columna, por posición (los nombres repetidos se pisan: el
    llamador usa nombres ya únicos)."""
    return {
        str(nombre): detect_column_type(df.iloc[:, i], cancel=cancel)
        for i, nombre in enumerate(df.columns)
    }


def _fallidas(original: np.ndarray, falla: np.ndarray) -> list[tuple[int, str]]:
    return [(int(i), str(original[i])) for i in np.flatnonzero(falla)]


def to_number(values: pd.Series, decimal_style: DecimalStyle | None = None) -> ConversionResult:
    """Convierte a float. Sin estilo pedido, lo detecta de la serie."""
    original = _como_texto(values)
    limpio = _strip(original)
    estilo = decimal_style or _estilo_de(limpio)
    vacio = (limpio == "").to_numpy(dtype=bool)
    pat_ent, pat_dec = _patrones_numero(estilo)
    ok = (_match(limpio, pat_ent) | _match(limpio, pat_dec)) & ~vacio
    if estilo == ",":
        normal = limpio.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    else:
        normal = limpio.str.replace(",", "", regex=False)
    numeros = pd.to_numeric(normal.where(ok), errors="coerce").astype("float64")
    return ConversionResult(
        values=pd.Series(numeros.to_numpy(), index=values.index),
        failed=_fallidas(original, ~ok & ~vacio),
    )


def to_date(values: pd.Series) -> ConversionResult:
    """Convierte a fecha o fecha-hora (datetime64)."""
    original = _como_texto(values)
    limpio = _strip(original)
    vacio = (limpio == "").to_numpy(dtype=bool)
    fechas, _ = _parse_fechas(limpio)
    ok = fechas.notna().to_numpy(dtype=bool)
    return ConversionResult(
        values=pd.Series(fechas.to_numpy(), index=values.index),
        failed=_fallidas(original, ~ok & ~vacio),
    )


def to_bool(values: pd.Series) -> ConversionResult:
    """Convierte sí/no, verdadero/falso, true/false a booleano (0 y 1 no valen: son enteros)."""
    original = _como_texto(values)
    limpio = _strip(original)
    vacio = (limpio == "").to_numpy(dtype=bool)
    bajo = limpio.str.lower()
    es_si = bajo.isin(_TRUE).to_numpy(dtype=bool)
    es_no = bajo.isin(_FALSE).to_numpy(dtype=bool)
    resultado = pd.array(np.where(es_si, True, False), dtype="boolean")
    resultado[~(es_si | es_no)] = pd.NA
    return ConversionResult(
        values=pd.Series(resultado, index=values.index),
        failed=_fallidas(original, ~(es_si | es_no) & ~vacio),
    )
