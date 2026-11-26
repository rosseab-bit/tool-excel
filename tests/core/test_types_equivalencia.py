"""F4-05: la conversión por valores distintos da lo mismo que la versión vieja.

La referencia es una copia de la lógica anterior a F4-05 (regex de Python, `str.strip` y una
pasada por todas las filas), escrita aparte para que no comparta código con `core/types.py`.
"""

from __future__ import annotations

import re
import time
from datetime import datetime

import numpy as np
import pandas as pd
import pytest

from excel_tool.core.io import read_table
from excel_tool.core.types import to_bool, to_date, to_number

_BASE = r"(?:0|[1-9]\d*|[1-9]\d{0,2}(?:%s\d{3})+)"


def _regex(estilo):
    miles, dec = (r"\.", ",") if estilo == "," else (",", r"\.")
    base = _BASE % miles
    return re.compile(rf"-?{base}"), re.compile(rf"-?{base}{dec}\d+")


def _texto(v):
    return "" if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NA else str(v)


def _estilo_viejo(limpios):
    con_sep = [v for v in limpios if "." in v or "," in v]
    if not con_sep:
        return ","
    cuentas = []
    for estilo in (",", "."):
        e, d = _regex(estilo)
        cuentas.append(sum(1 for v in con_sep if e.fullmatch(v) or d.fullmatch(v)))
    return "," if cuentas[0] == len(con_sep) else ("." if cuentas[1] > cuentas[0] else ",")


def numero_viejo(values, estilo=None):
    originales = [_texto(v) for v in values]
    limpios = [v.strip() for v in originales]
    estilo = estilo or _estilo_viejo(limpios)
    e, d = _regex(estilo)
    out, failed = [], []
    for i, (o, v) in enumerate(zip(originales, limpios, strict=True)):
        if v == "":
            out.append(np.nan)
        elif e.fullmatch(v) or d.fullmatch(v):
            out.append(
                float(v.replace(".", "").replace(",", ".") if estilo == "," else v.replace(",", ""))
            )
        else:
            out.append(np.nan)
            failed.append((i, o))
    return np.array(out, dtype="float64"), failed


_FORMATOS = (
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    "%d-%m-%Y %H:%M:%S",
    "%d-%m-%Y %H:%M",
    "%d-%m-%Y",
)


def fecha_vieja(values):
    out, failed = [], []
    for i, v in enumerate(values):
        o = _texto(v)
        s = o.strip()
        valor = pd.NaT
        if s:
            for f in _FORMATOS:
                if ("/" in s) != f.startswith("%d/"):
                    continue
                if s.count(":") != f.count(":") or (":" in s) != (":" in f):
                    continue
                try:
                    valor = pd.Timestamp(datetime.strptime(s, f))
                    break
                except ValueError:
                    continue
            if valor is pd.NaT:
                failed.append((i, o))
        out.append(valor)
    return pd.Series(out, dtype="datetime64[ns]"), failed


def bool_viejo(values):
    si, no = ("sí", "si", "verdadero", "true"), ("no", "falso", "false")
    out, failed = [], []
    for i, v in enumerate(values):
        o = _texto(v)
        s = o.strip().lower()
        if s in si:
            out.append(True)
        elif s in no:
            out.append(False)
        else:
            out.append(pd.NA)
            if o.strip():
                failed.append((i, o))
    return pd.array(out, dtype="boolean"), failed


NUMEROS = [
    "1.234,56", "10", "", "abc", "  7,5 ", "1234.5", "-3", "0", "007", "1.234.567", "1,5",
    "1.5", "12,", ",5", "1 000", "\t8,25\n", "-0,5", "1.234,56", "10", "   ", "9.999,99", "1e5",
    "٣", "NaN", "inf", "--1", "1.23", "1,234",
]  # fmt: skip
NUMEROS_INGLES = [
    "1,234.56",
    "10",
    "",
    "abc",
    "7.5",
    "1,234,567",
    "1234.5",
    "3.",
    "1.234,5",
    "  8 ",
]
FECHAS = [
    "31/12/2024", "01/02/2023 10:30", "03/07/2024 23:59:59", "31/02/2024", "25:00", "", "  ",
    "15-06-2022", "15-06-2022 08:00", "2024-01-01", "abc", " 05/05/2025 ", "29/02/2024",
    "29/02/2023", "1/1/2020", "31/12/2024", "12/31/2024", "10/10/2010 24:00",
]  # fmt: skip
BOOLS = ["Sí", "no", "TRUE", "falso", "", "0", "1", "  si ", "Verdadero", "xx", "NO", " "]


def _mismo_numero(serie, estilo=None):
    nuevo = to_number(serie, estilo)
    valores, failed = numero_viejo(serie.tolist(), estilo)
    np.testing.assert_array_equal(nuevo.values.to_numpy(), valores)
    assert nuevo.failed == failed
    assert list(nuevo.values.index) == list(serie.index)


def _series(valores):
    return [
        pd.Series(valores, dtype=object),
        pd.Series(valores, dtype="string[pyarrow]"),
        pd.Series(valores, dtype=object, index=range(100, 100 + len(valores))),
    ]


@pytest.mark.parametrize("estilo", [",", ".", None])
def test_numero_equivale_a_la_version_vieja(estilo):
    for s in _series(NUMEROS * 3):
        _mismo_numero(s, estilo)


@pytest.mark.parametrize("estilo", [",", ".", None])
def test_numero_estilo_ingles(estilo):
    for s in _series(NUMEROS_INGLES * 2):
        _mismo_numero(s, estilo)


def test_numero_estilo_detectado_pesa_por_repeticion():
    """Pocos valores distintos 'ingleses' muy repetidos ganan a muchos 'argentinos' únicos."""
    valores = ["1,234.5"] * 50 + ["1.234,5", "2.345,6", "3.456,7"] + ["1,5"] * 2
    _mismo_numero(pd.Series(valores, dtype=object))


def test_vacias_y_con_nulos():
    s = pd.Series([None, "5", np.nan, "x", pd.NA], dtype=object)
    _mismo_numero(s, ",")
    assert to_number(pd.Series([], dtype=object)).failed == []


def test_fecha_equivale_a_la_version_vieja():
    for s in _series(FECHAS * 3):
        nuevo = to_date(s)
        valores, failed = fecha_vieja(s.tolist())
        pd.testing.assert_series_equal(
            nuevo.values.reset_index(drop=True), valores, check_names=False
        )
        assert nuevo.failed == failed


def test_bool_equivale_a_la_version_vieja():
    for s in _series(BOOLS * 3):
        nuevo = to_bool(s)
        valores, failed = bool_viejo(s.tolist())
        assert nuevo.values.tolist() == list(valores)
        assert nuevo.failed == failed


def test_bloques_chicos_dan_lo_mismo(monkeypatch):
    """Los distintos se convierten de a bloques: con bloques de 3 da igual."""
    from excel_tool.core import types

    monkeypatch.setattr(types, "BLOQUE_UNICOS", 3)
    _mismo_numero(pd.Series(NUMEROS * 2, dtype=object))
    nuevo = to_date(pd.Series(FECHAS, dtype=object))
    valores, failed = fecha_vieja(FECHAS)
    pd.testing.assert_series_equal(nuevo.values, valores, check_names=False)
    assert nuevo.failed == failed


@pytest.mark.parametrize("nombre", ["para_explorar.csv", "para_explorar.xlsx"])
def test_fixtures_explorar_todas_las_columnas(fixtures_generados, nombre):
    """CSV y xlsx: cada columna del archivo de prueba (mixtas e inválidas incluidas)."""
    df = read_table(fixtures_generados / "explorar" / nombre).df
    for i in range(len(df.columns)):
        col = df.iloc[:, i]
        _mismo_numero(col)
        nuevo = to_date(col)
        valores, failed = fecha_vieja(col.tolist())
        pd.testing.assert_series_equal(
            nuevo.values.reset_index(drop=True), valores, check_names=False
        )
        assert nuevo.failed == failed
        b, fb = bool_viejo(col.tolist())
        nb = to_bool(col)
        assert nb.values.tolist() == list(b)
        assert nb.failed == fb


def test_no_es_mas_lento_con_valores_todos_distintos():
    """Sin repeticiones factorizar no ayuda: no puede costar más del doble que la versión vieja."""
    n = 60_000
    s = pd.Series([f"{i}.{i % 1000:03d},{i % 97}" for i in range(n)], dtype=object)
    inicio = time.perf_counter()
    to_number(s, ",")
    nuevo = time.perf_counter() - inicio
    inicio = time.perf_counter()
    numero_viejo(s.tolist(), ",")  # referencia de Python puro (más lenta que la vieja de pandas)
    assert nuevo < 2 * (time.perf_counter() - inicio) + 0.5
