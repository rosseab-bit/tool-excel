"""Tests de core/types.py: detección de tipos y conversiones."""

from __future__ import annotations

import json
import time
import warnings

import pandas as pd
import pytest

from excel_tool.core.io import read_table
from excel_tool.core.types import (
    classify_value,
    detect_column_type,
    detect_types,
    to_bool,
    to_date,
    to_number,
)


def serie(*valores: str) -> pd.Series:
    return pd.Series(list(valores), dtype=object)


# --- classify_value -------------------------------------------------------------------


@pytest.mark.parametrize(
    ("valor", "estilo", "esperado"),
    [
        ("0", ",", "entero"),
        ("00123", ",", "texto"),
        ("007", ".", "texto"),
        ("-5", ",", "entero"),
        ("1.234", ",", "entero"),
        ("1.234.567", ",", "entero"),
        ("1.23", ",", "texto"),
        ("1234,56", ",", "decimal"),
        ("1.234,56", ",", "decimal"),
        ("0,5", ",", "decimal"),
        ("1234.56", ",", "texto"),
        ("1234.56", ".", "decimal"),
        ("1,234.56", ".", "decimal"),
        ("1,234", ".", "entero"),
        ("1.234", ".", "decimal"),
        ("", ",", "vacia"),
        ("   ", ",", "vacia"),
        ("  42  ", ",", "entero"),
        ("03/07/2024", ",", "fecha"),
        ("3/7/2024", ",", "fecha"),
        ("3-7-2024", ",", "fecha"),
        ("03/07/2024 13:05", ",", "fecha_hora"),
        ("3/7/2024 13:05:59", ",", "fecha_hora"),
        ("31/02/2024", ",", "texto"),
        ("29/02/2023", ",", "texto"),
        ("29/02/2024", ",", "fecha"),
        ("03/07/2024 25:00", ",", "texto"),
        ("03/07/2024 12:60", ",", "texto"),
        ("1/2-2024", ",", "texto"),
        ("03/07/24", ",", "texto"),
        ("Sí", ",", "si_no"),
        ("si", ",", "si_no"),
        (" NO ", ",", "si_no"),
        ("Verdadero", ",", "si_no"),
        ("FALSE", ",", "si_no"),
        ("true", ",", "si_no"),
        ("1", ",", "entero"),
        ("0", ".", "entero"),
        ("hola", ",", "texto"),
        ("s/d", ",", "texto"),
    ],
)
def test_classify_value(valor, estilo, esperado):
    assert classify_value(valor, estilo) == esperado


# --- detect_column_type ---------------------------------------------------------------


def test_columna_vacia():
    d = detect_column_type(serie("", "  ", ""))
    assert d.detected == "vacia"
    assert d.breakdown == {}
    assert d.empty == 3
    assert d.decimal_style is None


def test_columna_sin_filas():
    assert detect_column_type(serie()).detected == "vacia"


def test_columna_una_sola_clase():
    d = detect_column_type(serie("1", "2", "", "3"))
    assert d.detected == "entero"
    assert d.breakdown == {"entero": 3}
    assert d.empty == 1
    assert d.decimal_style == ","


def test_entero_mas_decimal_es_decimal():
    d = detect_column_type(serie("1", "2,5", "1.000"))
    assert d.detected == "decimal"
    assert d.breakdown == {"entero": 2, "decimal": 1}


def test_fecha_mas_fecha_hora_es_fecha_hora():
    d = detect_column_type(serie("03/07/2024", "03/07/2024 10:00"))
    assert d.detected == "fecha_hora"
    assert d.breakdown == {"fecha": 1, "fecha_hora": 1}
    assert d.decimal_style is None


def test_otras_combinaciones_son_mixtas():
    d = detect_column_type(serie("1", "a", "a", "03/07/2024", "si"))
    assert d.detected == "mixto"
    assert d.breakdown == {"entero": 1, "texto": 2, "fecha": 1, "si_no": 1}
    assert detect_column_type(serie("1", "si")).detected == "mixto"
    assert detect_column_type(serie("1,5", "03/07/2024")).detected == "mixto"


def test_ceros_a_la_izquierda_es_texto():
    assert detect_column_type(serie("00123", "00456")).detected == "texto"
    assert detect_column_type(serie("0", "10")).detected == "entero"


def test_cero_uno_son_enteros():
    assert detect_column_type(serie("0", "1", "1", "0")).detected == "entero"


def test_fechas_imposibles_son_texto():
    d = detect_column_type(serie("31/02/2024", "03/07/2024"))
    assert d.detected == "mixto"
    assert d.breakdown == {"texto": 1, "fecha": 1}


def test_si_no_columna():
    d = detect_column_type(serie("Sí", "no", " VERDADERO ", "falso", "True", "si"))
    assert d.detected == "si_no"
    assert d.breakdown == {"si_no": 6}


def test_estilo_ambiguo_es_argentino():
    d = detect_column_type(serie("1.234", "5"))
    assert d.decimal_style == ","
    assert d.detected == "entero"


def test_estilo_punto_si_hay_evidencia():
    d = detect_column_type(serie("1234.56", "1,234.56", "7"))
    assert d.decimal_style == "."
    assert d.detected == "decimal"
    assert d.breakdown == {"decimal": 2, "entero": 1}


def test_estilo_decidido_por_columna_no_por_celda():
    # "1.234" sería entero en estilo argentino, pero la columna es de estilo punto.
    d = detect_column_type(serie("1.234", "2.5", "3.75"))
    assert d.decimal_style == "."
    assert d.detected == "decimal"


def test_estilo_coma_gana_si_hay_mayoria():
    d = detect_column_type(serie("1,5", "2,5", "3.5"))
    assert d.decimal_style == ","
    assert d.breakdown == {"decimal": 2, "texto": 1}


def test_acepta_nulos_de_pandas():
    d = detect_column_type(pd.Series(["1", None, float("nan"), "2"], dtype=object))
    assert d.detected == "entero"
    assert d.empty == 2


def test_detect_types_por_columna():
    df = pd.DataFrame({"a": ["1", "2"], "b": ["x", ""]}, dtype=str)
    r = detect_types(df)
    assert list(r) == ["a", "b"]
    assert r["a"].detected == "entero"
    assert r["b"].detected == "texto"
    assert r["b"].empty == 1


def test_detect_types_en_xlsx(tmp_path):
    ruta = tmp_path / "t.xlsx"
    pd.DataFrame({"n": ["1", "2,5"], "f": ["03/07/2024", "04/07/2024"]}).to_excel(ruta, index=False)
    df = pd.read_excel(ruta, dtype=str, keep_default_na=False, engine="openpyxl")
    r = detect_types(df)
    assert r["n"].detected == "decimal"
    assert r["f"].detected == "fecha"


# --- conversiones ---------------------------------------------------------------------


def test_to_number_estilo_argentino():
    r = to_number(serie("1.234,56", "10", "", "abc", "  7,5 ", "1234.5"), ",")
    v = r.values.tolist()
    assert v[0] == 1234.56
    assert v[1] == 10.0
    assert pd.isna(v[2])
    assert pd.isna(v[3])
    assert v[4] == 7.5
    assert r.failed == [(3, "abc"), (5, "1234.5")]
    assert r.values.dtype == "float64"


def test_to_number_estilo_punto_y_deteccion():
    r = to_number(serie("1,234.56", "0.5", "x"), ".")
    assert r.values.tolist()[:2] == [1234.56, 0.5]
    assert r.failed == [(2, "x")]
    auto = to_number(serie("1234.56", "0.5"))
    assert auto.values.tolist() == [1234.56, 0.5]
    assert auto.failed == []


def test_to_number_ambiguo_es_argentino_y_conserva_indice():
    s = pd.Series(["1.234", "5"], index=[10, 20], dtype=object)
    r = to_number(s)
    assert r.values.tolist() == [1234.0, 5.0]
    assert list(r.values.index) == [10, 20]


def test_to_number_ceros_a_la_izquierda_fallan():
    r = to_number(serie("00123", "0"))
    assert r.failed == [(0, "00123")]
    assert r.values.tolist()[1] == 0.0


def test_to_number_todo_vacio():
    r = to_number(serie("", " "))
    assert r.failed == []
    assert r.values.isna().all()


def test_to_date():
    r = to_date(serie("03/07/2024", "3-7-2024", "03/07/2024 13:05:09", "", "31/02/2024", "hola"))
    v = r.values
    assert v.iloc[0] == pd.Timestamp(2024, 7, 3)
    assert v.iloc[1] == pd.Timestamp(2024, 7, 3)
    assert v.iloc[2] == pd.Timestamp(2024, 7, 3, 13, 5, 9)
    assert v.iloc[3:].isna().all()
    assert r.failed == [(4, "31/02/2024"), (5, "hola")]
    assert str(v.dtype).startswith("datetime64")


def test_to_date_sin_fechas():
    r = to_date(serie("a", ""))
    assert r.failed == [(0, "a")]
    assert r.values.isna().all()


def test_to_bool():
    r = to_bool(serie("Sí", " no ", "VERDADERO", "falso", "", "1", "tal vez", "true", "False"))
    v = r.values.tolist()
    assert v[:4] == [True, False, True, False]
    assert v[4] is pd.NA
    assert v[5] is pd.NA
    assert v[7:] == [True, False]
    assert r.failed == [(5, "1"), (6, "tal vez")]
    assert str(r.values.dtype) == "boolean"


# --- contra perfil.json ---------------------------------------------------------------


@pytest.fixture(scope="module")
def perfil_explorar(fixtures_generados, fixtures_esperados):
    df = pd.read_csv(
        fixtures_generados / "explorar" / "para_explorar.csv",
        sep=";",
        encoding="utf-8-sig",
        dtype=str,
        keep_default_na=False,
    )
    perfil = json.loads((fixtures_esperados / "perfil.json").read_text(encoding="utf-8"))
    return df, perfil["columnas_detalle"]


def test_perfil_json(perfil_explorar):
    df, esperadas = perfil_explorar
    assert len(df.columns) == len(esperadas)
    for i, esperada in enumerate(esperadas):
        d = detect_column_type(df.iloc[:, i])
        etiqueta = f"columna {i + 1} ({esperada['nombre']!r})"
        assert d.detected == esperada["tipo"], etiqueta
        assert d.empty == esperada["vacios"], etiqueta
        if esperada["desglose_tipos"] is not None:
            assert d.breakdown == esperada["desglose_tipos"], etiqueta


# --- rendimiento ----------------------------------------------------------------------


def test_rendimiento_2m_filas(fixtures_grandes):
    """Objetivo de la Fase 1: <= 15 s. Se reporta; solo falla si pasa del doble."""
    df = read_table(fixtures_grandes / "rendimiento_2M.csv").df  # camino real de la app
    inicio = time.perf_counter()
    r = detect_types(df)
    segundos = time.perf_counter() - inicio
    print(f"\ndetect_types 2M filas: {segundos:.1f} s (objetivo 15 s, tope 30 s)")
    if segundos > 15:
        warnings.warn(f"detect_types 2M filas tardo {segundos:.1f} s (objetivo 15 s)", stacklevel=1)
    assert segundos <= 30
    assert r["ID"].detected == "entero"
    assert r["Fecha y hora"].detected == "fecha_hora"
