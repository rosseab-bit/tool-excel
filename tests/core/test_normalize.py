"""Tests de core/normalize (F4-01)."""

from __future__ import annotations

import json

import pandas as pd
import pytest

from excel_tool.core.io import WriteOptions, read_table, write_table
from excel_tool.core.normalize import KeyOptions, match_key_columns, normalize_keys

_STR = pd.StringDtype("pyarrow")


def _serie(valores):
    return pd.Series(valores, dtype=_STR)


def _norm(valores, **kw):
    return normalize_keys(_serie(valores), KeyOptions(**kw)).tolist()


def test_trim_espacios_duros_tabs_y_dobles():
    assert _norm([" 12 ", "\xa012\xa0", "\t12\t", "a  b", "a\xa0 \tb", "   "]) == [
        "12",
        "12",
        "12",
        "a b",
        "a b",
        "",
    ]


def test_sin_trim_deja_espacios():
    assert _norm([" 12 "], trim=False) == [" 12 "]


def test_ignore_case_y_tildes_se_conservan():
    assert _norm(["ABC", "Ñandú"]) == ["abc", "ñandú"]
    assert _norm(["ABC"], ignore_case=False) == ["ABC"]


def test_vacios_y_nulos_quedan_vacios():
    s = pd.Series(["", None, "x"], dtype=_STR)
    out = normalize_keys(s, KeyOptions(strip_leading_zeros=True))
    assert out.tolist() == ["", "", "x"]
    assert out.dtype == _STR
    assert out.index.equals(s.index)


def test_ceros_no_se_sacan_por_defecto():
    assert _norm(["00012", "12"]) == ["00012", "12"]


def test_strip_ceros_casos_borde():
    out = _norm(
        ["00012", "0", "000", "00A1", "0A", " 007 ", "12", "0.5", "00", "-007", "٠٠١"],
        strip_leading_zeros=True,
    )
    assert out == ["12", "0", "0", "00a1", "0a", "7", "12", "0.5", "0", "-007", "٠٠١"]


def test_strip_ceros_sin_trim_no_toca_valor_con_espacios():
    assert _norm([" 007"], trim=False, strip_leading_zeros=True) == [" 007"]


def test_123_punto_0_no_se_convierte():
    assert _norm(["123.0", "123"], strip_leading_zeros=True) == ["123.0", "123"]


def test_mantiene_indice():
    s = pd.Series(["A", "b"], index=[10, 20], dtype=_STR, name="k")
    out = normalize_keys(s, KeyOptions())
    assert list(out.index) == [10, 20] and out.name == "k"


def test_match_key_columns():
    assert match_key_columns(["ID cliente", "Región", "Monto"], ["id  Cliente", "REGION", "x"]) == [
        ("ID cliente", "id  Cliente"),
        ("Región", "REGION"),
    ]
    assert match_key_columns(["Nombre"], ["Apellido"]) == []
    assert match_key_columns([], []) == []


def test_match_key_columns_ambiguo_no_adivina():
    assert match_key_columns(["Id", "ID"], ["id"]) == []
    assert match_key_columns(["Id"], ["id", "ID "]) == []
    assert match_key_columns(["Id", "Otro"], ["id", "OTRO"]) == [("Id", "id"), ("Otro", "OTRO")]


@pytest.fixture(scope="module")
def _esperado(fixtures_esperados):
    ruta = fixtures_esperados / "escenario_3" / "diagnostico.json"
    return json.loads(ruta.read_text(encoding="utf-8"))["cobranzas"]


def _coincidencias(clientes, cobranzas, opciones):
    p = normalize_keys(clientes["ID cliente"], opciones)
    o = normalize_keys(cobranzas["Nro Cliente"], opciones)
    return int(p.isin(set(o)).sum())


@pytest.fixture(params=["xlsx", "csv"])
def _tablas(request, fixtures_generados, tmp_path):
    base = fixtures_generados / "escenario_3"
    clientes = read_table(base / "clientes.xlsx").df
    cob = read_table(base / "cobranzas.xlsx").df
    if request.param == "csv":
        destino = tmp_path / "cobranzas.csv"
        write_table(cob, destino, WriteOptions(format="csv_excel"))
        cob = read_table(destino).df
    return clientes, cob


def test_escenario_solo_recortar(_tablas, _esperado):
    clientes, cob = _tablas
    esperado = _esperado["solo_recortar_espacios"]["claves_con_coincidencia"]
    assert _coincidencias(clientes, cob, KeyOptions()) == esperado == 5


def test_escenario_sin_ceros(_tablas, _esperado):
    clientes, cob = _tablas
    esperado = _esperado["sin_ceros_a_la_izquierda"]["claves_con_coincidencia"]
    assert _coincidencias(clientes, cob, KeyOptions(strip_leading_zeros=True)) == esperado == 27
