import json

import numpy as np
import pandas as pd
import pytest

from excel_tool.core.errors import Cancelado, ExcelToolError
from excel_tool.core.explore import (
    DistinctOptions,
    FilterCondition,
    filter_rows,
    get_page,
    rows_with_value,
    sort_rows,
    value_counts,
)
from excel_tool.core.io import read_table
from excel_tool.core.progress import CancelToken, ProgressEvent

STR = pd.StringDtype("pyarrow")


def _df(**columnas) -> pd.DataFrame:
    return pd.DataFrame({k: pd.array(v, dtype=STR) for k, v in columnas.items()})


@pytest.fixture(scope="module", params=["csv", "xlsx"])
def explorar(request, fixtures_generados):
    return read_table(fixtures_generados / "explorar" / f"para_explorar.{request.param}").df


@pytest.fixture(scope="module")
def perfil(fixtures_esperados):
    d = json.loads((fixtures_esperados / "perfil.json").read_text(encoding="utf-8"))
    return {c["posicion"] - 1: c for c in d["columnas_detalle"]}, d["filas"]


# ----------------------------------------------------------------- value_counts


def test_value_counts_contra_perfil(explorar, perfil):
    detalle, filas = perfil
    assert len(explorar) == filas
    for pos, col in enumerate(explorar.columns):
        r = value_counts(explorar, [col])
        esperado = detalle[pos]
        tabla = r.table
        hay_vacio = bool((explorar.iloc[:, pos] == "").any())
        assert r.distinct_total - int(hay_vacio) == esperado["distintos"], col
        assert r.rows_total == filas
        assert int(tabla["Cantidad"].sum()) == filas
        vacios = tabla.loc[tabla[col] == "(vacío)", "Cantidad"]
        assert (int(vacios.iloc[0]) if len(vacios) else 0) == esperado["vacios"], col
        assert tabla["Porcentaje"].sum() == pytest.approx(100.0)
        assert not r.truncated


def test_value_counts_region_normalizado_y_variantes(explorar):
    sin = value_counts(explorar, ["Región"])
    con = value_counts(explorar, ["Región"], DistinctOptions(normalize=True))
    assert con.distinct_total < sin.distinct_total
    assert con.variant_groups == []
    assert sin.variant_groups  # "Norte" / "NORTE" / "Norte "
    for grupo in sin.variant_groups:
        assert len({v.strip().lower() for v in grupo}) == 1 and len(grupo) > 1
    assert int(con.table["Cantidad"].sum()) == 200


def test_value_counts_orden_vacio_busqueda(explorar):
    r = value_counts(explorar, ["Región"], DistinctOptions(include_empty=False))
    assert "(vacío)" not in set(r.table["Región"])
    cant = r.table["Cantidad"].tolist()
    assert cant == sorted(cant, reverse=True)
    alfa = value_counts(explorar, ["Región"], DistinctOptions(sort="alfabetico"))
    nombres = [x.lower() for x in alfa.table["Región"] if x != "(vacío)"]
    assert nombres == sorted(nombres)
    assert alfa.table["Región"].iloc[-1] == "(vacío)"
    b = value_counts(explorar, ["Región"], DistinctOptions(search="NOR"))
    assert len(b.table) and all("nor" in x.lower() for x in b.table["Región"])
    assert b.distinct_total == sin_busqueda(explorar)


def sin_busqueda(df):
    return value_counts(df, ["Región"]).distinct_total


def test_value_counts_limit_truncated():
    df = _df(a=["x", "y", "z", "x", "x", "y"])
    r = value_counts(df, ["a"], DistinctOptions(limit=2))
    assert r.truncated and len(r.table) == 2 and r.distinct_total == 3
    assert r.table["a"].tolist() == ["x", "y"]
    assert r.table["Cantidad"].tolist() == [3, 2]
    assert r.table["Porcentaje"].tolist() == pytest.approx([50.0, 100 / 3])
    r = value_counts(df, ["a"], DistinctOptions(limit=3))
    assert not r.truncated


def test_value_counts_varias_columnas(explorar):
    r = value_counts(explorar, ["Región", "Activo"])
    assert list(r.table.columns) == ["Región", "Activo", "Cantidad", "Porcentaje"]
    assert int(r.table["Cantidad"].sum()) == len(explorar)
    esperado = explorar.groupby(["Región", "Activo"], sort=False).size()
    assert r.distinct_total == len(esperado)
    for _, fila in r.table.iterrows():
        reg = "" if fila["Región"] == "(vacío)" else fila["Región"]
        act = "" if fila["Activo"] == "(vacío)" else fila["Activo"]
        assert esperado[(reg, act)] == fila["Cantidad"]


def test_value_counts_normalize_representante_y_espacios():
    df = _df(r=["Norte", "NORTE", "Norte ", "Norte", "Sur", "", "  "])
    r = value_counts(df, ["r"], DistinctOptions(normalize=True))
    assert r.table["r"].tolist() == ["Norte", "(vacío)", "Sur"]
    assert r.table["Cantidad"].tolist() == [4, 2, 1]
    sin = value_counts(df, ["r"])
    assert sorted(map(sorted, sin.variant_groups)) == [["", "  "], ["NORTE", "Norte", "Norte "]]


def test_value_counts_casos_borde():
    vacio = _df(a=[])
    r = value_counts(vacio, ["a"])
    assert len(r.table) == 0 and r.rows_total == 0 and r.distinct_total == 0
    assert list(r.table.columns) == ["a", "Cantidad", "Porcentaje"]
    todo_vacio = _df(a=["", "", ""])
    r = value_counts(todo_vacio, ["a"])
    assert r.table["a"].tolist() == ["(vacío)"] and r.table["Porcentaje"].tolist() == [100.0]
    assert len(value_counts(todo_vacio, ["a"], DistinctOptions(include_empty=False)).table) == 0
    with pytest.raises(ExcelToolError):
        value_counts(todo_vacio, ["no_existe"])
    with pytest.raises(ExcelToolError):
        value_counts(todo_vacio, [])


def test_value_counts_progreso_y_cancelacion():
    df = _df(a=["1", "2"], b=["x", "y"])
    eventos: list[ProgressEvent] = []
    value_counts(df, ["a", "b"], progress=eventos.append)
    assert eventos and eventos[-1].done == eventos[-1].total
    assert any("a" in e.message for e in eventos)
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        value_counts(df, ["a"], cancel=token)
    token = CancelToken()

    def cancelar_al_primer_evento(_):
        token.cancel()

    with pytest.raises(Cancelado):
        value_counts(df, ["a", "b"], progress=cancelar_al_primer_evento, cancel=token)


# ----------------------------------------------------------------- rows_with_value


def test_rows_with_value(explorar):
    esperado = np.flatnonzero((explorar["Región"] == "Norte").to_numpy())
    assert rows_with_value(explorar, ["Región"], ("Norte",)).tolist() == esperado.tolist()
    nor = rows_with_value(explorar, ["Región"], ("Norte",), normalize=True)
    todas = explorar["Región"].str.strip().str.lower() == "norte"
    assert nor.tolist() == np.flatnonzero(todas.to_numpy()).tolist()
    assert len(nor) >= len(esperado)
    vac = rows_with_value(explorar, ["Vacía"], ("(vacío)",))
    assert len(vac) == 200
    dos = rows_with_value(explorar, ["Región", "Activo"], ("Norte", "Sí"))
    assert all(explorar.iloc[i]["Región"] == "Norte" for i in dos)
    assert len(rows_with_value(explorar, ["Región"], ("no hay",))) == 0
    with pytest.raises(ExcelToolError):
        rows_with_value(explorar, ["Región", "Activo"], ("Norte",))


# ----------------------------------------------------------------- filter_rows


TIPOS = {
    "ID": "entero",
    "Importe": "decimal",
    "Fecha": "fecha",
    "Mixta": "mixto",
    "Región": "texto",
    "Descripción": "texto",
}


def test_filtros_texto(explorar):
    col = explorar["Descripción"]
    ig = filter_rows(explorar, [FilterCondition("Descripción", "igual", "PEGAMENTO CHICO")], TIPOS)
    assert ig.tolist() == np.flatnonzero((col.str.lower() == "pegamento chico").to_numpy()).tolist()
    assert len(ig)
    di = filter_rows(
        explorar, [FilterCondition("Descripción", "distinto", "pegamento chico")], TIPOS
    )
    assert len(ig) + len(di) == 200
    co = filter_rows(explorar, [FilterCondition("Descripción", "contiene", "ÑANDÚ")], TIPOS)
    assert co.tolist() == np.flatnonzero(col.str.lower().str.contains("ñandú").to_numpy()).tolist()
    va = filter_rows(explorar, [FilterCondition("Activo", "vacio")], TIPOS)
    nv = filter_rows(explorar, [FilterCondition("Activo", "no_vacio")], TIPOS)
    assert len(va) == 18 and len(va) + len(nv) == 200


def test_filtros_numero_y_fecha(explorar):
    num = explorar["Importe"].str.replace(",", ".").replace("", np.nan).astype(float)
    may = filter_rows(explorar, [FilterCondition("Importe", "mayor", "5000,5")], TIPOS)
    assert may.tolist() == np.flatnonzero((num > 5000.5).to_numpy()).tolist()
    men = filter_rows(explorar, [FilterCondition("Importe", "menor", "5000,5")], TIPOS)
    assert men.tolist() == np.flatnonzero((num < 5000.5).to_numpy()).tolist()
    fec = pd.to_datetime(explorar["Fecha"].replace("", np.nan), format="%d/%m/%Y")
    may_f = filter_rows(explorar, [FilterCondition("Fecha", "mayor", "01/07/2024")], TIPOS)
    assert may_f.tolist() == np.flatnonzero((fec > pd.Timestamp("2024-07-01")).to_numpy()).tolist()
    assert len(may_f)


def test_filtro_y_entre_condiciones(explorar):
    a = FilterCondition("Importe", "mayor", "5000")
    b = FilterCondition("Región", "igual", "sur")
    r = filter_rows(explorar, [a, b], TIPOS)
    assert set(r) == set(filter_rows(explorar, [a], TIPOS)) & set(filter_rows(explorar, [b], TIPOS))
    assert filter_rows(explorar, [], TIPOS).tolist() == list(range(200))


def test_filtro_columna_mixta(explorar):
    r = filter_rows(explorar, [FilterCondition("Mixta", "mayor", "5000")], TIPOS)
    serie = pd.to_numeric(explorar["Mixta"], errors="coerce")
    assert r.tolist() == np.flatnonzero((serie > 5000).to_numpy()).tolist()
    assert len(r)
    # los textos ("s/d") y los vacíos no pasan nunca
    assert all(explorar["Mixta"].iloc[i].isdigit() for i in r)
    # valor de filtro que no es número: se compara como texto
    t = filter_rows(explorar, [FilterCondition("Mixta", "mayor", "s")], TIPOS)
    assert all(explorar["Mixta"].iloc[i].lower() > "s" for i in t)


def test_filtro_valor_invalido():
    df = _df(n=["1", "2", "3"])
    with pytest.raises(ExcelToolError, match="número"):
        filter_rows(df, [FilterCondition("n", "mayor", "abc")], {"n": "entero"})
    df = _df(f=["01/01/2024"])
    with pytest.raises(ExcelToolError, match="fecha"):
        filter_rows(df, [FilterCondition("f", "mayor", "hola")], {"f": "fecha"})
    with pytest.raises(ExcelToolError):
        filter_rows(df, [FilterCondition("zzz", "igual", "x")], {})


def test_filtro_mayor_texto_y_vacios_no_pasan():
    df = _df(t=["b", "a", "", "C", "  "])
    r = filter_rows(df, [FilterCondition("t", "mayor", "a")], {"t": "texto"})
    assert r.tolist() == [0, 3]
    r = filter_rows(df, [FilterCondition("t", "menor", "c")], {"t": "texto"})
    assert r.tolist() == [0, 1]


def test_filtro_0_filas():
    df = _df(a=[])
    assert len(filter_rows(df, [FilterCondition("a", "contiene", "x")], {})) == 0
    assert len(filter_rows(df, [FilterCondition("a", "mayor", "1")], {"a": "entero"})) == 0
    assert len(sort_rows(df, pd.RangeIndex(0), "a", True, "texto")) == 0


# ----------------------------------------------------------------- sort_rows


def test_orden_numerico_estable_y_al_final():
    df = _df(n=["10", "2", "", "x", "2", "1.000,5", "-3"])
    todas = pd.RangeIndex(len(df))
    asc = sort_rows(df, todas, "n", True, "mixto")
    assert asc.tolist() == [6, 1, 4, 0, 5, 3, 2]  # -3, 2, 2 (estable), 10, 1000,5, texto, vacío
    desc = sort_rows(df, todas, "n", False, "mixto")
    assert desc.tolist() == [5, 0, 1, 4, 6, 3, 2]  # descendente: iguales siguen estables
    num = sort_rows(df, todas, "n", True, "decimal")
    assert num.tolist() == [6, 1, 4, 0, 5, 3, 2]


def test_orden_fecha_texto_y_subconjunto(explorar):
    r = sort_rows(explorar, pd.RangeIndex(200), "Fecha", True, "fecha")
    fechas = pd.to_datetime(explorar["Fecha"].iloc[r].replace("", np.nan), format="%d/%m/%Y")
    validas = fechas.dropna()
    assert validas.is_monotonic_increasing
    assert fechas.iloc[len(validas) :].isna().all()  # vacíos al final
    d = sort_rows(explorar, pd.RangeIndex(200), "Fecha", False, "fecha")
    fd = pd.to_datetime(explorar["Fecha"].iloc[d].replace("", np.nan), format="%d/%m/%Y")
    assert fd.dropna().is_monotonic_decreasing and fd.iloc[len(fd.dropna()) :].isna().all()
    sub = filter_rows(explorar, [FilterCondition("Región", "igual", "Sur")], TIPOS)
    t = sort_rows(explorar, sub, "Descripción", True, "texto")
    assert sorted(t.tolist()) == sorted(sub.tolist())
    nombres = [explorar["Descripción"].iloc[i].lower() for i in t]
    llenos = [n for n in nombres if n]
    assert llenos == sorted(llenos) and nombres[: len(llenos)] == llenos


def test_orden_texto_ignora_mayusculas():
    df = _df(t=["b", "A", "a", "B"])
    assert sort_rows(df, pd.RangeIndex(4), "t", True, "texto").tolist() == [1, 2, 0, 3]
    assert sort_rows(df, pd.RangeIndex(4), "t", False, "texto").tolist() == [0, 3, 1, 2]


# ----------------------------------------------------------------- get_page


def test_get_page(explorar):
    todas = pd.RangeIndex(len(explorar))
    p0 = get_page(explorar, todas, 0, 50)
    assert list(p0.columns) == ["N.º fila", *explorar.columns]
    assert p0["N.º fila"].tolist() == list(range(1, 51))
    p3 = get_page(explorar, todas, 3, 50)
    assert p3["N.º fila"].tolist() == list(range(151, 201))
    assert p3["ID"].tolist() == explorar["ID"].iloc[150:].tolist()
    filtradas = pd.Index([5, 2, 9])
    p = get_page(explorar, filtradas, 0, 2)
    assert p["N.º fila"].tolist() == [6, 3]
    assert p["ID"].tolist() == explorar["ID"].iloc[[5, 2]].tolist()
    assert get_page(explorar, filtradas, 1, 2)["N.º fila"].tolist() == [10]


def test_get_page_fuera_de_rango_y_vacio(explorar):
    p = get_page(explorar, pd.RangeIndex(200), 99, 50)
    assert len(p) == 0 and list(p.columns) == ["N.º fila", *explorar.columns]
    assert len(get_page(explorar, pd.Index([], dtype="int64"), 0)) == 0
    assert len(get_page(explorar.iloc[:0], pd.RangeIndex(0), 0)) == 0
    assert len(get_page(explorar, pd.RangeIndex(200), -1, 50)) == 0
    with pytest.raises(ValueError):
        get_page(explorar, pd.RangeIndex(200), 0, 0)


def test_get_page_no_modifica_la_tabla(explorar):
    columnas = list(explorar.columns)
    get_page(explorar, pd.RangeIndex(200), 0, 10)
    assert list(explorar.columns) == columnas
