from collections import Counter

import numpy as np
import pandas as pd
import pytest

from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io import read_table
from excel_tool.core.split import SplitSpec, plan_split, split_dataframe

STR = pd.StringDtype("pyarrow")


def _df(**columnas) -> pd.DataFrame:
    return pd.DataFrame({k: pd.array(v, dtype=STR) for k, v in columnas.items()})


@pytest.fixture(scope="module", params=["csv", "xlsx"])
def tabla(request, fixtures_generados):
    return read_table(fixtures_generados / "explorar" / f"para_explorar.{request.param}").df


def _partes(df, spec, base="datos", max_rows=None):
    plan = plan_split(df, spec, base_name=base, max_rows_per_part=max_rows)
    reales = list(split_dataframe(df, spec, base_name=base, max_rows_per_part=max_rows))
    assert plan == [p for p, _ in reales]  # mismo plan, mismo orden
    for p, d in reales:
        assert p.rows == len(d)
    return reales


def _ids(partes) -> list:
    return [x for _, d in partes for x in d.iloc[:, 0]]


def test_rows(tabla):
    partes = _partes(tabla, SplitSpec("rows", rows_per_part=30))
    assert sum(p.rows for p, _ in partes) == len(tabla)
    assert all(p.rows <= 30 for p, _ in partes)
    assert _ids(partes) == list(tabla.iloc[:, 0])  # mismo orden
    assert partes[0][0].name == f"datos_parte_01_de_{len(partes):02d}"
    assert (partes[0][0].first_row, partes[0][0].last_row) == (1, 30)
    assert partes[-1][0].last_row == len(tabla)


def test_rows_respeta_limite_maximo(tabla):
    partes = _partes(tabla, SplitSpec("rows", rows_per_part=1000), max_rows=40)
    assert all(p.rows <= 40 for p, _ in partes)
    assert sum(p.rows for p, _ in partes) == len(tabla)


def test_rows_sin_cantidad_da_error(tabla):
    with pytest.raises(ExcelToolError):
        plan_split(tabla, SplitSpec("rows"), base_name="x")


def test_sheets_nombres_validos(tabla):
    base = "Un nombre bastante largo: para [hoja]"
    partes = _partes(tabla, SplitSpec("sheets", rows_per_part=50), base=base)
    nombres = [p.name for p, _ in partes]
    assert len(set(n.lower() for n in nombres)) == len(nombres)
    for n in nombres:
        assert len(n) <= 31 and not set("[]:*?/\\") & set(n)
    assert nombres[0].endswith(f"_parte_01_de_{len(partes):02d}")
    assert sum(p.rows for p, _ in partes) == len(tabla)


def test_keep_groups_no_corta_grupos(tabla):
    col = "Región"
    partes = _partes(tabla, SplitSpec("rows_keep_groups", rows_per_part=60, group_column=col))
    assert sum(p.rows for p, _ in partes) == len(tabla)
    vistos = [set(d[col]) for _, d in partes]
    for i, a in enumerate(vistos):
        for b in vistos[i + 1 :]:
            assert not a & b
    # el orden dentro del grupo es el original
    ids = _ids(partes)
    assert sorted(ids) == sorted(tabla.iloc[:, 0])


def test_keep_groups_no_contiguos_y_grupo_grande():
    df = _df(id=[str(i) for i in range(10)], g=list("ABABAACABA"))
    # A: filas 0,2,4,5,7,9 (6) - B: 1,3,8 (3) - C: 6 (1). Orden de primera aparición.
    partes = _partes(df, SplitSpec("rows_keep_groups", rows_per_part=4, group_column="g"))
    assert [p.rows for p, _ in partes] == [6, 4]  # A supera el límite: va sola
    assert partes[0][1]["id"].tolist() == ["0", "2", "4", "5", "7", "9"]
    assert partes[1][1]["id"].tolist() == ["1", "3", "8", "6"]
    assert [(p.first_row, p.last_row) for p, _ in partes] == [(1, 10), (2, 9)]
    partes = _partes(df, SplitSpec("rows_keep_groups", rows_per_part=7, group_column="g"))
    assert [p.rows for p, _ in partes] == [6, 4]
    partes = _partes(df, SplitSpec("rows_keep_groups", rows_per_part=10, group_column="g"))
    assert [p.rows for p, _ in partes] == [10]


def test_by_value(tabla):
    col = "Región"
    partes = _partes(tabla, SplitSpec("by_value", value_columns=[col]))
    assert sum(p.rows for p, _ in partes) == len(tabla)
    valores = [p.value[0] for p, _ in partes]
    assert valores == sorted(valores, key=lambda v: (v == "", v.casefold()))
    for p, d in partes:
        assert set(d[col]) == {p.value[0]}
        assert p.first_row == tabla.index[tabla[col] == p.value[0]][0] + 1
    filas = Counter(tuple(f) for _, d in partes for f in d.itertuples(index=False))
    assert filas == Counter(tuple(f) for f in tabla.itertuples(index=False))


def test_by_value_subdivide_parte_grande(tabla):
    col = "Región"
    mayor = int(tabla[col].value_counts().max())
    limite = max(1, mayor // 2)
    partes = _partes(tabla, SplitSpec("by_value", value_columns=[col], rows_per_part=limite))
    assert sum(p.rows for p, _ in partes) == len(tabla)
    assert all(p.rows <= limite for p, _ in partes)
    assert any("_parte_1_de_" in p.name for p, _ in partes)
    partes2 = _partes(tabla, SplitSpec("by_value", value_columns=[col]), max_rows=limite)
    assert [p for p, _ in partes2] == [p for p, _ in partes]


def test_by_value_orden_vacio_y_varias_columnas():
    df = _df(
        a=["b", "", "a", "B", "a", "b", "a"],
        b=["x", "x", "y", "x", "x", "x", ""],
        n=list("1234567"),
    )
    partes = _partes(df, SplitSpec("by_value", value_columns=["a"], name_pattern="{valor}"))
    assert [p.value for p, _ in partes] == [("a",), ("b",), ("B",), ("",)] or [
        p.value for p, _ in partes
    ] == [("a",), ("B",), ("b",), ("",)]
    assert partes[-1][0].name == "vacio"
    assert partes[0][1]["n"].tolist() == ["3", "5", "7"]  # orden original dentro de la parte
    dos = _partes(df, SplitSpec("by_value", value_columns=["a", "b"]))
    assert sum(p.rows for p, _ in dos) == 7
    assert dos[0][0].value == ("a", "x")
    assert dos[-1][0].value[0] == ""
    assert dos[0][0].name == "datos_a - x"


def test_by_value_columna_inexistente():
    with pytest.raises(ExcelToolError):
        plan_split(_df(a=["1"]), SplitSpec("by_value", value_columns=["z"]), base_name="d")
    with pytest.raises(ExcelToolError):
        plan_split(_df(a=["1"]), SplitSpec("by_value"), base_name="d")


def test_nombres_validos_en_windows():
    valores = ["Norte/Sur", "CON", "aux.txt", "x" * 300, "a/b", "a_b", "A_B", "", "nul"]
    df = _df(v=valores)
    partes = _partes(df, SplitSpec("by_value", value_columns=["v"], name_pattern="{valor}"))
    nombres = [p.name for p, _ in partes]
    assert len(set(n.lower() for n in nombres)) == len(nombres)
    for n in nombres:
        assert len(n) <= 150
        assert not set('<>:"/\\|?*') & set(n)
        assert n.split(".")[0].upper() not in {"CON", "AUX", "NUL", "PRN"}
        assert n == n.rstrip(" .")
    assert "Norte_Sur" in nombres
    assert "a_b" in nombres and "a_b (3)" in nombres
    assert "vacio" in nombres


def test_nombres_hoja_repetidos():
    df = _df(v=["a/b", "a_b", "x"])
    partes = _partes(df, SplitSpec("by_value", value_columns=["v"], name_pattern="{valor}"))
    assert {p.name for p, _ in partes} == {"a_b", "a_b (2)", "x"}


def test_cero_filas_no_da_error():
    df = _df(a=[], b=[])
    for spec in (
        SplitSpec("rows", rows_per_part=10),
        SplitSpec("sheets", rows_per_part=10),
        SplitSpec("rows_keep_groups", rows_per_part=10, group_column="a"),
        SplitSpec("by_value", value_columns=["a"]),
    ):
        assert plan_split(df, spec, base_name="d") == []
        assert list(split_dataframe(df, spec, base_name="d")) == []


def test_split_dataframe_es_perezoso():
    df = _df(a=[str(i) for i in range(100)])
    it = split_dataframe(df, SplitSpec("rows", rows_per_part=10), base_name="d")
    p, d = next(it)
    assert len(d) == 10 and d["a"].iloc[0] == "0"
    assert d.index.tolist() == list(range(10))


def test_patron_invalido():
    with pytest.raises(ExcelToolError):
        plan_split(
            _df(a=["1"]), SplitSpec("rows", rows_per_part=1, name_pattern="{nada}"), base_name="d"
        )


def test_valores_nulos_se_tratan_como_vacios():
    df = pd.DataFrame({"a": pd.array(["x", None, "x", None], dtype=STR)})
    partes = _partes(df, SplitSpec("by_value", value_columns=["a"]))
    assert [(p.value, p.rows) for p, _ in partes] == [(("x",), 2), (("",), 2)]


def test_grande_sintetico():
    rng = np.random.default_rng(1)
    n = 100_000
    df = _df(id=[str(i) for i in range(n)], g=[str(x) for x in rng.integers(0, 500, n)])
    p = _partes(df, SplitSpec("rows_keep_groups", rows_per_part=30_000, group_column="g"))
    assert sum(x.rows for x, _ in p) == n and all(x.rows <= 30_000 for x, _ in p)
    gs = [set(d["g"]) for _, d in p]
    assert sum(len(g) for g in gs) == len(set().union(*gs))
