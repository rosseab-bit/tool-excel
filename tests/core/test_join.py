"""Tests de core/join (F4-03): diagnóstico y unión de varios archivos."""

from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path

import pandas as pd
import pytest

from excel_tool import platform_utils as pu
from excel_tool.core import join as join_mod
from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy
from excel_tool.core.errors import Cancelado, ExcelToolError
from excel_tool.core.io import ReadOptions, WriteOptions, read_table, write_table
from excel_tool.core.join import (
    JoinPlan,
    SecondarySpec,
    SourceSpec,
    diagnose_join,
    final_column_names,
    run_join,
    suggest_reductions,
)
from excel_tool.core.normalize import KeyOptions
from excel_tool.core.progress import CancelToken

_STR = pd.StringDtype("pyarrow")
_RO = ReadOptions()
CEROS = KeyOptions(strip_leading_zeros=True)


def _tabla(**cols):
    return pd.DataFrame({k: pd.Series(v, dtype=_STR) for k, v in cols.items()})


def _dec(texto: str) -> Decimal:
    return Decimal(texto.replace(".", "").replace(",", ".")) if "," in texto else Decimal(texto)


def _src(nombre: str, df: pd.DataFrame, columnas: list[str] | None = None) -> SourceSpec:
    return SourceSpec(nombre, Path(nombre), _RO, list(df.columns) if columnas is None else columnas)


def _sec(nombre, df, pares, columnas=(), **kw) -> SecondarySpec:
    return SecondarySpec(nombre, Path(nombre), _RO, list(columnas), list(pares), **kw)


# --------------------------------------------------------------- escenario 3


@pytest.fixture(params=["csv", "xlsx"])
def escenario(request, fixtures_generados, tmp_path):
    base = fixtures_generados / "escenario_3"
    clientes = read_table(base / "clientes.xlsx").df
    ventas = read_table(base / "ventas.csv").df
    cobranzas = read_table(base / "cobranzas.xlsx").df
    if request.param == "xlsx":
        ventas_x = tmp_path / "ventas.xlsx"
        write_table(ventas, ventas_x, WriteOptions(format="xlsx"))
        ventas = read_table(ventas_x).df
    else:
        cob_c = tmp_path / "cobranzas.csv"
        write_table(cobranzas, cob_c, WriteOptions(format="csv_excel"))
        cobranzas = read_table(cob_c).df
    return {"clientes.xlsx": clientes, "ventas": ventas, "cobranzas": cobranzas}


@pytest.fixture(scope="module")
def esperado(fixtures_esperados):
    base = fixtures_esperados / "escenario_3"

    def leer(nombre):
        return json.loads((base / nombre).read_text(encoding="utf-8"))

    return {
        "diag": leer("diagnostico.json"),
        "ventas": leer("resumen_ventas.json"),
        "cobranzas": leer("resumen_cobranzas.json"),
        "region": leer("resumen_final_por_region.json"),
    }


def _principal(t):
    return _src("clientes.xlsx", t["clientes.xlsx"])


SUMA_VENTAS = AggregationSpec(
    [GroupBy("ID cliente")], [Aggregation("Importe", "sum", "Total comprado")]
)
SUMA_COBRANZAS = AggregationSpec(
    [GroupBy("Nro Cliente")], [Aggregation("Monto", "sum", "Total cobrado")]
)


def _ventas_detalle(t, **kw):
    return _sec(
        "ventas",
        t["ventas"],
        [("ID cliente", "ID cliente")],
        ["Fecha", "Artículo", "Importe"],
        **kw,
    )


def _ventas_resumen(t, resumen=SUMA_VENTAS, **kw):
    return _sec(
        "ventas",
        t["ventas"],
        [("ID cliente", "ID cliente")],
        mode="summary",
        summary=resumen,
        **kw,
    )


def _cobranzas(t, modo="detail", columnas=("Fecha", "Monto"), **kw):
    if modo == "summary":
        kw.setdefault("summary", SUMA_COBRANZAS)
    return _sec(
        "cobranzas", t["cobranzas"], [("ID cliente", "Nro Cliente")], columnas, mode=modo, **kw
    )


def test_detalle_de_ventas_y_resumen_de_cobranzas(escenario, esperado):
    plan = JoinPlan(
        _principal(escenario), CEROS, [_ventas_detalle(escenario), _cobranzas(escenario, "summary")]
    )
    res = run_join(plan, escenario)
    assert len(res.df) == esperado["diag"]["ventas"]["filas_resultado_detalle"] == 88
    assert res.diagnosis.rows_result == 88
    esp = esperado["cobranzas"]["por_cliente"]
    for _, fila in res.df.iterrows():
        e = esp.get(fila["ID cliente"])
        if e is None:
            assert fila["Total cobrado"] == "0"  # ADR-037
        else:
            assert _dec(fila["Total cobrado"]) == Decimal(e["suma_monto"])
    assert res.column_types["Total cobrado"] == "decimal"
    # la clave sale una sola vez, con el valor del principal
    assert list(res.df.columns).count("ID cliente") == 1
    assert "Nro Cliente" not in res.df.columns


def test_resumen_de_ventas_por_cliente(escenario, esperado):
    plan = JoinPlan(_principal(escenario), KeyOptions(), [_ventas_resumen(escenario)])
    res = run_join(plan, escenario)
    assert len(res.df) == len(escenario["clientes.xlsx"]) == 30
    assert list(res.df["ID cliente"]) == list(escenario["clientes.xlsx"]["ID cliente"])
    esp = esperado["ventas"]["por_cliente"]
    for _, fila in res.df.iterrows():
        e = esp.get(fila["ID cliente"])
        if e is None:
            assert fila["Total comprado"] == "0"  # ADR-037
        else:
            assert _dec(fila["Total comprado"]) == Decimal(e["suma_importe"])
    # los 3 importes que no se pueden leer se informan (el vacío se informa aparte)
    assert any("3 celda" in w and "Importe" in w for w in res.warnings)


def test_resumen_por_cliente_y_articulo(escenario):
    resumen = AggregationSpec(
        [GroupBy("ID cliente"), GroupBy("Artículo")],
        [Aggregation("Importe", "sum", "Total comprado")],
    )
    plan = JoinPlan(_principal(escenario), KeyOptions(), [_ventas_resumen(escenario, resumen)])
    res = run_join(plan, escenario)
    ventas = escenario["ventas"]
    ids = set(escenario["clientes.xlsx"]["ID cliente"])
    combos = {(c, a) for c, a in zip(ventas["ID cliente"], ventas["Artículo"], strict=True)}
    combos = {x for x in combos if x[0] in ids}
    sin_ventas = len(ids) - len({c for c, _ in combos})
    assert len(res.df) == len(combos) + sin_ventas
    assert res.diagnosis.rows_result == len(res.df)
    con_datos = res.df[res.df["Artículo"] != ""]
    assert {
        (c, a) for c, a in zip(con_datos["ID cliente"], con_datos["Artículo"], strict=True)
    } == combos
    # la suma de todo lo cobrado es la suma por cliente
    total = sum(_dec(x) for x in res.df["Total comprado"] if x)
    por_cliente = run_join(
        JoinPlan(_principal(escenario), KeyOptions(), [_ventas_resumen(escenario)]), escenario
    )
    assert total == sum(_dec(x) for x in por_cliente.df["Total comprado"] if x)


def test_resumen_final_por_region(escenario, esperado):
    final = AggregationSpec(
        [GroupBy("Región")], [Aggregation("Total comprado", "sum", "Total vendido")]
    )
    plan = JoinPlan(
        _principal(escenario), KeyOptions(), [_ventas_resumen(escenario)], final_summary=final
    )
    diag = diagnose_join(plan, escenario)
    assert diag.rows_result == 5
    assert diag.columns_result == 2
    res = run_join(plan, escenario)
    assert len(res.df) == diag.rows_result == 5
    esp = esperado["region"]["total_importe_por_region"]
    obtenido = dict(zip(res.df["Región"], res.df["Total vendido"], strict=True))
    assert {k: _dec(v) for k, v in obtenido.items()} == {k: Decimal(v) for k, v in esp.items()}
    assert res.column_types["Total vendido"] == "decimal"


def test_dos_detalles_avisa_y_estima(escenario, esperado):
    plan = JoinPlan(
        _principal(escenario), CEROS, [_ventas_detalle(escenario), _cobranzas(escenario)]
    )
    diag = diagnose_join(plan, escenario)
    esp = esperado["diag"]["cobranzas"]
    assert diag.many_details
    assert any("todas con todas" in w for w in diag.warnings)
    assert diag.secondaries[0].rows_after == esperado["diag"]["ventas"]["filas_resultado_detalle"]
    assert (
        diag.rows_result
        == esp["sin_ceros_a_la_izquierda"]["filas_resultado_ventas_detalle_y_cobranzas_detalle"]
    )
    res = run_join(plan, escenario)
    assert len(res.df) == diag.rows_result == 210
    # sin la opción de ceros
    plan2 = JoinPlan(
        _principal(escenario), KeyOptions(), [_ventas_detalle(escenario), _cobranzas(escenario)]
    )
    d2 = diagnose_join(plan2, escenario)
    assert (
        d2.rows_result
        == esp["solo_recortar_espacios"]["filas_resultado_ventas_detalle_y_cobranzas_detalle"]
    )
    assert len(run_join(plan2, escenario).df) == d2.rows_result == 91


def test_un_solo_detalle_no_avisa(escenario):
    plan = JoinPlan(
        _principal(escenario), CEROS, [_ventas_detalle(escenario), _cobranzas(escenario, "summary")]
    )
    diag = diagnose_join(plan, escenario)
    assert not diag.many_details
    assert not any("todas con todas" in w for w in diag.warnings)


def test_ceros_a_la_izquierda(escenario, esperado):
    esp = esperado["diag"]["cobranzas"]
    sin = JoinPlan(_principal(escenario), KeyOptions(), [_cobranzas(escenario)])
    d = diagnose_join(sin, escenario)
    s = d.secondaries[0]
    assert s.keys_matched == esp["solo_recortar_espacios"]["claves_con_coincidencia"] == 5
    assert s.keys_unmatched_sample[:3] == ["1", "2", "3"]
    assert s.secondary_orphans == len(
        esp["solo_recortar_espacios"]["claves_del_secundario_sin_pareja"]
    )
    assert s.keys_matched_without_zeros == 27
    assert any(
        "si sacás los ceros a la izquierda, coinciden 27 claves en vez de 5" in w
        for w in d.warnings
    )
    assert s.rows_after == esp["solo_recortar_espacios"]["filas_resultado_detalle"]

    con = JoinPlan(_principal(escenario), CEROS, [_cobranzas(escenario)])
    d2 = diagnose_join(con, escenario)
    s2 = d2.secondaries[0]
    assert s2.keys_matched == esp["sin_ceros_a_la_izquierda"]["claves_con_coincidencia"] == 27
    assert s2.keys_matched_without_zeros is None
    assert not any("ceros" in w for w in d2.warnings)
    assert (
        s2.keys_with_many_rows
        == esp["sin_ceros_a_la_izquierda"]["claves_con_varias_filas_en_secundario"]
    )
    assert s2.rows_after == esp["sin_ceros_a_la_izquierda"]["filas_resultado_detalle"]
    assert s2.secondary_orphans_sample == ["00077"]
    assert s2.example_many is not None and s2.example_many[1] >= 2


def test_diagnostico_de_ventas(escenario, esperado):
    esp = esperado["diag"]["ventas"]
    plan = JoinPlan(_principal(escenario), KeyOptions(), [_ventas_detalle(escenario)])
    d = diagnose_join(plan, escenario)
    s = d.secondaries[0]
    assert s.keys_principal == esp["claves_principal"]
    assert s.keys_matched == esp["claves_con_coincidencia"]
    assert s.keys_unmatched_sample == esp["claves_sin_coincidencia"]
    assert s.secondary_orphans == 2
    assert sorted(s.secondary_orphans_sample) == sorted(esp["claves_del_secundario_sin_pareja"])
    assert s.keys_with_many_rows == esp["claves_con_varias_filas_en_secundario"]
    assert s.rows_after == d.rows_result == esp["filas_resultado_detalle"]
    assert s.keys_matched_without_zeros is None  # sin ceros no hay nada que sugerir
    assert d.columns_result == 4 + 3
    assert d.size.rows == 88 and d.size.columns == 7 and d.size.approx_mb_csv > 0
    assert d.memory.ok


def test_sin_pareja_para_hojas_extra(escenario):
    plan = JoinPlan(_principal(escenario), KeyOptions(), [_ventas_detalle(escenario)])
    res = run_join(plan, escenario)
    sin = res.unmatched["ventas"]
    assert sorted(sin["ID cliente"]) == ["98", "99"]
    assert "Importe" in sin.columns


def test_inner_contra_left(escenario):
    izq = JoinPlan(_principal(escenario), KeyOptions(), [_ventas_detalle(escenario)])
    inn = JoinPlan(_principal(escenario), KeyOptions(), [_ventas_detalle(escenario, how="inner")])
    r_left, r_inner = run_join(izq, escenario), run_join(inn, escenario)
    assert len(r_left.df) == 88
    assert len(r_inner.df) == 85  # 88 menos los 3 clientes sin ventas
    assert "5" in set(r_left.df["ID cliente"]) and "5" not in set(r_inner.df["ID cliente"])
    assert diagnose_join(inn, escenario).rows_result == 85
    # inner con un resumen: solo clientes con ventas
    inn2 = JoinPlan(_principal(escenario), KeyOptions(), [_ventas_resumen(escenario, how="inner")])
    assert len(run_join(inn2, escenario).df) == 27


def test_primera_fila(escenario):
    plan = JoinPlan(
        _principal(escenario),
        KeyOptions(),
        [
            _sec(
                "ventas",
                escenario["ventas"],
                [("ID cliente", "ID cliente")],
                ["Artículo"],
                mode="first",
            )
        ],
    )
    res = run_join(plan, escenario)
    assert len(res.df) == 30
    ventas = escenario["ventas"]
    primera = ventas.drop_duplicates("ID cliente").set_index("ID cliente")["Artículo"]
    for _, fila in res.df.iterrows():
        assert fila["Artículo"] == primera.get(fila["ID cliente"], "")


def test_columnas_con_el_mismo_nombre(escenario):
    plan = JoinPlan(
        _principal(escenario),
        CEROS,
        [
            _ventas_detalle(escenario),
            _cobranzas(escenario, columnas=("Fecha", "Monto"), prefix=None),
        ],
    )
    res = run_join(plan, escenario)
    assert "Fecha" in res.df.columns
    assert "Fecha (cobranzas)" in res.df.columns
    assert len(set(res.df.columns)) == len(res.df.columns)
    con_prefijo = JoinPlan(
        _principal(escenario),
        CEROS,
        [_ventas_detalle(escenario), _cobranzas(escenario, prefix="Cobro ")],
    )
    res2 = run_join(con_prefijo, escenario)
    assert {"Cobro Fecha", "Cobro Monto", "Fecha"} <= set(res2.df.columns)


def test_orden_de_columnas(escenario):
    plan = JoinPlan(
        _principal(escenario),
        KeyOptions(),
        [_ventas_resumen(escenario)],
        output_column_order=["Total comprado", "Nombre", "ID cliente", "No existe"],
    )
    res = run_join(plan, escenario)
    assert list(res.df.columns)[:3] == ["Total comprado", "Nombre", "ID cliente"]
    assert set(res.df.columns) == {"ID cliente", "Nombre", "Región", "Vendedor", "Total comprado"}
    assert plan.output_column_order is not None
    assert diagnose_join(plan, escenario).columns_result == 5


# --------------------------------------------------------------- casos chicos


def test_claves_multiples():
    p = _tabla(suc=["A", "A", "B", "B"], cli=["1", "2", "1", "2"], nombre=["a1", "a2", "b1", "b2"])
    s = _tabla(Cli=["1", "2", "1", "9"], Suc=["a", "A", "B", "B"], dato=["x", "y", "z", "w"])
    plan = JoinPlan(
        _src("p", p),
        KeyOptions(),
        [_sec("s", s, [("suc", "Suc"), ("cli", "Cli")], ["dato"])],
    )
    res = run_join(plan, {"p": p, "s": s})
    assert list(res.df["nombre"]) == ["a1", "a2", "b1", "b2"]
    # (A,1) coincide con (a,1) porque se ignoran las mayúsculas; (B,2) y (9) no tienen pareja
    assert list(res.df["dato"]) == ["x", "y", "z", ""]


def test_claves_multiples_ignora_mayusculas():
    p = _tabla(suc=["A", "B"], cli=["1", "1"])
    s = _tabla(Suc=["a", "b"], Cli=["1", "2"], dato=["x", "y"])
    plan = JoinPlan(
        _src("p", p), KeyOptions(), [_sec("s", s, [("suc", "Suc"), ("cli", "Cli")], ["dato"])]
    )
    d = diagnose_join(plan, {"p": p, "s": s})
    assert d.secondaries[0].keys_matched == 1
    res = run_join(plan, {"p": p, "s": s})
    assert list(res.df["dato"]) == ["x", ""]


def test_claves_vacias_no_emparejan():
    p = _tabla(k=["1", "", " ", "2"], v=["a", "b", "c", "d"])
    s = _tabla(k=["", "1", "  ", "2"], w=["x", "y", "z", "t"])
    for modo in ("detail", "first"):
        plan = JoinPlan(_src("p", p), KeyOptions(), [_sec("s", s, [("k", "k")], ["w"], mode=modo)])
        res = run_join(plan, {"p": p, "s": s})
        assert list(res.df["w"]) == ["y", "", "", "t"], modo
        assert list(res.unmatched["s"]["w"]) == ["x", "z"]
    resumen = AggregationSpec([GroupBy("k")], [Aggregation("", "count", "Cant")])
    plan = JoinPlan(
        _src("p", p), KeyOptions(), [_sec("s", s, [("k", "k")], mode="summary", summary=resumen)]
    )
    res = run_join(plan, {"p": p, "s": s})
    assert list(res.df["Cant"]) == ["1", "0", "0", "1"]  # ADR-037
    d = diagnose_join(plan, {"p": p, "s": s})
    assert d.secondaries[0].keys_principal == 2  # las claves vacías no cuentan
    assert d.secondaries[0].secondary_orphans == 0


def test_resumen_agrupa_la_clave_normalizada():
    p = _tabla(k=["ab", "cd"])
    s = _tabla(k=[" AB ", "ab", "Ab", "cd"], imp=["1,50", "2,50", "1", "7"])
    resumen = AggregationSpec([GroupBy("k")], [Aggregation("imp", "sum", "Total")])
    plan = JoinPlan(
        _src("p", p), KeyOptions(), [_sec("s", s, [("k", "k")], mode="summary", summary=resumen)]
    )
    res = run_join(plan, {"p": p, "s": s})
    assert len(res.df) == 2  # " AB ", "ab" y "Ab" son la misma clave: no duplica filas
    assert [_dec(x) for x in res.df["Total"]] == [Decimal("5"), Decimal("7")]


def test_claves_repetidas_en_el_principal():
    p = _tabla(k=["1", "1", "2", "3"], v=["a", "b", "c", "d"])
    s = _tabla(k=["1", "1", "1", "2"], w=["x", "y", "z", "t"])
    plan = JoinPlan(_src("p", p), KeyOptions(), [_sec("s", s, [("k", "k")], ["w"])])
    d = diagnose_join(plan, {"p": p, "s": s})
    assert d.rows_result == 2 * 3 + 1 + 1  # dos filas de "1" x 3 + "2" + "3"
    aviso = [w for w in d.warnings if "repetida" in w]
    assert len(aviso) == 1 and "«1»" in aviso[0] and "2 filas" in aviso[0]
    assert "se multiplican" in aviso[0]
    res = run_join(plan, {"p": p, "s": s})
    assert len(res.df) == 8
    # en resumen no se multiplican las filas del principal
    resumen = AggregationSpec([GroupBy("k")], [Aggregation("", "count", "Cant")])
    plan2 = JoinPlan(
        _src("p", p), KeyOptions(), [_sec("s", s, [("k", "k")], mode="summary", summary=resumen)]
    )
    d2 = diagnose_join(plan2, {"p": p, "s": s})
    assert d2.rows_result == 4
    assert len(run_join(plan2, {"p": p, "s": s}).df) == 4


def test_control_de_filas_en_codigo():
    assert join_mod._control_de_filas(
        JoinPlan(
            _src("p", _tabla(k=["1"])),
            KeyOptions(),
            [_sec("s", _tabla(k=["1"]), [("k", "k")], mode="first")],
        )
    )
    detalle = _sec("s", _tabla(k=["1"]), [("k", "k")], mode="detail")
    assert not join_mod._control_de_filas(
        JoinPlan(_src("p", _tabla(k=["1"])), KeyOptions(), [detalle])
    )
    final = AggregationSpec([GroupBy("k")], [Aggregation("", "count")])
    primero = _sec("s", _tabla(k=["1"]), [("k", "k")], mode="first")
    assert not join_mod._control_de_filas(
        JoinPlan(_src("p", _tabla(k=["1"])), KeyOptions(), [primero], final_summary=final)
    )


def test_sin_detalle_ni_resumen_final_mismas_filas_que_el_principal(escenario):
    resumen_total = AggregationSpec(
        [GroupBy("ID cliente")],
        [Aggregation("Importe", "sum", "Total comprado"), Aggregation("", "count", "Compras")],
    )
    plan = JoinPlan(
        _principal(escenario),
        CEROS,
        [_ventas_resumen(escenario, resumen_total), _cobranzas(escenario, "summary")],
    )
    res = run_join(plan, escenario)
    assert len(res.df) == len(escenario["clientes.xlsx"]) == 30


def test_el_control_de_filas_corta_si_algo_falla(monkeypatch):
    p = _tabla(k=["1", "2"])
    s = _tabla(k=["1", "2"], w=["x", "y"])
    plan = JoinPlan(_src("p", p), KeyOptions(), [_sec("s", s, [("k", "k")], ["w"], mode="first")])
    original = join_mod._unir

    def roto(pos, c, t, cancel):
        r = original(pos, c, t, cancel)
        return join_mod._Posiciones(r.pidx, r.sidx, r.n + 1)

    monkeypatch.setattr(join_mod, "_unir", roto)
    with pytest.raises(ExcelToolError):
        run_join(plan, {"p": p, "s": s})


def test_errores_amigables():
    p = _tabla(k=["1"])
    s = _tabla(k=["1"], w=["x"])
    con = {"p": p, "s": s}
    with pytest.raises(ExcelToolError, match="No encuentro la columna «zzz»"):
        diagnose_join(JoinPlan(_src("p", p), KeyOptions(), [_sec("s", s, [("k", "zzz")])]), con)
    with pytest.raises(ExcelToolError, match="Falta cargar"):
        diagnose_join(JoinPlan(_src("p", p), KeyOptions(), [_sec("x", s, [("k", "k")])]), con)
    with pytest.raises(ExcelToolError, match="Falta definir el resumen"):
        diagnose_join(
            JoinPlan(_src("p", p), KeyOptions(), [_sec("s", s, [("k", "k")], mode="summary")]), con
        )
    with pytest.raises(ExcelToolError, match="por qué columna"):
        diagnose_join(JoinPlan(_src("p", p), KeyOptions(), [_sec("s", s, [])]), con)


def test_memoria_insuficiente_no_arranca(monkeypatch):
    p = _tabla(k=["1", "2"])
    s = _tabla(k=["1", "2"], w=["x", "y"])
    plan = JoinPlan(_src("p", p), KeyOptions(), [_sec("s", s, [("k", "k")], ["w"])])
    monkeypatch.setattr(pu, "available_memory", lambda: 10)
    d = diagnose_join(plan, {"p": p, "s": s})
    assert not d.memory.ok
    with pytest.raises(ExcelToolError, match="memoria"):
        run_join(plan, {"p": p, "s": s})


def test_progreso_y_cancelar(escenario):
    plan = JoinPlan(
        _principal(escenario), CEROS, [_ventas_detalle(escenario), _cobranzas(escenario, "summary")]
    )
    eventos = []
    run_join(plan, escenario, progress=eventos.append)
    assert eventos and eventos[-1].done == eventos[-1].total
    hechos = [e.done for e in eventos]
    assert hechos == sorted(hechos)
    assert all(e.message for e in eventos)
    cancel = CancelToken()
    cancel.cancel()
    with pytest.raises(Cancelado):
        run_join(plan, escenario, cancel=cancel)
    with pytest.raises(Cancelado):
        diagnose_join(plan, escenario, cancel=cancel)


def test_sin_secundarios_devuelve_el_principal(escenario):
    res = run_join(JoinPlan(_principal(escenario), KeyOptions(), []), escenario)
    assert len(res.df) == 30
    assert list(res.df.columns) == list(escenario["clientes.xlsx"].columns)


def test_las_columnas_del_principal_se_eligen(escenario):
    principal = _src("clientes.xlsx", escenario["clientes.xlsx"], ["Nombre"])
    plan = JoinPlan(principal, KeyOptions(), [_ventas_resumen(escenario)])
    res = run_join(plan, escenario)
    # la clave sale siempre aunque no se haya marcado
    assert list(res.df.columns) == ["ID cliente", "Nombre", "Total comprado"]


# --------------------------------------------------------------- escenario grande


_GRANDES = Path(__file__).resolve().parents[1] / "fixtures" / "generados" / "grandes"


@pytest.mark.skipif(
    not (_GRANDES / "ventas_grande.csv").is_file(), reason="falta el escenario grande"
)
@pytest.mark.parametrize("modo", ["detail", "summary"])
def test_escenario_grande_filas_estimadas_igual_a_reales(modo, fixtures_esperados):
    esperado = json.loads((fixtures_esperados / "grandes.json").read_text(encoding="utf-8"))
    clientes = read_table(_GRANDES / "clientes_grande.csv").df
    ventas = read_table(_GRANDES / "ventas_grande.csv").df
    if modo == "detail":
        sec = _sec(
            "ventas", ventas, [("ID cliente", "ID cliente")], ["Fecha", "Importe"], mode="detail"
        )
    else:
        sec = _sec(
            "ventas",
            ventas,
            [("ID cliente", "ID cliente")],
            mode="summary",
            summary=AggregationSpec([GroupBy("ID cliente")], [Aggregation("", "count", "Compras")]),
        )
    plan = JoinPlan(_src("clientes", clientes), KeyOptions(), [sec])
    tablas = {"clientes": clientes, "ventas": ventas}
    diag = diagnose_join(plan, tablas)
    res = run_join(plan, tablas)
    assert len(res.df) == diag.rows_result
    if modo == "detail":
        assert diag.rows_result == esperado["detalle_left"]["total"]
    else:
        assert diag.rows_result == len(clientes)


# --------------------------------------------------------------- ADR-037: ceros sin pareja

SIN_VENTAS = ["5", "17", "28"]
VARIOS = AggregationSpec(
    [GroupBy("ID cliente")],
    [
        Aggregation("Importe", "sum", "Total comprado"),
        Aggregation("", "count", "Cantidad de filas"),
        Aggregation("Artículo", "count_distinct", "Artículos"),
        Aggregation("Importe", "mean", "Promedio"),
        Aggregation("Fecha", "max", "Última compra"),
        Aggregation("Artículo", "first", "Primer artículo"),
    ],
)


def _sin_ventas(res, col):
    df = res.df
    return df.loc[df["ID cliente"].isin(SIN_VENTAS), col]


def _controlar_ceros(res, nombres):
    total, filas, arts, prom, ultima, primero = nombres
    for col in (total, filas, arts):
        sub = _sin_ventas(res, col)
        assert len(sub) == 3 and (sub == "0").all(), col
    for col in (prom, ultima, primero):
        assert (_sin_ventas(res, col) == "").all(), col
    con = res.df[~res.df["ID cliente"].isin(SIN_VENTAS)]
    assert (con[filas] != "0").all()
    assert res.column_types[total] == "decimal"
    assert res.column_types[filas] == "entero"


NOMBRES = [
    "Total comprado",
    "Cantidad de filas",
    "Artículos",
    "Promedio",
    "Última compra",
    "Primer artículo",
]


def test_sin_pareja_ceros_en_sumas_y_cuentas(escenario):
    plan = JoinPlan(_principal(escenario), KeyOptions(), [_ventas_resumen(escenario, VARIOS)])
    res = run_join(plan, escenario)
    assert len(res.df) == 30
    _controlar_ceros(res, NOMBRES)


def test_sin_pareja_ceros_con_prefijo(escenario):
    plan = JoinPlan(
        _principal(escenario), KeyOptions(), [_ventas_resumen(escenario, VARIOS, prefix="V - ")]
    )
    res = run_join(plan, escenario)
    _controlar_ceros(res, ["V - " + n for n in NOMBRES])


def test_sin_pareja_ceros_con_nombres_repetidos(escenario):
    # el principal ya tiene "Total comprado": el del resumen se renombra con el archivo
    t = dict(escenario)
    clientes = t["clientes.xlsx"].copy()
    clientes["Total comprado"] = "x"
    t["clientes.xlsx"] = clientes
    plan = JoinPlan(
        _src("clientes.xlsx", clientes, ["ID cliente", "Total comprado"]),
        KeyOptions(),
        [_ventas_resumen(t, VARIOS)],
    )
    res = run_join(plan, t)
    assert (res.df["Total comprado"] == "x").all()
    _controlar_ceros(res, ["Total comprado (ventas)", *NOMBRES[1:]])


def test_resumen_final_con_ceros_suma_igual(escenario, esperado):
    final = AggregationSpec(
        [GroupBy("Región")],
        [Aggregation("Total comprado", "sum", "Total"), Aggregation("", "count", "Clientes")],
    )
    plan = JoinPlan(
        _principal(escenario),
        KeyOptions(),
        [_ventas_resumen(escenario, VARIOS)],
        final_summary=final,
    )
    res = run_join(plan, escenario)
    assert sum(int(x) for x in res.df["Clientes"]) == 30  # cuenta filas, no cambia


# --------------------------------------------------------------- F5-01: achicar el resultado


def _plan_dos_detalles(t):
    return JoinPlan(
        _principal(t),
        CEROS,
        [_ventas_detalle(t), _cobranzas(t)],
    )


def test_reducciones_ventas_en_detalle_da_una_fila_por_cliente(escenario):
    ventas = _sec(
        "ventas",
        escenario["ventas"],
        [("ID cliente", "ID cliente")],
        ["Fecha", "Cantidad", "Importe"],
    )
    plan = JoinPlan(_principal(escenario), KeyOptions(), [ventas])
    base = diagnose_join(plan, escenario)
    sugeridas = suggest_reductions(plan, escenario)
    assert len(sugeridas) == 1
    r = sugeridas[0]
    assert r.description == "Cambiar ventas a resumen por ID cliente (una fila por cliente)"
    assert r.rows == 30 < base.rows_result
    assert r.size.rows == 30 and r.columns == r.size.columns
    sec = r.plan.secondaries[0]
    assert sec.mode == "summary" and sec.summary is not None
    nombres = [a.func for a in sec.summary.aggregations]
    assert nombres[0] == "count"
    assert [a.column for a in sec.summary.aggregations if a.func == "sum"] == [
        "Cantidad",
        "Importe",
    ]
    # el plan alternativo, ejecutado, da exactamente esas filas y columnas
    res = run_join(r.plan, escenario)
    assert len(res.df) == r.rows == res.diagnosis.rows_result
    assert len(res.df.columns) == r.columns
    # el plan original no se tocó
    assert plan.secondaries[0].mode == "detail"


def test_reducciones_dos_detalles_todos_menos_uno(escenario):
    plan = _plan_dos_detalles(escenario)
    base = diagnose_join(plan, escenario)
    assert base.many_details
    sugeridas = suggest_reductions(plan, escenario)
    todas = [r for r in sugeridas if r.description.startswith("Dejar el detalle de")]
    assert len(todas) == 1
    r = todas[0]
    modos = {s.name: s.mode for s in r.plan.secondaries}
    assert sorted(modos.values()) == ["detail", "summary"]
    # se deja el detalle de la tabla con más filas (ventas: 87 > cobranzas: 67)
    assert modos["ventas"] == "detail"
    assert "cobranzas" in r.description
    assert not run_join(r.plan, escenario).diagnosis.many_details
    assert len(run_join(r.plan, escenario).df) == r.rows < base.rows_result
    # ordenadas de la que más achica a la que menos, todas más chicas que la actual
    pesos = [x.size.approx_mb_csv for x in sugeridas]
    assert pesos == sorted(pesos)
    assert all(x.rows <= base.rows_result for x in sugeridas)
    # no se repite el mismo plan
    assert len({repr(x.plan) for x in sugeridas}) == len(sugeridas)


def test_reducciones_sin_detalles_no_sugiere_resumenes(escenario):
    plan = JoinPlan(_principal(escenario), KeyOptions(), [_ventas_resumen(escenario)])
    assert suggest_reductions(plan, escenario) == []
    assert suggest_reductions(JoinPlan(_principal(escenario), KeyOptions(), []), escenario) == []


@pytest.fixture(params=["csv", "xlsx"])
def con_texto_largo(request, tmp_path):
    n = 300
    clientes = _tabla(
        **{
            "ID": [str(i) for i in range(n)],
            "Nombre": [f"Cliente {i}" for i in range(n)],
            "Observaciones": [f"Observación larga del cliente número {i} " * 3 for i in range(n)],
            "Descripción": [f"Descripción comercial extensa del rubro {i} " * 2 for i in range(n)],
        }
    )
    if request.param == "xlsx":
        write_table(clientes, tmp_path / "c.xlsx", WriteOptions(format="xlsx"))
        clientes = read_table(tmp_path / "c.xlsx").df
    else:
        write_table(clientes, tmp_path / "c.csv", WriteOptions(format="csv_excel"))
        clientes = read_table(tmp_path / "c.csv").df
    ventas = _tabla(**{"ID": [str(i % n) for i in range(n)], "Importe": ["1"] * n})
    return {"c": clientes, "v": ventas}


def test_reducciones_quitar_columnas_pesadas(con_texto_largo):
    t = con_texto_largo
    plan = JoinPlan(
        _src("c", t["c"]),
        KeyOptions(),
        [_sec("v", t["v"], [("ID", "ID")], ["Importe"], mode="first")],
    )
    sugeridas = suggest_reductions(plan, t)
    assert len(sugeridas) == 1
    r = sugeridas[0]
    assert (
        r.description == ("Quitar las columnas Observaciones y Descripción (las que más ocupan)")
        or r.description == "Quitar las columnas Descripción y Observaciones (las que más ocupan)"
    )
    assert r.plan.principal.columns == ["ID", "Nombre"]
    assert r.rows == 300 and r.columns == 3
    assert r.size.approx_mb_csv < 0.85 * diagnose_join(plan, t).size.approx_mb_csv
    res = run_join(r.plan, t)
    assert list(res.df.columns) == ["ID", "Nombre", "Importe"] and len(res.df) == 300
    assert r.columns == len(res.df.columns)


def test_reducciones_no_quita_claves_ni_columnas_del_resumen_final(con_texto_largo):
    t = con_texto_largo
    final = AggregationSpec([GroupBy("Observaciones")], [Aggregation("", "count")])
    plan = JoinPlan(_src("c", t["c"]), KeyOptions(), [], final_summary=final)
    assert suggest_reductions(plan, t) == []
    largas = _tabla(ID=["a" * 60 + str(i) for i in range(50)], Otro=["x"] * 50)
    plan = JoinPlan(
        _src("l", largas),
        KeyOptions(),
        [_sec("v", largas, [("ID", "ID")], ["Otro"], mode="first")],
    )
    assert suggest_reductions(plan, {"l": largas, "v": largas}) == []


def test_reducciones_progreso_y_cancelacion(escenario):
    plan = _plan_dos_detalles(escenario)
    eventos = []
    suggest_reductions(plan, escenario, progress=eventos.append)
    assert eventos and eventos[-1].done == eventos[-1].total
    cancel = CancelToken()
    cancel.cancel()
    with pytest.raises(Cancelado):
        suggest_reductions(plan, escenario, cancel=cancel)


@pytest.mark.skipif(
    not (_GRANDES / "ventas_grande.csv").is_file(), reason="falta el escenario grande"
)
def test_reducciones_escenario_grande_es_rapido():
    import time

    clientes = read_table(_GRANDES / "clientes_grande.csv").df
    ventas = read_table(_GRANDES / "ventas_grande.csv").df
    sec = _sec("ventas", ventas, [("ID cliente", "ID cliente")], ["Fecha", "Cantidad", "Importe"])
    plan = JoinPlan(_src("clientes", clientes), KeyOptions(), [sec])
    t0 = time.perf_counter()
    sugeridas = suggest_reductions(plan, {"clientes": clientes, "ventas": ventas})
    assert time.perf_counter() - t0 < 5
    assert len(sugeridas) >= 1
    assert sugeridas[0].rows == 200_000
    assert sugeridas[0].plan.secondaries[0].mode == "summary"


# --------------------------------------------------------------- F5-03: nombres finales


def _repetidos(t):
    t = dict(t)
    clientes = t["clientes.xlsx"].copy()
    clientes["Total comprado"] = "x"
    clientes["Fecha"] = "y"
    t["clientes.xlsx"] = clientes
    return t


def _igual_que_run_join(plan, tablas):
    nombres = final_column_names(plan, tablas)
    assert [f for _, _, f in nombres] == list(run_join(plan, tablas).df.columns)
    return nombres


def test_nombres_finales_detalle_y_clave_sin_archivo(escenario):
    plan = JoinPlan(_principal(escenario), CEROS, [_ventas_detalle(escenario)])
    nombres = _igual_que_run_join(plan, escenario)
    assert "ID cliente" in {o for a, o, _ in nombres if a is None}
    assert ("ventas", "Importe", "Importe") in nombres


def test_nombres_finales_resumen_y_primera_fila(escenario):
    plan = JoinPlan(
        _principal(escenario), CEROS, [_ventas_resumen(escenario), _cobranzas(escenario, "first")]
    )
    nombres = _igual_que_run_join(plan, escenario)
    assert ("ventas", "Total comprado", "Total comprado") in nombres


def test_nombres_finales_prefix_y_repetidos(escenario):
    t = _repetidos(escenario)
    principal = _src("clientes.xlsx", t["clientes.xlsx"])
    plan = JoinPlan(
        principal,
        CEROS,
        [_ventas_detalle(t), _cobranzas(t, "summary", prefix="Cob. ")],
    )
    nombres = _igual_que_run_join(plan, t)
    assert ("cobranzas", "Total cobrado", "Cob. Total cobrado") in nombres
    assert ("ventas", "Fecha", "Fecha (ventas)") in nombres


def test_nombres_finales_con_orden(escenario):
    plan = JoinPlan(
        _principal(escenario),
        CEROS,
        [_ventas_detalle(escenario)],
        output_column_order=["Importe", "inexistente", "ID cliente", "Importe"],
    )
    nombres = _igual_que_run_join(plan, escenario)
    assert nombres[0][2] == "Importe"
    assert nombres[1][2] == "ID cliente"


def test_nombres_finales_con_resumen_final(escenario):
    final = AggregationSpec(
        [GroupBy("Región")], [Aggregation("Total comprado", "sum", "Total vendido")]
    )
    plan = JoinPlan(
        _principal(escenario), KeyOptions(), [_ventas_resumen(escenario)], final_summary=final
    )
    nombres = _igual_que_run_join(plan, escenario)
    assert nombres == [(None, "Región", "Región"), (None, "Total vendido", "Total vendido")]


def test_nombres_finales_grande_no_lee_filas():
    import time

    n = 1_000_000
    p = pd.DataFrame(
        {"id": pd.Series(["1"] * n, dtype=_STR), "a": pd.Series(["x"] * n, dtype=_STR)}
    )
    s = pd.DataFrame(
        {"id": pd.Series(["1"] * n, dtype=_STR), "a": pd.Series(["y"] * n, dtype=_STR)}
    )
    plan = JoinPlan(_src("p.csv", p), KeyOptions(), [_sec("s.xlsx", s, [("id", "id")], ["a"])])
    t0 = time.perf_counter()
    nombres = final_column_names(plan, {"p.csv": p, "s.xlsx": s})
    assert time.perf_counter() - t0 < 0.5
    assert [f for _, _, f in nombres] == ["id", "a", "a (s)"]


# ---------------------------------------------------------------- F5-05: resumen final sin detalle


def _resumenes_finales():
    return [
        AggregationSpec(
            [GroupBy("Región")],
            [Aggregation("Importe", "sum", "Total"), Aggregation("", "count")],
        ),
        # columnas de tres archivos (clientes, ventas y cobranzas) y un período
        AggregationSpec(
            [GroupBy("Región"), GroupBy("Fecha", "month")],
            [
                Aggregation("Importe", "sum"),
                Aggregation("Total cobrado", "sum"),
                Aggregation("Artículo", "count_distinct"),
            ],
        ),
        AggregationSpec([GroupBy("Artículo")], [Aggregation("Importe", "max")]),
    ]


@pytest.mark.parametrize("indice", [0, 1, 2])
def test_resumen_final_con_solo_las_columnas_usadas_da_lo_mismo(escenario, monkeypatch, indice):
    final = _resumenes_finales()[indice]
    plan = JoinPlan(
        _principal(escenario),
        CEROS,
        [
            _ventas_detalle(escenario, how="inner" if indice == 2 else "left"),
            _cobranzas(escenario, "summary"),
        ],
        final_summary=final,
    )
    nuevo = run_join(plan, escenario)
    # la versión de antes: armar todas las columnas
    monkeypatch.setattr(join_mod, "_columnas_del_resumen", lambda spec: None)
    viejo = run_join(plan, escenario)
    pd.testing.assert_frame_equal(nuevo.df, viejo.df)
    assert nuevo.column_types == viejo.column_types
    assert nuevo.warnings == viejo.warnings
    assert nuevo.diagnosis.rows_result == viejo.diagnosis.rows_result == len(nuevo.df)
    assert nuevo.unmatched.keys() == viejo.unmatched.keys()
    for nombre, df in nuevo.unmatched.items():
        pd.testing.assert_frame_equal(df, viejo.unmatched[nombre])


def test_resumen_final_no_arma_las_columnas_que_no_usa(escenario, monkeypatch):
    plan = JoinPlan(
        _principal(escenario),
        CEROS,
        [_ventas_detalle(escenario)],
        final_summary=_resumenes_finales()[0],
    )
    pedidas: list[set[str] | None] = []
    original = join_mod._juntar

    def espia(a, tablas, pos, solo):
        pedidas.append(solo)
        return original(a, tablas, pos, solo)

    monkeypatch.setattr(join_mod, "_juntar", espia)
    run_join(plan, escenario)
    # la última llamada es la de la unión; las anteriores son las muestras del diagnóstico
    assert pedidas[-1] == {"Región", "Importe"}
