"""Tests de core/aggregate (F4-02)."""

from __future__ import annotations

import json
from decimal import Decimal

import pandas as pd
import pytest

from excel_tool.core.aggregate import (
    Aggregation,
    AggregationSpec,
    GroupBy,
    default_name,
    result_columns,
    summarize,
)
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io import WriteOptions, read_table, write_table
from excel_tool.core.normalize import KeyOptions, normalize_keys
from excel_tool.core.progress import CancelToken, ProgressEvent

_STR = pd.StringDtype("pyarrow")


def _tabla(**cols):
    return pd.DataFrame({k: pd.Series(v, dtype=_STR) for k, v in cols.items()})


def _una(df, group, func, col="v", **kw):
    spec = AggregationSpec([GroupBy(group)] if group else [], [Aggregation(col, func)])
    return summarize(df, spec, **kw)


def _dec(texto: str) -> Decimal:
    return Decimal(texto.replace(",", "."))


def _a_dict(res, clave, col):
    return dict(zip(res.df[clave], res.df[col], strict=True))


# --------------------------------------------------------------- escenario 3


@pytest.fixture(params=["csv", "xlsx"])
def ventas(request, fixtures_generados, tmp_path):
    df = read_table(fixtures_generados / "escenario_3" / "ventas.csv").df
    if request.param == "xlsx":
        destino = tmp_path / "ventas.xlsx"
        write_table(df, destino, WriteOptions(format="xlsx"))
        df = read_table(destino).df
    return df


@pytest.fixture(scope="module")
def esperado_ventas(fixtures_esperados):
    ruta = fixtures_esperados / "escenario_3" / "resumen_ventas.json"
    return json.loads(ruta.read_text(encoding="utf-8"))


def test_ventas_por_cliente_contra_esperado(ventas, esperado_ventas):
    spec = AggregationSpec(
        [GroupBy("ID cliente")],
        [
            Aggregation("Importe", "sum"),
            Aggregation("", "count"),
            Aggregation("Artículo", "count_distinct"),
            Aggregation("Fecha", "max"),
        ],
    )
    res = summarize(ventas, spec)
    esp = esperado_ventas["por_cliente"]
    assert list(res.df.columns) == [
        "ID cliente",
        "Suma de Importe",
        "Cantidad de filas",
        "Distintos de Artículo",
        "Máximo de Fecha",
    ]
    assert len(res.df) == len(esp)
    # Los clientes fuera de clientes.xlsx (98, 99) no están en el esperado: se ignoran.
    for _, fila in res.df.iterrows():
        e = esp.get(fila["ID cliente"])
        if e is None:
            continue
        assert _dec(fila["Suma de Importe"]) == Decimal(e["suma_importe"])
        assert int(fila["Cantidad de filas"]) == e["filas"]
        assert int(fila["Distintos de Artículo"]) == e["articulos_distintos"]
        assert fila["Máximo de Fecha"] == e["ultima_fecha"]
    assert res.column_types["Máximo de Fecha"] == "fecha"
    assert res.column_types["Cantidad de filas"] == "entero"
    assert res.column_types["Suma de Importe"] == "decimal"


def test_invalidos_3_mas_un_vacio(ventas):
    res = _una(ventas, "ID cliente", "sum", "Importe")
    assert res.not_numeric == {"Importe": 3}
    assert any("3 celda" in w and "número" in w for w in res.warnings)
    assert any("1 celda" in w and "vacía" in w for w in res.warnings)
    # el cliente 7 tiene el vacío: suma lo válido
    assert _dec(_a_dict(res, "ID cliente", "Suma de Importe")["7"]) == Decimal("17097.46")


def test_ventas_por_cliente_y_articulo(ventas):
    spec = AggregationSpec(
        [GroupBy("ID cliente"), GroupBy("Artículo")], [Aggregation("Importe", "sum")]
    )
    res = summarize(ventas, spec)
    # esperado calculado aparte, con Decimal
    esp: dict[tuple[str, str], Decimal] = {}
    orden: list[tuple[str, str]] = []
    for cli, art, imp in zip(
        ventas["ID cliente"], ventas["Artículo"], ventas["Importe"], strict=True
    ):
        clave = (cli, art)
        if clave not in esp:
            esp[clave] = Decimal(0)
            orden.append(clave)
        try:
            esp[clave] += Decimal(imp.replace(".", "").replace(",", "."))
        except Exception:
            pass
    assert len(res.df) == len(esp)
    obtenido = [(c, a) for c, a in zip(res.df["ID cliente"], res.df["Artículo"], strict=True)]
    assert obtenido == orden  # orden de primera aparición
    for (c, a), v in zip(obtenido, res.df["Suma de Importe"], strict=True):
        assert _dec(v) == esp[(c, a)]


@pytest.fixture(params=["xlsx", "csv"])
def cobranzas(request, fixtures_generados, tmp_path):
    df = read_table(fixtures_generados / "escenario_3" / "cobranzas.xlsx").df
    if request.param == "csv":
        destino = tmp_path / "cobranzas.csv"
        write_table(df, destino, WriteOptions(format="csv_excel"))
        df = read_table(destino).df
    return df


def test_cobranzas_por_cliente(cobranzas, fixtures_esperados):
    esp = json.loads(
        (fixtures_esperados / "escenario_3" / "resumen_cobranzas.json").read_text(encoding="utf-8")
    )["por_cliente"]
    df = cobranzas.copy()
    df["cliente"] = normalize_keys(df["Nro Cliente"], KeyOptions(strip_leading_zeros=True))
    columna_monto = next(c for c in df.columns if c.lower().startswith("monto"))
    spec = AggregationSpec(
        [GroupBy("cliente")], [Aggregation("", "count"), Aggregation(columna_monto, "sum")]
    )
    res = summarize(df, spec)
    assert len(res.df) == len(esp)
    for _, fila in res.df.iterrows():
        e = esp[fila["cliente"]]
        assert int(fila["Cantidad de filas"]) == e["filas"]
        assert _dec(fila[f"Suma de {columna_monto}"]) == Decimal(e["suma_monto"])


def test_resumen_final_por_region_con_tabla_armada(ventas, fixtures_generados, fixtures_esperados):
    esp = json.loads(
        (fixtures_esperados / "escenario_3" / "resumen_final_por_region.json").read_text(
            encoding="utf-8"
        )
    )
    clientes = read_table(fixtures_generados / "escenario_3" / "clientes.xlsx").df
    por_cliente = summarize(
        ventas,
        AggregationSpec([GroupBy("ID cliente")], [Aggregation("Importe", "sum", "Total comprado")]),
    ).df
    unida = clientes.merge(por_cliente, on="ID cliente", how="left")
    unida["Total comprado"] = unida["Total comprado"].fillna("0").astype(_STR)
    final = summarize(
        unida,
        AggregationSpec(
            [GroupBy("Región")],
            [Aggregation("Total comprado", "sum"), Aggregation("", "count", "Clientes")],
        ),
    )
    por_region = _a_dict(final, "Región", "Suma de Total comprado")
    assert {r: _dec(v) for r, v in por_region.items()} == {
        r: Decimal(v) for r, v in esp["total_importe_por_region"].items()
    }
    clientes_region = _a_dict(final, "Región", "Clientes")
    assert {r: int(v) for r, v in clientes_region.items()} == esp["clientes_por_region"]


# --------------------------------------------------------------- operaciones


def test_sum_y_formato_argentino():
    df = _tabla(g=["a", "a", "b"], v=["1.234,50", "0,25", "10"])
    res = _una(df, "g", "sum")
    assert res.df["Suma de v"].tolist() == ["1234,75", "10,00"]
    assert res.column_types["Suma de v"] == "decimal"


def test_sum_enteros_sin_decimales_y_exacta():
    df = _tabla(g=["a", "a", "b"], v=["1", "2", "5"])
    res = _una(df, "g", "sum")
    assert res.df["Suma de v"].tolist() == ["3", "5"]
    assert res.column_types["Suma de v"] == "entero"
    df = _tabla(g=["a", "a"], v=["0,1", "0,2"])
    assert _una(df, "g", "sum").df["Suma de v"].tolist() == ["0,30"]


def test_sum_negativos_y_estilo_ingles():
    df = _tabla(g=["a", "a"], v=["-1,000.50", "0.25"])
    assert _una(df, "g", "sum").df["Suma de v"].tolist() == ["-1000,25"]
    df = _tabla(g=["a"], v=["-0,5"])
    assert _una(df, "g", "sum").df["Suma de v"].tolist() == ["-0,50"]


def test_sum_grupo_sin_validos_da_cero():
    df = _tabla(g=["a", "b"], v=["x", "5"])
    res = _una(df, "g", "sum")
    assert res.df["Suma de v"].tolist() == ["0", "5"]
    assert res.not_numeric == {"v": 1}


def test_count_cuenta_filas():
    df = _tabla(g=["a", "a", "b"], v=["", "x", "y"])
    res = summarize(df, AggregationSpec([GroupBy("g")], [Aggregation("", "count")]))
    assert res.df["Cantidad de filas"].tolist() == ["2", "1"]
    assert res.column_types["Cantidad de filas"] == "entero"


def test_count_distinct_ignora_vacios():
    df = _tabla(g=["a", "a", "a", "b"], v=["x", "x", "", "  "])
    res = _una(df, "g", "count_distinct")
    assert res.df["Distintos de v"].tolist() == ["1", "0"]


def test_mean():
    df = _tabla(g=["a", "a", "b", "c"], v=["1", "2", "3,5", "x"])
    res = _una(df, "g", "mean")
    assert res.df["Promedio de v"].tolist() == ["1,50", "3,50", ""]
    df = _tabla(g=["a", "a", "a"], v=["1", "1", "2"])
    assert _una(df, "g", "mean").df["Promedio de v"].tolist() == ["1,333333"]


def test_min_max_numeros_no_alfabeticos():
    df = _tabla(g=["a", "a", "a"], v=["9", "10", "100"])
    assert _una(df, "g", "min").df["Mínimo de v"].tolist() == ["9"]
    assert _una(df, "g", "max").df["Máximo de v"].tolist() == ["100"]


def test_min_max_fechas():
    df = _tabla(g=["a", "a", "b"], v=["15/03/2024", "02/12/2023", ""])
    mn = _una(df, "g", "min")
    mx = _una(df, "g", "max")
    assert mn.df["Mínimo de v"].tolist() == ["02/12/2023", ""]
    assert mx.df["Máximo de v"].tolist() == ["15/03/2024", ""]
    assert mx.column_types["Máximo de v"] == "fecha"


def test_min_max_texto():
    df = _tabla(g=["a", "a", "a"], v=["pera", "manzana", ""])
    assert _una(df, "g", "min").df["Mínimo de v"].tolist() == ["manzana"]
    assert _una(df, "g", "max").df["Máximo de v"].tolist() == ["pera"]


def test_first_last_con_contenido():
    df = _tabla(g=["a", "a", "a", "b"], v=["", "uno", "dos", ""])
    assert _una(df, "g", "first").df["Primero de v"].tolist() == ["uno", ""]
    assert _una(df, "g", "last").df["Último de v"].tolist() == ["dos", ""]


def test_first_last_numero_y_fecha_con_formato():
    df = _tabla(g=["a", "a"], v=["1.234,5", "2"])
    res = _una(df, "g", "first")
    assert res.df["Primero de v"].tolist() == ["1234,50"]
    assert res.column_types["Primero de v"] == "decimal"
    assert _una(df, "g", "last").df["Último de v"].tolist() == [
        "2"
    ]  # el tipo sale de los valores elegidos
    df = _tabla(g=["a", "a"], v=["1/3/2024", "05/04/2024"])
    res = _una(df, "g", "first")
    assert res.df["Primero de v"].tolist() == ["01/03/2024"]
    assert res.column_types["Primero de v"] == "fecha"


def test_join_text_sin_repetidos_en_orden():
    df = _tabla(g=["a", "a", "a", "b", "a"], v=["B", "A", "B", "Z", "C"])
    res = _una(df, "g", "join_text")
    assert res.df["Lista de v"].tolist() == ["B, A, C", "Z"]


def test_join_text_tope_de_largo():
    valores = [f"valor numero {i:04d}" for i in range(200)]
    df = _tabla(g=["a"] * 200, v=valores)
    res = _una(df, "g", "join_text")
    texto = res.df["Lista de v"].iloc[0]
    assert len(texto) <= 1000
    assert texto.endswith("…")
    assert texto.startswith("valor numero 0000, valor numero 0001")
    assert any("1.000" in w for w in res.warnings)


# --------------------------------------------------------------- grupos


def test_agrupar_por_mes_y_por_anio():
    df = _tabla(
        f=["05/01/2024", "20/01/2024", "03/02/2024", "x", "", "10/01/2025"],
        v=["1", "2", "3", "4", "5", "6"],
    )
    spec = AggregationSpec([GroupBy("f", "month")], [Aggregation("v", "sum")])
    res = summarize(df, spec)
    assert list(res.df.columns) == ["f (mes)", "Suma de v"]
    assert res.df["f (mes)"].tolist() == ["2024-01", "2024-02", "2025-01", "(vacío)", "(sin fecha)"]
    assert res.df["Suma de v"].tolist() == ["3", "3", "6", "5", "4"]
    assert any("(sin fecha)" in w for w in res.warnings)
    spec = AggregationSpec([GroupBy("f", "year")], [Aggregation("", "count")])
    res = summarize(df, spec)
    assert list(res.df.columns) == ["f (año)", "Cantidad de filas"]
    assert res.df["f (año)"].tolist() == ["2024", "2025", "(vacío)", "(sin fecha)"]
    assert res.df["Cantidad de filas"].tolist() == ["3", "1", "1", "1"]


def test_anio_limpio_es_entero():
    df = _tabla(f=["05/01/2024", "10/01/2025"])
    res = summarize(df, AggregationSpec([GroupBy("f", "year")], [Aggregation("", "count")]))
    assert res.column_types["f (año)"] == "entero"


def test_clave_vacia_va_a_vacio_y_orden_de_aparicion():
    df = _tabla(g=["b", "", "a", " ", "b"], v=["1"] * 5)
    res = summarize(df, AggregationSpec([GroupBy("g")], [Aggregation("", "count")]))
    assert res.df["g"].tolist() == ["b", "(vacío)", "a"]
    assert res.df["Cantidad de filas"].tolist() == ["2", "2", "1"]
    assert res.column_types["g"] == "texto"


def test_sin_group_by_una_sola_fila():
    df = _tabla(v=["1", "2", "x"])
    res = summarize(df, AggregationSpec([], [Aggregation("v", "sum"), Aggregation("", "count")]))
    assert len(res.df) == 1
    assert res.df.iloc[0].tolist() == ["3", "3"]


def test_sin_group_by_ni_calculos_es_error():
    with pytest.raises(ExcelToolError, match="Elegí por qué columnas agrupar"):
        summarize(_tabla(v=["1"]), AggregationSpec([], []))


def test_solo_group_by_da_combinaciones():
    df = _tabla(g=["a", "b", "a"])
    res = summarize(df, AggregationSpec([GroupBy("g")], []))
    assert res.df["g"].tolist() == ["a", "b"]


def test_columna_inexistente_y_falta_columna():
    with pytest.raises(ExcelToolError):
        summarize(_tabla(v=["1"]), AggregationSpec([GroupBy("zz")], []))
    with pytest.raises(ExcelToolError):
        summarize(_tabla(v=["1"]), AggregationSpec([], [Aggregation("", "sum")]))


def test_tabla_vacia():
    df = _tabla(g=[], v=[])
    res = summarize(df, AggregationSpec([GroupBy("g")], [Aggregation("v", "sum")]))
    assert len(res.df) == 0
    assert list(res.df.columns) == ["g", "Suma de v"]


def test_column_types_dados_no_se_detectan():
    df = _tabla(g=["a", "a"], v=["10", "9"])
    res = _una(df, "g", "max", column_types={"v": "texto", "g": "texto"})
    assert res.df["Máximo de v"].tolist() == ["9"]  # como texto "9" > "10"


# --------------------------------------------------------------- nombres


def test_nombres_automaticos():
    casos = {
        "sum": "Suma de Importe",
        "count_distinct": "Distintos de Importe",
        "mean": "Promedio de Importe",
        "min": "Mínimo de Importe",
        "max": "Máximo de Importe",
        "first": "Primero de Importe",
        "last": "Último de Importe",
        "join_text": "Lista de Importe",
    }
    for func, nombre in casos.items():
        assert default_name(Aggregation("Importe", func)) == nombre
    assert default_name(Aggregation("", "count")) == "Cantidad de filas"


def test_nombres_repetidos_y_explicitos():
    spec = AggregationSpec(
        [GroupBy("g")],
        [
            Aggregation("v", "sum"),
            Aggregation("v", "sum"),
            Aggregation("v", "sum", "Total"),
            Aggregation("", "count", "g"),
        ],
    )
    assert result_columns(spec) == ["g", "Suma de v", "Suma de v (2)", "Total", "g (2)"]
    res = summarize(_tabla(g=["a"], v=["1"]), spec)
    assert list(res.df.columns) == result_columns(spec)
    assert set(res.column_types) == set(res.df.columns)


# --------------------------------------------------------------- progreso y cancelar


def test_progreso_y_cancelar():
    df = _tabla(g=["a"], v=["1"])
    eventos: list[ProgressEvent] = []
    _una(df, "g", "sum", progress=eventos.append)
    assert eventos and eventos[-1].done == eventos[-1].total
    token = CancelToken()
    token.cancel()
    from excel_tool.core.errors import Cancelado

    with pytest.raises(Cancelado):
        _una(df, "g", "sum", cancel=token)


def test_periodos_en_orden_cronologico_con_meses_mezclados():
    df = _tabla(
        f=["15/03/2024", "x", "10/01/2025", "02/01/2024", "", "20/12/2023", "05/03/2024"],
        v=["1", "2", "3", "4", "5", "6", "7"],
    )
    res = summarize(df, AggregationSpec([GroupBy("f", "month")], [Aggregation("v", "sum")]))
    assert res.df["f (mes)"].tolist() == [
        "2023-12",
        "2024-01",
        "2024-03",
        "2025-01",
        "(vacío)",
        "(sin fecha)",
    ]
    assert res.df["Suma de v"].tolist() == ["6", "4", "8", "3", "5", "2"]


def test_periodos_por_anio_en_orden_cronologico():
    df = _tabla(f=["01/01/2025", "01/01/2023", "x", "01/01/2024"], v=["1", "1", "1", "1"])
    res = summarize(df, AggregationSpec([GroupBy("f", "year")], [Aggregation("", "count")]))
    assert res.df["f (año)"].tolist() == ["2023", "2024", "2025", "(sin fecha)"]


def test_periodo_y_otra_columna_conservan_aparicion_dentro_del_periodo():
    df = _tabla(
        f=["01/02/2024", "01/01/2024", "02/02/2024", "02/01/2024", "03/01/2024"],
        c=["B", "Z", "A", "A", "B"],
        v=["1"] * 5,
    )
    spec = AggregationSpec([GroupBy("c"), GroupBy("f", "month")], [Aggregation("v", "count")])
    res = summarize(df, spec)
    assert list(zip(res.df["f (mes)"], res.df["c"], strict=True)) == [
        ("2024-01", "Z"),
        ("2024-01", "A"),
        ("2024-01", "B"),
        ("2024-02", "B"),
        ("2024-02", "A"),
    ] or list(zip(res.df["c"], res.df["f (mes)"], strict=True)) == [
        ("Z", "2024-01"),
        ("A", "2024-01"),
        ("B", "2024-01"),
        ("B", "2024-02"),
        ("A", "2024-02"),
    ]


def test_grupos_sin_periodo_mantienen_primera_aparicion():
    df = _tabla(c=["B", "A", "B", "C"], v=["1"] * 4)
    res = summarize(df, AggregationSpec([GroupBy("c")], [Aggregation("v", "count")]))
    assert res.df["c"].tolist() == ["B", "A", "C"]


def test_aviso_celdas_vacias_singular_y_plural():
    uno = summarize(
        _tabla(c=["a", "a"], v=["1", ""]),
        AggregationSpec([GroupBy("c")], [Aggregation("v", "sum")]),
    )
    assert "1 celda vacía de v: no se suma" in " ".join(uno.warnings)
    tres = summarize(
        _tabla(c=["a"] * 4, v=["1", "", "", ""]),
        AggregationSpec([GroupBy("c")], [Aggregation("v", "sum")]),
    )
    assert "3 celdas vacías de v: no se suman" in " ".join(tres.warnings)


def test_aviso_celdas_ilegibles_y_fechas_singular_y_plural():
    spec = AggregationSpec([GroupBy("c")], [Aggregation("v", "sum")])
    uno = summarize(_tabla(c=["a"] * 3, v=["1", "2", "zz"]), spec)
    assert "1 celda con contenido que no se pudo leer como números: no se suma" in " ".join(
        uno.warnings
    )
    dos = summarize(_tabla(c=["a"] * 4, v=["1", "2", "zz", "yy"]), spec)
    assert "2 celdas con contenido que no se pudieron leer como números: no se suman" in " ".join(
        dos.warnings
    )
    f1 = summarize(
        _tabla(f=["01/01/2024", "x"], v=["1", "1"]),
        AggregationSpec([GroupBy("f", "month")], [Aggregation("v", "sum")]),
    )
    assert any("1 fecha que no se pudo leer" in w for w in f1.warnings)
    f2 = summarize(
        _tabla(f=["01/01/2024", "x", "y"], v=["1"] * 3),
        AggregationSpec([GroupBy("f", "month")], [Aggregation("v", "sum")]),
    )
    assert any("2 fechas que no se pudieron leer" in w for w in f2.warnings)
