"""Integración de la Fase 4: leer -> diagnosticar -> unir -> guardar -> releer, con el core REAL.

Escenario de 3 archivos (`clientes.xlsx` principal, `ventas.csv` y `cobranzas.xlsx`), con las ventas
también como xlsx. Los resultados esperados salen de `tests/fixtures/esperados/escenario_3/*.json`
(calculados por el generador desde la definición de los datos) o se recuentan acá con el módulo
`csv` y `Decimal`, sin pasar por el código que se prueba.
"""

from __future__ import annotations

import csv
import dataclasses
import json
import re
import shutil
import subprocess
from collections import Counter, defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy, summarize
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io import ReadOptions, WriteOptions, read_table, write_table
from excel_tool.core.join import (
    JoinPlan,
    SecondarySpec,
    SourceSpec,
    diagnose_join,
    run_join,
)
from excel_tool.core.normalize import KeyOptions
from excel_tool.core.output import OutputSpec, save_result
from excel_tool.core.recipe import load_recipe, save_recipe
from excel_tool.core.split import SplitSpec

RO = ReadOptions()
CEROS = KeyOptions(strip_leading_zeros=True)
SIN_CEROS = KeyOptions()
RAIZ = Path(__file__).resolve().parents[2]


# ---- utilidades


def _dec(texto: str) -> Decimal:
    """'2.473,77' / '922,24' / '922.24' -> Decimal."""
    t = texto.strip()
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    return Decimal(t)


def _es_importe(texto: str) -> bool:
    return bool(re.fullmatch(r"\d{1,3}(\.\d{3})*,\d{2}|\d+,\d{2}", texto))


def _esperado(fixtures_esperados: Path, nombre: str) -> dict:
    ruta = fixtures_esperados / "escenario_3" / f"{nombre}.json"
    return json.loads(ruta.read_text(encoding="utf-8"))


@dataclass
class Escenario:
    formato: str
    carpeta: Path
    tablas: dict[str, pd.DataFrame]
    rutas: dict[str, Path]
    esp_ventas: dict
    esp_cobranzas: dict
    esp_diag: dict
    esp_region: dict
    ventas_csv: list[dict]  # recuento independiente con el módulo csv

    def principal(self) -> SourceSpec:
        df = self.tablas["clientes"]
        return SourceSpec("clientes", self.rutas["clientes"], RO, list(df.columns))

    def ventas(self, **kw) -> SecondarySpec:
        return SecondarySpec(
            "ventas",
            self.rutas["ventas"],
            RO,
            ["Fecha", "Artículo", "Cantidad", "Importe"],
            [("ID cliente", "ID cliente")],
            **kw,
        )

    def cobranzas(self, **kw) -> SecondarySpec:
        return SecondarySpec(
            "cobranzas",
            self.rutas["cobranzas"],
            RO,
            ["Fecha", "Medio de pago", "Monto"],
            [("ID cliente", "Nro Cliente")],
            **kw,
        )


def resumen_ventas_por_cliente(**kw) -> dict:
    return {
        "mode": "summary",
        "summary": AggregationSpec(
            [GroupBy("ID cliente")],
            [
                Aggregation("Importe", "sum", "Total comprado"),
                Aggregation("", "count", "Compras"),
                Aggregation("Artículo", "count_distinct", "Artículos distintos"),
                Aggregation("Fecha", "max", "Última compra"),
            ],
        ),
        **kw,
    }


def resumen_cobranzas_por_cliente(**kw) -> dict:
    return {
        "mode": "summary",
        "summary": AggregationSpec(
            [GroupBy("Nro Cliente")],
            [Aggregation("Monto", "sum", "Total cobrado"), Aggregation("", "count", "Cobros")],
        ),
        **kw,
    }


@pytest.fixture(params=["csv", "xlsx"])
def esc(request, fixtures_generados, fixtures_esperados, tmp_path) -> Escenario:
    """`csv`: ventas.csv tal cual. `xlsx`: las mismas ventas guardadas como xlsx."""
    base = fixtures_generados / "escenario_3"
    rutas = {
        "clientes": base / "clientes.xlsx",
        "ventas": base / "ventas.csv",
        "cobranzas": base / "cobranzas.xlsx",
    }
    if request.param == "xlsx":
        crudas = read_table(rutas["ventas"]).df
        rutas["ventas"] = tmp_path / "ventas.xlsx"
        write_table(crudas, rutas["ventas"], WriteOptions(format="xlsx"))
    tablas = {n: read_table(r).df for n, r in rutas.items()}
    with (base / "ventas.csv").open(encoding="cp1252", newline="") as f:
        lector = csv.reader(f, delimiter=";")
        next(lector)
        ventas_csv = [
            {"id": r[0], "fecha": r[1], "articulo": r[2], "cantidad": r[3], "importe": r[4]}
            for r in lector
        ]
    return Escenario(
        request.param,
        tmp_path,
        tablas,
        rutas,
        _esperado(fixtures_esperados, "resumen_ventas"),
        _esperado(fixtures_esperados, "resumen_cobranzas"),
        _esperado(fixtures_esperados, "diagnostico"),
        _esperado(fixtures_esperados, "resumen_final_por_region"),
        ventas_csv,
    )


def _guardar_y_releer(res, destino: Path, formato: str = "xlsx") -> pd.DataFrame:
    """Guarda con los tipos del join y relee. Verifica que el contenido sea el mismo."""
    spec = OutputSpec(
        destino,
        "xlsx" if formato == "xlsx" else "csv_excel",
        "excel",
        column_types=res.column_types,
    )
    guardado = save_result(res.df, spec)
    assert guardado.rows_total == len(res.df) == sum(p.rows for p in guardado.parts)
    releido = read_table(guardado.files[0]).df
    assert list(releido.columns) == list(res.df.columns)
    assert len(releido) == len(res.df)
    for col in res.df.columns:
        tipo = res.column_types.get(col)
        for a, b in zip(res.df[col].fillna(""), releido[col].fillna(""), strict=True):
            if tipo in ("decimal", "entero") and a != "":
                assert _dec(a) == _dec(b), (col, a, b)
            else:
                assert a == b, (col, a, b)
    return releido


def _fila_por_cliente(df: pd.DataFrame, col: str = "ID cliente") -> dict[str, pd.DataFrame]:
    return {k: g for k, g in df.groupby(col, sort=False)}


# ---- 1. detalle + resumen


def test_detalle_de_ventas_mas_resumen_de_cobranzas(esc, tmp_path):
    plan = JoinPlan(
        esc.principal(), CEROS, [esc.ventas(), esc.cobranzas(**resumen_cobranzas_por_cliente())]
    )
    diag = diagnose_join(plan, esc.tablas)
    esperadas = esc.esp_diag["cobranzas"]["sin_ceros_a_la_izquierda"][
        "filas_resultado_ventas_detalle_y_cobranzas_resumen"
    ]
    assert diag.rows_result == esperadas == esc.esp_diag["ventas"]["filas_resultado_detalle"] == 88
    assert not diag.many_details

    res = run_join(plan, esc.tablas)
    df = res.df
    assert len(df) == diag.rows_result == 88
    # Una fila por compra (más una por cliente sin compras) y cada una con el total cobrado.
    por_cliente = _fila_por_cliente(df)
    assert set(por_cliente) == {c["id"] for c in _clientes(esc)}
    for cid, grupo in por_cliente.items():
        esperadas_filas = esc.esp_ventas["por_cliente"].get(cid, {"filas": 1})["filas"]
        assert len(grupo) == esperadas_filas, cid
        cobro = esc.esp_cobranzas["por_cliente"].get(cid)
        for valor in grupo["Total cobrado"].fillna(""):
            if cobro is None:
                assert valor == "0", cid  # ADR-037: Sumar sin pareja = 0
            else:
                assert _dec(valor) == Decimal(cobro["suma_monto"]), cid
        for valor in grupo["Cobros"].fillna(""):
            assert valor == ("0" if cobro is None else str(cobro["filas"])), cid
    # Las compras salen todas, con el importe original (recuento independiente).
    importes = Counter(x for x in df["Importe"].fillna("") if x != "")
    propias = [v for v in esc.ventas_csv if v["id"] in por_cliente]
    assert importes == Counter(v["importe"] for v in propias if v["importe"] != "")
    assert [c for c in df.columns][:4] == ["ID cliente", "Nombre", "Región", "Vendedor"]
    _guardar_y_releer(res, tmp_path / "detalle_resumen.xlsx", "xlsx")
    _guardar_y_releer(res, tmp_path / "detalle_resumen.csv", "csv")


def _clientes(esc: Escenario) -> list[dict]:
    df = esc.tablas["clientes"]
    return [{"id": i, "region": r} for i, r in zip(df["ID cliente"], df["Región"], strict=True)]


# ---- 2. resumen por cliente


def test_resumen_de_ventas_por_cliente_mismas_filas_que_clientes(esc, tmp_path):
    plan = JoinPlan(esc.principal(), SIN_CEROS, [esc.ventas(**resumen_ventas_por_cliente())])
    diag = diagnose_join(plan, esc.tablas)
    res = run_join(plan, esc.tablas)
    df = res.df
    assert len(df) == len(esc.tablas["clientes"]) == 30 == diag.rows_result
    assert df["ID cliente"].tolist() == esc.tablas["clientes"]["ID cliente"].tolist()
    por = esc.esp_ventas["por_cliente"]
    for _, fila in df.iterrows():
        cid = fila["ID cliente"]
        if cid not in por:  # cliente sin compras (ADR-037): 0 en Sumar/Contar, vacío en el resto
            assert fila["Total comprado"] == "0" and fila["Compras"] == "0"
            assert fila["Artículos distintos"] == "0"
            continue
        e = por[cid]
        assert _dec(fila["Total comprado"]) == Decimal(e["suma_importe"]), cid
        assert int(fila["Compras"]) == e["filas"], cid
        assert int(fila["Artículos distintos"]) == e["articulos_distintos"], cid
        assert fila["Última compra"] == e["ultima_fecha"], cid
    assert {c for c in df["ID cliente"] if c not in por} == {"5", "17", "28"}
    assert res.column_types["Total comprado"] == "decimal"
    assert res.column_types["Última compra"] == "fecha"
    _guardar_y_releer(res, tmp_path / "res_ventas.xlsx", "xlsx")
    _guardar_y_releer(res, tmp_path / "res_ventas.csv", "csv")


# ---- 3. cliente + artículo


def test_resumen_de_ventas_por_cliente_y_articulo(esc):
    sec = esc.ventas(
        mode="summary",
        summary=AggregationSpec(
            [GroupBy("ID cliente"), GroupBy("Artículo")],
            [Aggregation("Importe", "sum", "Total comprado"), Aggregation("", "count", "Compras")],
        ),
    )
    plan = JoinPlan(esc.principal(), SIN_CEROS, [sec])
    diag = diagnose_join(plan, esc.tablas)
    res = run_join(plan, esc.tablas)
    # Recuento independiente (csv + Decimal) de cada par cliente / artículo.
    ids = {c["id"] for c in _clientes(esc)}
    pares: dict[tuple[str, str], dict] = defaultdict(lambda: {"suma": Decimal(0), "n": 0})
    for v in esc.ventas_csv:
        if v["id"] in ids:
            p = pares[(v["id"], v["articulo"])]
            p["n"] += 1
            if _es_importe(v["importe"]):
                p["suma"] += _dec(v["importe"])
    sin_ventas = ids - {k for k, _ in pares}
    assert sin_ventas == {"5", "17", "28"}
    esperadas = len(pares) + len(sin_ventas)
    assert len(res.df) == esperadas == diag.rows_result  # ADR-036: filas a propósito, las estimó
    assert len(res.df) > 30
    vistos = {}
    for _, f in res.df.iterrows():
        if f["ID cliente"] in sin_ventas:
            assert not f["Artículo"] and f["Total comprado"] == "0" and f["Compras"] == "0"
            continue
        vistos[(f["ID cliente"], f["Artículo"])] = (_dec(f["Total comprado"]), int(f["Compras"]))
    assert vistos == {k: (v["suma"], v["n"]) for k, v in pares.items()}
    # el diagnóstico avisa que cambian las filas respecto del principal
    assert any("filas" in w for w in diag.warnings) or diag.rows_result != 30


# ---- 4. resumen final


def test_resumen_final_por_region(esc, tmp_path):
    final = AggregationSpec(
        [GroupBy("Región")],
        [
            Aggregation("Total comprado", "sum", "Total vendido"),
            Aggregation("ID cliente", "count_distinct", "Clientes"),
        ],
    )
    plan = JoinPlan(
        esc.principal(),
        SIN_CEROS,
        [esc.ventas(**resumen_ventas_por_cliente())],
        final_summary=final,
    )
    diag = diagnose_join(plan, esc.tablas)
    res = run_join(plan, esc.tablas)
    assert len(res.df) == 5 == diag.rows_result
    total = {f["Región"]: _dec(f["Total vendido"]) for _, f in res.df.iterrows()}
    esperado = {r: Decimal(v) for r, v in esc.esp_region["total_importe_por_region"].items()}
    assert total == esperado
    clientes = {f["Región"]: int(f["Clientes"]) for _, f in res.df.iterrows()}
    assert clientes == esc.esp_region["clientes_por_region"]
    # Con lo excluido (ventas de IDs que no son clientes) suma todas las compras válidas.
    todas = sum(
        (_dec(v["importe"]) for v in esc.ventas_csv if _es_importe(v["importe"])), Decimal(0)
    )
    assert (
        sum(total.values()) + Decimal(esc.esp_region["importe_excluido_ids_fuera_de_clientes"])
        == todas
    )
    _guardar_y_releer(res, tmp_path / "por_region.xlsx", "xlsx")


def test_ventas_agrupadas_por_mes_salen_en_orden_cronologico(esc):
    """F4-B: los meses salen de enero a junio (no por orden de aparición); recuento con csv."""
    res = summarize(
        esc.tablas["ventas"],
        AggregationSpec([GroupBy("Fecha", "month")], [Aggregation("", "count")]),
    )
    col = res.df.columns[0]
    etiquetas = res.df[col].tolist()
    fechas = [v["fecha"] for v in esc.ventas_csv if v["fecha"]]
    esperados = sorted({"-".join(reversed(f.split("/")[1:])) for f in fechas})
    sin_fecha = [x for x in etiquetas if x in ("(sin fecha)", "(vacío)")]
    assert [x for x in etiquetas if x not in sin_fecha] == esperados
    assert etiquetas[len(etiquetas) - len(sin_fecha) :] == sin_fecha  # los vacíos, al final
    assert sum(int(x) for x in res.df.iloc[:, 1]) == len(esc.ventas_csv)
    # Cruzando de año y con fecha mala: dic 2023 antes que ene 2024; "(sin fecha)" al final.
    chica = pd.DataFrame(
        {"Fecha": ["15/02/2024", "x", "10/12/2023", "03/01/2024", "20/11/2023"]}, dtype="string"
    )
    r2 = summarize(chica, AggregationSpec([GroupBy("Fecha", "month")], [Aggregation("", "count")]))
    assert r2.df.iloc[:, 0].tolist()[:4] == ["2023-11", "2023-12", "2024-01", "2024-02"]
    assert r2.df.iloc[:, 0].tolist()[-1] == "(sin fecha)"


# ---- 5. dos en detalle


def test_dos_archivos_en_detalle_avisa_y_estima_bien(esc):
    plan = JoinPlan(esc.principal(), CEROS, [esc.ventas(), esc.cobranzas()])
    diag = diagnose_join(plan, esc.tablas)
    esperadas = esc.esp_diag["cobranzas"]["sin_ceros_a_la_izquierda"][
        "filas_resultado_ventas_detalle_y_cobranzas_detalle"
    ]
    assert diag.many_details is True
    assert diag.rows_result == esperadas
    assert esperadas > esc.esp_diag["ventas"]["filas_resultado_detalle"]
    assert any("detalle" in w.lower() for w in diag.warnings), diag.warnings
    res = run_join(plan, esc.tablas)
    assert len(res.df) == esperadas
    assert any("detalle" in w.lower() for w in res.warnings)
    # filas por cliente = filas de ventas x filas de cobranzas (cada lado al menos 1)
    cuenta_v = Counter(v["id"] for v in esc.ventas_csv)
    cuenta_c = {k: e["filas"] for k, e in esc.esp_cobranzas["por_cliente"].items()}
    for cid, grupo in _fila_por_cliente(res.df).items():
        assert len(grupo) == max(1, cuenta_v.get(cid, 0)) * max(1, cuenta_c.get(cid, 0)), cid


def test_dos_archivos_en_detalle_sin_la_opcion_de_ceros(esc):
    plan = JoinPlan(esc.principal(), SIN_CEROS, [esc.ventas(), esc.cobranzas()])
    diag = diagnose_join(plan, esc.tablas)
    esperadas = esc.esp_diag["cobranzas"]["solo_recortar_espacios"][
        "filas_resultado_ventas_detalle_y_cobranzas_detalle"
    ]
    assert diag.rows_result == esperadas
    assert len(run_join(plan, esc.tablas).df) == esperadas


# ---- 6. ceros (ADR-034)


def test_ceros_a_la_izquierda_sin_y_con_la_opcion(esc):
    sin = esc.esp_diag["cobranzas"]["solo_recortar_espacios"]
    con = esc.esp_diag["cobranzas"]["sin_ceros_a_la_izquierda"]
    assert (sin["claves_con_coincidencia"], con["claves_con_coincidencia"]) == (5, 27)
    spec = esc.cobranzas(**resumen_cobranzas_por_cliente())

    d0 = diagnose_join(JoinPlan(esc.principal(), SIN_CEROS, [spec]), esc.tablas).secondaries[0]
    assert d0.keys_matched == 5
    assert d0.keys_matched_without_zeros == 27
    assert set(d0.keys_unmatched_sample) <= set(sin["claves_sin_coincidencia"])
    assert len(d0.keys_unmatched_sample) == min(10, 25)

    diag1 = diagnose_join(JoinPlan(esc.principal(), CEROS, [spec]), esc.tablas)
    d1 = diag1.secondaries[0]
    assert d1.keys_matched == 27
    assert d1.keys_matched_without_zeros is None
    assert sorted(d1.keys_unmatched_sample, key=int) == con["claves_sin_coincidencia"]
    assert d1.secondary_orphans == len(con["claves_del_secundario_sin_pareja"]) == 1

    diag0 = diagnose_join(JoinPlan(esc.principal(), SIN_CEROS, [spec]), esc.tablas)
    sugerencia = [w for w in diag0.warnings if "ceros" in w]
    assert len(sugerencia) == 1 and "27" in sugerencia[0] and "5" in sugerencia[0], diag0.warnings
    assert not [w for w in diag1.warnings if "ceros" in w]

    # Resultados: solo 5 clientes con total cobrado sin la opción; 27 con la opción.
    r0 = run_join(JoinPlan(esc.principal(), SIN_CEROS, [spec]), esc.tablas)
    r1 = run_join(JoinPlan(esc.principal(), CEROS, [spec]), esc.tablas)
    assert len(r0.df) == len(r1.df) == 30
    assert int((~r0.df["Total cobrado"].fillna("").isin(["", "0"])).sum()) == 5
    assert int((~r1.df["Total cobrado"].fillna("").isin(["", "0"])).sum()) == 27
    for _, f in r1.df.iterrows():
        e = esc.esp_cobranzas["por_cliente"].get(f["ID cliente"])
        if e is None:
            assert f["Total cobrado"] == "0"  # ADR-037
        else:
            assert _dec(f["Total cobrado"]) == Decimal(e["suma_monto"])
    # Sin la opción solo emparejan las filas cuya clave es exactamente el ID (con espacios): el
    # total de esos clientes es parcial. Se recuenta leyendo el xlsx con openpyxl.
    parcial: dict[str, Decimal] = defaultdict(Decimal)
    filas_parcial = 0
    hoja = openpyxl.load_workbook(esc.rutas["cobranzas"]).worksheets[0]
    for clave, _f, _m, monto in hoja.iter_rows(min_row=2, values_only=True):
        if str(clave).strip() == str(clave).strip().lstrip("0"):
            parcial[str(clave).strip()] += Decimal(str(monto))
            filas_parcial += 1
    vistos = {
        f["ID cliente"]: _dec(f["Total cobrado"])
        for _, f in r0.df.iterrows()
        if f["Total cobrado"] not in ("", "0")
    }
    assert vistos == dict(parcial) and len(vistos) == 5
    # Las filas de cobranzas sin pareja (para la hoja extra) son las que no emparejaron.
    total = esc.esp_cobranzas["filas_total"]
    assert len(r0.unmatched["cobranzas"]) == total - filas_parcial
    assert len(r1.unmatched["cobranzas"]) == 2  # las dos cobranzas de "00077"


# ---- 7. importes inválidos


def test_importes_invalidos_se_informan_y_el_vacio_aparte(esc):
    e = esc.esp_ventas
    invalidas = [c for c in e["celdas_invalidas"] if not c["vacia"]]
    vacias = [c for c in e["celdas_invalidas"] if c["vacia"]]
    assert (len(invalidas), len(vacias)) == (3, 1)
    assert sum(r["importes_invalidos"] for r in e["por_cliente"].values()) == 3
    assert sum(r["importes_vacios"] for r in e["por_cliente"].values()) == 1
    assert sorted(c["valor"] for c in invalidas) == ["12,34,5", "abc", "n/a"]

    directo = summarize(
        esc.tablas["ventas"],
        AggregationSpec([GroupBy("ID cliente")], [Aggregation("Importe", "sum")]),
    )
    assert directo.not_numeric == {"Importe": 3}
    assert any(
        re.search(r"\b3 celdas con contenido que no se pudieron leer", w) for w in directo.warnings
    )
    assert any(
        re.search(r"\bHay 1 celda vacía de Importe: no se suma", w) for w in directo.warnings
    )
    for _, f in directo.df.iterrows():
        if f["ID cliente"] in e["por_cliente"]:
            assert _dec(f["Suma de Importe"]) == Decimal(
                e["por_cliente"][f["ID cliente"]]["suma_importe"]
            )

    res = run_join(
        JoinPlan(esc.principal(), SIN_CEROS, [esc.ventas(**resumen_ventas_por_cliente())]),
        esc.tablas,
    )
    assert any("3 celda" in w for w in res.warnings), res.warnings
    # los inválidos no rompen la suma: cada total es el esperado
    for _, f in res.df.iterrows():
        if f["ID cliente"] in e["por_cliente"]:
            assert _dec(f["Total comprado"]) == Decimal(
                e["por_cliente"][f["ID cliente"]]["suma_importe"]
            )


# ---- 8. control de filas (ADR-036)


@pytest.mark.parametrize(
    "como",
    ["resumen_resumen", "resumen_primera", "primera", "resumen_con_prefijo"],
)
def test_sin_detalle_el_resultado_tiene_las_filas_del_principal(esc, como):
    """ADR-036: todos `left`, resúmenes solo por la clave (o `first`): filas del principal."""
    n = len(esc.tablas["clientes"])
    if como == "resumen_resumen":
        secundarios = [
            esc.ventas(**resumen_ventas_por_cliente()),
            esc.cobranzas(**resumen_cobranzas_por_cliente()),
        ]
    elif como == "resumen_primera":
        secundarios = [esc.ventas(**resumen_ventas_por_cliente()), esc.cobranzas(mode="first")]
    elif como == "primera":
        secundarios = [esc.ventas(mode="first"), esc.cobranzas(mode="first")]
    else:
        secundarios = [
            esc.ventas(**resumen_ventas_por_cliente(prefix="V")),
            esc.cobranzas(mode="first", prefix="C"),
        ]
    plan = JoinPlan(esc.principal(), CEROS, secundarios)
    diag = diagnose_join(plan, esc.tablas)
    res = run_join(plan, esc.tablas)
    assert diag.rows_result == len(res.df) == n
    assert res.df["ID cliente"].tolist() == esc.tablas["clientes"]["ID cliente"].tolist()
    assert not diag.many_details


def test_primera_fila_trae_la_primera_del_archivo(esc):
    res = run_join(JoinPlan(esc.principal(), SIN_CEROS, [esc.ventas(mode="first")]), esc.tablas)
    primera: dict[str, dict] = {}
    for v in esc.ventas_csv:
        primera.setdefault(v["id"], v)
    for _, f in res.df.iterrows():
        v = primera.get(f["ID cliente"])
        if v is None:
            assert not f["Artículo"]
        else:
            assert (f["Fecha"], f["Artículo"], f["Cantidad"]) == (
                v["fecha"],
                v["articulo"],
                v["cantidad"],
            )
            assert (f["Importe"] or "") == v["importe"]


def test_los_casos_que_cambian_las_filas_las_estiman_exacto(esc):
    ids = {c["id"] for c in _clientes(esc)}
    con_ventas = {v["id"] for v in esc.ventas_csv} & ids
    # inner con resumen: solo los clientes que tienen compras
    inner = JoinPlan(
        esc.principal(), SIN_CEROS, [esc.ventas(how="inner", **resumen_ventas_por_cliente())]
    )
    d = diagnose_join(inner, esc.tablas)
    r = run_join(inner, esc.tablas)
    assert d.rows_result == len(r.df) == len(con_ventas) == 27
    assert set(r.df["ID cliente"]) == con_ventas
    # inner con detalle: una fila por compra de clientes del principal
    inner_det = JoinPlan(esc.principal(), SIN_CEROS, [esc.ventas(how="inner")])
    d = diagnose_join(inner_det, esc.tablas)
    assert (
        d.rows_result
        == len(run_join(inner_det, esc.tablas).df)
        == sum(1 for v in esc.ventas_csv if v["id"] in ids)
    )


def test_el_control_de_filas_corta_aunque_la_estimacion_coincida(esc, monkeypatch):
    """Con una fila de más y la estimación coincidiendo, el control de ADR-036 igual corta.

    Se simula el error en `_unir` y en `_analizar`."""
    from excel_tool.core import join as join_mod

    plan = JoinPlan(esc.principal(), CEROS, [esc.ventas(**resumen_ventas_por_cliente())])
    unir, analizar = join_mod._unir, join_mod._analizar

    def unir_roto(pos, c, t, cancel):
        r = unir(pos, c, t, cancel)
        return join_mod._Posiciones(r.pidx, r.sidx, r.n + 1)

    def analizar_roto(*args, **kwargs):
        a = analizar(*args, **kwargs)
        a.rows_join += 1  # la "estimación" coincide con el error
        return a

    monkeypatch.setattr(join_mod, "_unir", unir_roto)
    monkeypatch.setattr(join_mod, "_analizar", analizar_roto)
    with pytest.raises(ExcelToolError) as e:
        run_join(plan, esc.tablas)
    assert "mismas filas" in e.value.mensaje and "No se guardó nada" in e.value.mensaje


# ---- 9. recetas


def _plan_con_salida(esc: Escenario, salida: Path) -> JoinPlan:
    return JoinPlan(
        esc.principal(),
        CEROS,
        [
            esc.ventas(**resumen_ventas_por_cliente()),
            esc.cobranzas(**resumen_cobranzas_por_cliente()),
        ],
        output_column_order=["ID cliente", "Nombre", "Total comprado", "Total cobrado"],
        output=OutputSpec(
            salida / "resultado.csv",
            "csv_excel",
            "excel",
            split=SplitSpec("rows", rows_per_part=12),
        ),
    )


def test_receta_guardada_y_cargada_da_el_mismo_resultado_con_division(esc, tmp_path):
    out1, out2 = tmp_path / "salida1", tmp_path / "salida2"
    out1.mkdir()
    out2.mkdir()
    plan = _plan_con_salida(esc, out1)
    receta = tmp_path / "recetas" / "mensual.json"
    receta.parent.mkdir()
    save_recipe(plan, receta)
    cargada = load_recipe(receta)
    assert cargada.missing == [] and cargada.warnings == []
    assert cargada.plan == plan

    r1 = run_join(plan, esc.tablas)
    r2 = run_join(cargada.plan, esc.tablas)
    assert r1.df.equals(r2.df) and r1.column_types == r2.column_types
    assert list(r1.df.columns)[:4] == ["ID cliente", "Nombre", "Total comprado", "Total cobrado"]
    g1 = save_result(r1.df, plan.output)
    g2 = save_result(r2.df, dataclasses.replace(cargada.plan.output, path=out2 / "resultado.csv"))
    assert len(g1.parts) == len(g2.parts) == 3  # 30 filas en partes de 12
    assert [(p.path.name, p.rows) for p in g1.parts] == [(p.path.name, p.rows) for p in g2.parts]
    assert g1.rows_total == g2.rows_total == 30
    for a, b in zip(g1.files, g2.files, strict=True):
        assert a.read_bytes() == b.read_bytes()


def test_receta_de_linux_se_abre_en_otra_carpeta_y_con_archivos_nuevos(esc, tmp_path):
    """Rutas con `/`, relativas si se puede: se puede mover toda la carpeta (Linux -> Windows)."""
    carpeta = tmp_path / "mes"
    carpeta.mkdir()
    for r in esc.rutas.values():
        shutil.copy2(r, carpeta / r.name)
    esc2 = dataclasses.replace(esc, rutas={n: carpeta / r.name for n, r in esc.rutas.items()})
    plan = _plan_con_salida(esc2, carpeta / "salida")
    save_recipe(plan, carpeta / "receta.json")
    texto = (carpeta / "receta.json").read_text(encoding="utf-8")
    assert "\\" not in texto and str(tmp_path) not in texto  # nada absoluto ni con `\`
    assert json.loads(texto)["version"] == 1

    movida = tmp_path / "otra_pc" / "Documentos"
    shutil.copytree(carpeta, movida)
    cargada = load_recipe(movida / "receta.json")
    assert cargada.missing == []
    assert cargada.plan.principal.path == movida / "clientes.xlsx"
    assert cargada.plan.output.path == movida / "salida" / "resultado.csv"
    assert run_join(cargada.plan, esc.tablas).df.equals(run_join(plan, esc.tablas).df)


@pytest.mark.parametrize("estilo", ["barra", "contrabarra", "unc"])
def test_receta_con_rutas_de_windows_carga_en_linux(esc, tmp_path, estilo):
    plan = _plan_con_salida(esc, tmp_path)
    receta = tmp_path / "windows.json"
    save_recipe(plan, receta)
    datos = json.loads(receta.read_text(encoding="utf-8"))

    def ruta(nombre: str) -> str:
        return {
            "barra": f"D:/Mis datos/Clientes y ventas/{nombre}",
            "contrabarra": f"C:\\Users\\Ana Pérez\\Documents\\{nombre}",
            "unc": f"\\\\servidor\\compartida\\{nombre}",
        }[estilo]

    p = datos["plan"]
    for fuente in [p["principal"], *p["secondaries"]]:
        fuente["path"] = ruta(Path(fuente["path"]).name)
    p["output"]["path"] = ruta("resultado.csv")
    receta.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")

    cargada = load_recipe(receta)
    assert sorted(cargada.missing) == ["clientes", "cobranzas", "ventas"]
    assert cargada.warnings  # avisa cuáles faltan
    assert [s.name for s in cargada.plan.secondaries] == ["ventas", "cobranzas"]
    assert cargada.plan.secondaries[0].summary == plan.secondaries[0].summary
    assert cargada.plan.output.split == plan.output.split
    # El usuario elige los archivos de este mes y la receta corre igual.
    nuevo = dataclasses.replace(
        cargada.plan,
        principal=dataclasses.replace(cargada.plan.principal, path=esc.rutas["clientes"]),
        secondaries=[
            dataclasses.replace(s, path=esc.rutas[s.name]) for s in cargada.plan.secondaries
        ],
    )
    assert run_join(nuevo, esc.tablas).df.equals(run_join(plan, esc.tablas).df)


def test_receta_rota_o_de_otra_version_da_un_error_amigable(tmp_path):
    for contenido in ("no es json", '{"version": 99, "plan": {}}', '{"plan": {}}', "[]"):
        ruta = tmp_path / "mala.json"
        ruta.write_text(contenido, encoding="utf-8")
        with pytest.raises(ExcelToolError) as e:
            load_recipe(ruta)
        assert "receta" in e.value.mensaje.lower()
    with pytest.raises(ExcelToolError):
        load_recipe(tmp_path / "no_existe.json")


# ---- 10. guardar con tipos reales


def test_guardar_en_xlsx_con_los_tipos_del_join_da_numeros_y_fechas_reales(esc, tmp_path):
    plan = JoinPlan(
        esc.principal(),
        SIN_CEROS,
        [esc.ventas(**resumen_ventas_por_cliente()), esc.cobranzas(mode="first")],
    )
    res = run_join(plan, esc.tablas)
    destino = tmp_path / "tipos.xlsx"
    save_result(res.df, OutputSpec(destino, "xlsx", "excel", column_types=res.column_types))
    ws = openpyxl.load_workbook(destino).worksheets[0]
    cols = [c.value for c in ws[1]]
    assert cols == list(res.df.columns)
    assert ws.max_row == 31 and ws.freeze_panes == "A2"
    por = esc.esp_ventas["por_cliente"]
    chequeadas = 0
    for fila in ws.iter_rows(min_row=2):
        celda = dict(zip(cols, fila, strict=True))
        cid = celda["ID cliente"].value
        assert isinstance(cid, int | float)  # entero real
        e = por.get(str(int(cid)))
        if e is None:
            assert celda["Total comprado"].value == 0  # ADR-037
            assert celda["Compras"].value == 0
            continue
        chequeadas += 1
        total = celda["Total comprado"].value
        assert isinstance(total, float) and Decimal(str(total)) == Decimal(e["suma_importe"])
        assert celda["Compras"].data_type == "n" and celda["Compras"].value == e["filas"]
        ultima = celda["Última compra"]
        assert ultima.data_type == "d" and ultima.value.strftime("%d/%m/%Y") == e["ultima_fecha"]
        assert isinstance(celda["Nombre"].value, str)
    assert chequeadas == 27


# ---- escenario grande


GRANDE_ESTIMADAS = {"detalle": 2_519_998, "resumen": 200_000}
TOPE_MB = 1536  # 1,5 GB (ADR-023)


@pytest.mark.parametrize("modo", ["detalle", "resumen"])
def test_escenario_grande_estimacion_igual_a_lo_real_y_memoria_dentro_de_1_5_gb(
    fixtures_grandes, fixtures_esperados, modo
):
    """200.000 x 2.500.000: en otro proceso, para medir el pico con `/usr/bin/time -v`."""
    exe = Path("/usr/bin/time")
    if not exe.is_file():
        pytest.skip("no está /usr/bin/time (solo Linux)")
    esperado = json.loads((fixtures_esperados / "grandes.json").read_text(encoding="utf-8"))
    assert esperado["detalle_left"]["total"] == GRANDE_ESTIMADAS["detalle"]
    script = Path(__file__).with_name("medicion_fase4.py")
    import os
    import sys

    entorno = {**os.environ, "PYTHONPATH": str(RAIZ / "src")}
    proc = subprocess.run(
        [str(exe), "-v", sys.executable, str(script), modo],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=entorno,
        timeout=600,
        check=False,
    )
    assert proc.returncode == 0, proc.stderr[-2000:]
    linea = next(x for x in proc.stdout.splitlines() if x.startswith("RESULTADO "))
    datos = json.loads(linea.removeprefix("RESULTADO "))
    assert datos["filas_estimadas"] == datos["filas_reales"] == GRANDE_ESTIMADAS[modo]
    assert datos["memoria_ok"] is True
    pico = re.search(r"Maximum resident set size \(kbytes\): (\d+)", proc.stderr)
    assert pico, proc.stderr[-1500:]
    mb = int(pico.group(1)) / 1024
    print(f"\n[{modo}] pico {mb:.0f} MB; {datos}")
    assert mb < TOPE_MB, f"pico de {mb:.0f} MB, más de 1,5 GB"
