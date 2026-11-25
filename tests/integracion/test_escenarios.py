"""Los archivos del escenario de 3 archivos y los casos difíciles son consistentes con sus
esperados. Se recalcula con pandas directo (dtype=str), sin código del producto."""

from __future__ import annotations

import json
import re
from decimal import Decimal

import pandas as pd
import pytest

IMPORTE_VALIDO = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$")


def _json(esperados, *partes):
    ruta = esperados.joinpath(*partes)
    return json.loads(ruta.read_text(encoding="utf-8"))


def _excel(ruta, hoja=0):
    return pd.read_excel(ruta, sheet_name=hoja, dtype=str, engine="calamine")


@pytest.fixture(scope="module")
def e3(fixtures_generados):
    base = fixtures_generados / "escenario_3"
    return {
        "clientes": _excel(base / "clientes.xlsx"),
        "ventas": pd.read_csv(
            base / "ventas.csv", sep=";", encoding="cp1252", dtype=str, keep_default_na=False
        ),
        "cobranzas": _excel(base / "cobranzas.xlsx"),
    }


def _a_decimal(texto: str) -> Decimal:
    return Decimal(texto.replace(".", "").replace(",", "."))


def test_estructura_del_escenario(e3):
    c, v, cb = e3["clientes"], e3["ventas"], e3["cobranzas"]
    assert list(c.columns) == ["ID cliente", "Nombre", "Región", "Vendedor"]
    assert list(v.columns) == ["ID cliente", "Fecha", "Artículo", "Cantidad", "Importe"]
    assert list(cb.columns) == ["Nro Cliente", "Fecha", "Medio de pago", "Monto"]
    assert len(c) == 30 and c["ID cliente"].is_unique
    assert 4 <= c["Región"].nunique() <= 5
    assert c["Nombre"].str.contains("[ñÑáéíóúÁ]", regex=True).any()
    assert (v["Fecha"].str.fullmatch(r"\d{2}/\d{2}/\d{4}")).all()
    # claves de cobranzas: con ceros y con espacios
    assert cb["Nro Cliente"].str.fullmatch(r"0\d{4}").any()
    assert cb["Nro Cliente"].str.fullmatch(r" \d+ ").any()


def test_ventas_casos_pedidos(e3):
    c, v = e3["clientes"], e3["ventas"]
    ids = set(c["ID cliente"])
    por_cliente = v["ID cliente"].value_counts()
    assert por_cliente.max() <= 8
    assert len(ids - set(v["ID cliente"])) >= 1  # clientes sin ventas
    assert len(set(v["ID cliente"]) - ids) >= 1  # ventas de IDs que no están
    invalidos = v[~v["Importe"].str.fullmatch(IMPORTE_VALIDO)]
    assert 3 <= len(invalidos) <= 5
    assert "" in set(invalidos["Importe"])
    assert "12,34,5" in set(invalidos["Importe"])


def test_resumen_ventas_coincide(e3, fixtures_esperados):
    esp = _json(fixtures_esperados, "escenario_3", "resumen_ventas.json")
    assert "_nota" in esp
    v = e3["ventas"]
    valido = v["Importe"].str.fullmatch(IMPORTE_VALIDO)
    assert set(esp["por_cliente"]) == set(v["ID cliente"])
    for cid, grupo in v.groupby("ID cliente"):
        r = esp["por_cliente"][cid]
        buenos = grupo[valido.loc[grupo.index]]
        assert r["filas"] == len(grupo)
        assert Decimal(r["suma_importe"]) == sum(
            (_a_decimal(x) for x in buenos["Importe"]), Decimal(0)
        )
        assert r["importes_invalidos"] == len(grupo) - len(buenos)
        assert r["articulos_distintos"] == grupo["Artículo"].nunique()
        ultima = pd.to_datetime(grupo["Fecha"], format="%d/%m/%Y").max()
        assert r["ultima_fecha"] == ultima.strftime("%d/%m/%Y")
    esperadas = [(c["fila_en_archivo"], c["valor"]) for c in esp["celdas_invalidas"]]
    reales = [(i + 2, x) for i, x in enumerate(v["Importe"]) if not IMPORTE_VALIDO.fullmatch(x)]
    assert esperadas == reales
    assert esp["filas_total"] == len(v)


def test_resumen_cobranzas_coincide(e3, fixtures_esperados):
    esp = _json(fixtures_esperados, "escenario_3", "resumen_cobranzas.json")
    cb = e3["cobranzas"].copy()
    cb["id"] = cb["Nro Cliente"].str.strip().str.lstrip("0")
    assert set(esp["por_cliente"]) == set(cb["id"])
    for cid, grupo in cb.groupby("id"):
        assert esp["por_cliente"][cid]["filas"] == len(grupo)
        suma = sum((Decimal(x) for x in grupo["Monto"]), Decimal(0))
        assert Decimal(esp["por_cliente"][cid]["suma_monto"]) == suma


def _diag(ids, claves):
    cuenta = pd.Series(claves).value_counts()
    return {
        "con": sum(1 for k in ids if k in cuenta.index),
        "sin": [k for k in ids if k not in cuenta.index],
        "huerfanas": set(cuenta.index) - set(ids),
        "filas": sum(max(1, int(cuenta.get(k, 0))) for k in ids),
        "cuenta": cuenta,
    }


def test_diagnostico_coincide(e3, fixtures_esperados):
    esp = _json(fixtures_esperados, "escenario_3", "diagnostico.json")
    ids = list(e3["clientes"]["ID cliente"])
    dv = _diag(ids, list(e3["ventas"]["ID cliente"]))
    ev = esp["ventas"]
    assert ev["claves_con_coincidencia"] == dv["con"]
    assert ev["claves_sin_coincidencia"] == dv["sin"]
    assert set(ev["claves_del_secundario_sin_pareja"]) == dv["huerfanas"]
    assert ev["claves_repetidas_en_principal"] == []
    assert ev["filas_resultado_detalle"] == dv["filas"]

    brutas = list(e3["cobranzas"]["Nro Cliente"])
    variantes = {
        "solo_recortar_espacios": [k.strip() for k in brutas],
        "sin_ceros_a_la_izquierda": [k.strip().lstrip("0") for k in brutas],
    }
    for nombre, claves in variantes.items():
        dc = _diag(ids, claves)
        ec = esp["cobranzas"][nombre]
        assert ec["claves_con_coincidencia"] == dc["con"]
        assert ec["claves_sin_coincidencia"] == dc["sin"]
        assert set(ec["claves_del_secundario_sin_pareja"]) == dc["huerfanas"]
        assert ec["filas_resultado_detalle"] == dc["filas"]
        multiplicado = sum(
            max(1, int(dv["cuenta"].get(k, 0))) * max(1, int(dc["cuenta"].get(k, 0))) for k in ids
        )
        assert ec["filas_resultado_ventas_detalle_y_cobranzas_detalle"] == multiplicado
        assert ec["filas_resultado_ventas_detalle_y_cobranzas_resumen"] == dv["filas"]
        assert ec["filas_resultado_ventas_resumen_y_cobranzas_resumen"] == len(ids) == 30
    # con solo recortar espacios casi nada coincide; sin ceros coincide
    assert (
        esp["cobranzas"]["solo_recortar_espacios"]["claves_con_coincidencia"]
        < esp["cobranzas"]["sin_ceros_a_la_izquierda"]["claves_con_coincidencia"]
    )
    # la multiplicación existe: más filas que las de ventas solas
    assert (
        esp["cobranzas"]["sin_ceros_a_la_izquierda"][
            "filas_resultado_ventas_detalle_y_cobranzas_detalle"
        ]
        > ev["filas_resultado_detalle"]
    )


def test_resumen_final_por_region(e3, fixtures_esperados):
    esp = _json(fixtures_esperados, "escenario_3", "resumen_final_por_region.json")
    c, v = e3["clientes"], e3["ventas"]
    v = v[v["Importe"].str.fullmatch(IMPORTE_VALIDO)].copy()
    v["n"] = v["Importe"].map(_a_decimal)
    por_cliente = v.groupby("ID cliente")["n"].sum()
    c = c.assign(total=c["ID cliente"].map(por_cliente).fillna(Decimal(0)))
    por_region = c.groupby("Región")["total"].sum()
    assert set(esp["total_importe_por_region"]) == set(por_region.index)
    for region, total in por_region.items():
        assert Decimal(esp["total_importe_por_region"][region]) == total
    fuera = v[~v["ID cliente"].isin(c["ID cliente"])]["n"].sum()
    assert Decimal(esp["importe_excluido_ids_fuera_de_clientes"]) == Decimal(fuera)


# --- Casos difíciles


def test_dificiles_claves_duplicadas(fixtures_generados, fixtures_esperados):
    esp = _json(fixtures_esperados, "dificiles.json")
    df = pd.read_csv(
        fixtures_generados / "dificiles/principal_claves_duplicadas.csv",
        sep=";",
        encoding="utf-8-sig",
        dtype=str,
    )
    e = esp["principal_claves_duplicadas.csv"]
    assert len(df) == e["filas_de_datos"]
    cuenta = df["Clave"].value_counts()
    assert dict(cuenta[cuenta > 1]) == e["claves_duplicadas"]
    assert df["Clave"].nunique() == e["claves_distintas"]


def test_dificiles_claves_sucias(fixtures_generados, fixtures_esperados):
    e = _json(fixtures_esperados, "dificiles.json")["claves_sucias"]
    csv = pd.read_csv(
        fixtures_generados / "dificiles/claves_sucias.csv",
        sep=";",
        encoding="utf-8-sig",
        dtype=str,
    )
    xlsx = _excel(fixtures_generados / "dificiles/claves_sucias.xlsx")
    pd.testing.assert_frame_equal(csv, xlsx)
    assert len(csv) == e["filas_de_datos"]
    assert csv["Clave"].nunique() == e["claves_distintas_sin_normalizar"]
    grupos = csv.groupby(csv["Clave"].str.strip().str.casefold())["Clave"].apply(list)
    assert grupos.to_dict() == e["grupos_por_clave_normalizada"]
    sin_ceros = csv.groupby(csv["Clave"].str.strip().str.casefold().str.lstrip("0"))["Clave"].apply(
        list
    )
    assert sin_ceros.to_dict() == e["grupos_sin_ceros_a_la_izquierda"]
    assert {"00123", "Ñandú", "  Peña "} <= set(csv["Clave"])


@pytest.mark.parametrize("archivo", ["con_basura.csv", "con_basura.xlsx"])
def test_dificiles_con_basura(fixtures_generados, fixtures_esperados, archivo):
    e = _json(fixtures_esperados, "dificiles.json")[archivo]
    ruta = fixtures_generados / "dificiles" / archivo
    if archivo.endswith(".csv"):
        crudo = pd.read_csv(
            ruta, sep=";", encoding="utf-8-sig", dtype=str, header=None, keep_default_na=False
        )
    else:
        crudo = pd.read_excel(
            ruta, dtype=str, engine="calamine", header=None, keep_default_na=False
        ).fillna("")
    assert e["fila_encabezado"] == 4
    assert crudo.iloc[0, 0].startswith("Listado")
    assert crudo.iloc[1, 0].startswith("Emitido")
    assert (crudo.iloc[2] == "").all()
    assert list(crudo.iloc[e["fila_encabezado"] - 1]) == e["columnas"]
    datos = crudo.iloc[e["fila_encabezado"] :]
    assert len(datos) == e["filas_de_datos"]
    assert "00123" in set(datos[0]) and "Ñandú Pérez" in set(datos[1])
