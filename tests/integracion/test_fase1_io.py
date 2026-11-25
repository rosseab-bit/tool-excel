"""Integración de la Fase 1: solo API pública de excel_tool.core.io y excel_tool.core.types,
contra los archivos generados y los esperados (calculados aparte, sin usar el producto)."""

from __future__ import annotations

import json
import time
from decimal import Decimal
from pathlib import Path

import pandas as pd
import pytest

from excel_tool.core.errors import ArchivoIlegible
from excel_tool.core.io import (
    ReadOptions,
    WriteOptions,
    list_sheets,
    read_table,
    write_sheets,
    write_table,
)
from excel_tool.core.types import detect_types, to_bool, to_date, to_number

MOJIBAKE = ("Ã", "Â", "�")
FORMATOS_SALIDA = ["xlsx", "csv_excel", "csv_google", "csv_standard"]
EXT = {"xlsx": ".xlsx", "csv_excel": ".csv", "csv_google": ".csv", "csv_standard": ".csv"}


def _json(carpeta: Path, nombre: str) -> dict:
    return json.loads((carpeta / nombre).read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def fase1(fixtures_esperados):
    return _json(fixtures_esperados, "fase1.json")


@pytest.fixture(scope="module")
def dificiles(fixtures_esperados):
    return _json(fixtures_esperados, "dificiles.json")


def _todos(gen: Path) -> list[Path]:
    """Todos los archivos de prueba legibles (sin los grandes ni los que deben fallar)."""
    ext = {".csv", ".txt", ".xlsx", ".xls", ".xlsb"}
    excluidos = {"vacio.csv", "danado.xlsx"}
    return sorted(
        p
        for carpeta in ("formatos", "csv", "explorar", "escenario_3", "dificiles", "fase1")
        for p in (gen / carpeta).rglob("*")
        if p.suffix in ext and p.name not in excluidos
    )


def _ids(paths):
    return [p.name for p in paths]


# ---------------------------------------------------------------- lectura de todos los fixtures


def test_todos_los_fixtures_se_leen_sin_mojibake(fixtures_generados):
    archivos = _todos(fixtures_generados)
    assert len(archivos) >= 20
    for p in archivos:
        r = read_table(p)
        assert all(str(c) for c in r.df.columns), p.name
        for col in r.df.columns:
            assert not any(m in str(col) for m in MOJIBAKE), (p.name, col)
        for v in r.df.to_numpy().ravel():
            assert isinstance(v, str), (p.name, v)
            assert not any(m in v for m in MOJIBAKE), (p.name, v)


@pytest.mark.parametrize(
    "ruta",
    [
        "formatos/chico.xls",
        "formatos/chico.xlsb",
        "formatos/varias_hojas.xlsx",
        "csv/coma_utf8.csv",
        "csv/excel_ar.csv",
        "csv/tabulado.txt",
        "csv/con_bom.csv",
    ],
)
def test_tildes_enie_y_ceros_intactos(fixtures_generados, ruta):
    df = read_table(fixtures_generados / ruta).df
    assert list(df.columns)[:3] == ["Código", "Nombre", "Localidad"]
    assert df.shape == (12, 4)
    assert df.iloc[0]["Código"] == "00123"
    assert df.iloc[0]["Nombre"] == "Ñandú Pérez"
    assert df.iloc[0]["Localidad"] == "Córdoba"
    assert "José Muñoz" in set(df["Nombre"])
    assert "María Núñez" in set(df["Nombre"])
    assert all(len(c) == 5 for c in df["Código"])


@pytest.mark.parametrize("nombre", ["con_basura.csv", "con_basura.xlsx"])
def test_encabezado_con_basura(fixtures_generados, dificiles, nombre):
    esp = dificiles[nombre]
    r = read_table(fixtures_generados / "dificiles" / nombre)
    assert r.header_row == esp["fila_encabezado"] - 1
    assert list(r.df.columns) == esp["columnas"]
    assert len(r.df) == esp["filas_de_datos"]
    assert r.df.iloc[0]["Código"] == "00123"
    assert r.df.iloc[0]["Nombre"] == "Ñandú Pérez"


@pytest.mark.parametrize("nombre", ["claves_sucias.csv", "claves_sucias.xlsx"])
def test_claves_sucias_se_leen_tal_cual(fixtures_generados, dificiles, nombre):
    esp = dificiles["claves_sucias"]
    df = read_table(fixtures_generados / "dificiles" / nombre).df
    assert len(df) == esp["filas_de_datos"]
    assert df[esp["columna_clave"]].nunique() == esp["claves_distintas_sin_normalizar"]
    assert {"  Peña ", "PEÑA", "00123", "123", "0123", "Núñez "} <= set(df["Clave"])


def test_claves_duplicadas(fixtures_generados, dificiles):
    esp = dificiles["principal_claves_duplicadas.csv"]
    df = read_table(fixtures_generados / "dificiles" / "principal_claves_duplicadas.csv").df
    assert len(df) == esp["filas_de_datos"]
    assert df["Clave"].value_counts()[lambda s: s > 1].to_dict() == esp["claves_duplicadas"]


@pytest.mark.parametrize("nombre", ["excel_tipos.xlsx", "excel_tipos.xls"])
def test_excel_tipos_valores(fixtures_generados, fase1, nombre):
    esp = fase1[nombre]
    r = read_table(fixtures_generados / "fase1" / nombre)
    assert r.sheet == esp["hoja_principal"]
    assert list(r.df.columns) == esp["columnas"]
    assert r.df.values.tolist() == esp["filas"]


def test_excel_tipos_hoja_con_vacias(fixtures_generados, fase1):
    esp = fase1["excel_tipos.xlsx"]["hoja_con_vacias"]
    r = read_table(
        fixtures_generados / "fase1" / "excel_tipos.xlsx", ReadOptions(sheet="Con vacias")
    )
    assert list(r.df.columns) == esp["columnas"]
    assert r.df.values.tolist() == esp["filas"]


@pytest.mark.parametrize(
    "ruta", ["csv/pipe.csv", "csv/estilo_punto.csv", "csv/solo_encabezado.csv"]
)
def test_csv_fase1(fixtures_generados, fase1, ruta):
    esp = fase1[ruta]
    r = read_table(fixtures_generados / "fase1" / ruta)
    assert r.csv_options.separator == esp["separador"]
    assert r.csv_options.encoding == esp["encoding"]
    assert list(r.df.columns) == esp["columnas"]
    assert r.df.values.tolist() == esp["filas"]


@pytest.mark.parametrize("ruta", ["csv/vacio.csv", "danado.xlsx"])
def test_errores_amigables(fixtures_generados, ruta):
    with pytest.raises(ArchivoIlegible) as e:
        read_table(fixtures_generados / "fase1" / ruta)
    assert (
        e.value.mensaje
        and not e.value.mensaje.isascii()
        or e.value.mensaje.startswith("El archivo")
    )


def test_tipos_esperados_de_fase1(fixtures_generados, fase1):
    for nombre, esp in fase1.items():
        if nombre.startswith("_") or "tipos" not in esp:
            continue
        ruta = fixtures_generados / "fase1" / nombre
        det = detect_types(read_table(ruta).df)
        assert {c: d.detected for c, d in det.items()} == esp["tipos"], nombre


# ---------------------------------------------------------------- list_sheets


def test_list_sheets_varias_hojas(fixtures_generados):
    hojas = list_sheets(fixtures_generados / "formatos" / "varias_hojas.xlsx")
    assert [(h.name, h.rows, h.columns) for h in hojas] == [
        ("Clientes", 12, 4),
        ("Ventas", 25, 6),
        ("Resumen año", 2, 2),
    ]


def test_list_sheets_fase1(fixtures_generados, fase1):
    esp = fase1["excel_tipos.xlsx"]["hojas"]
    hojas = list_sheets(fixtures_generados / "fase1" / "excel_tipos.xlsx")
    assert {h.name: {"filas": h.rows, "columnas": h.columns} for h in hojas} == {
        k: {"filas": v["filas"], "columnas": v["columnas"]} for k, v in esp.items()
    }


def test_list_sheets_csv(fixtures_generados):
    hojas = list_sheets(fixtures_generados / "csv" / "excel_ar.csv")
    assert [(h.name, h.rows, h.columns) for h in hojas] == [("excel_ar.csv", 12, 4)]


# ---------------------------------------------------------------- ida y vuelta


@pytest.mark.parametrize("formato", FORMATOS_SALIDA)
@pytest.mark.parametrize(
    "ruta",
    [
        "formatos/chico.xls",
        "csv/excel_ar.csv",
        "dificiles/claves_sucias.csv",
        "dificiles/con_basura.xlsx",
        "escenario_3/ventas.csv",
        "explorar/para_explorar.csv",
        "fase1/excel_tipos.xlsx",
        "fase1/csv/estilo_punto.csv",
    ],
)
def test_ida_y_vuelta(fixtures_generados, tmp_path, ruta, formato):
    original = read_table(fixtures_generados / ruta).df
    destino = tmp_path / ("salida" + EXT[formato])
    res = write_table(original, destino, WriteOptions(format=formato))
    assert res.rows == len(original) and res.unconverted == {}
    vuelto = read_table(destino).df
    pd.testing.assert_frame_equal(vuelto, original, check_dtype=False)


def test_write_sheets_ida_y_vuelta(fixtures_generados, tmp_path):
    a = read_table(fixtures_generados / "csv" / "excel_ar.csv").df
    b = read_table(fixtures_generados / "escenario_3" / "ventas.csv").df
    destino = tmp_path / "dos_hojas.xlsx"
    resultados = write_sheets({"Clientes": a, "Ventas": b}, destino)
    assert resultados
    assert [(h.name, h.rows, h.columns) for h in list_sheets(destino)] == [
        ("Clientes", len(a), a.shape[1]),
        ("Ventas", len(b), b.shape[1]),
    ]
    pd.testing.assert_frame_equal(
        read_table(destino, ReadOptions(sheet="Ventas")).df, b, check_dtype=False
    )


def test_escritura_con_tipos_informa_no_convertidas(fixtures_generados, tmp_path):
    df = read_table(fixtures_generados / "escenario_3" / "ventas.csv").df
    res = write_table(
        df, tmp_path / "t.xlsx", WriteOptions(format="xlsx", column_types={"Importe": "decimal"})
    )
    assert res.rows == len(df)
    assert sorted(v for _, v in res.unconverted["Importe"]) == sorted(["n/a", "12,34,5", "abc"])


def test_xlsx_con_tipos_se_relee_como_texto_numerico(tmp_path):
    df = pd.DataFrame({"N": ["100", "1.5"], "F": ["05/03/2024", "31/12/2023"]})
    write_table(
        df,
        tmp_path / "t.xlsx",
        WriteOptions(format="xlsx", column_types={"N": "decimal", "F": "fecha"}),
    )
    vuelto = read_table(tmp_path / "t.xlsx").df
    assert vuelto["N"].tolist() == ["100", "1.5"]
    assert vuelto["F"].tolist() == ["05/03/2024", "31/12/2023"]


# ---------------------------------------------------------------- tipos vs perfil.json


@pytest.mark.parametrize("nombre", ["para_explorar.csv", "para_explorar.xlsx"])
def test_tipos_coinciden_con_perfil(fixtures_generados, fixtures_esperados, nombre):
    perfil = _json(fixtures_esperados, "perfil.json")
    r = read_table(fixtures_generados / "explorar" / nombre)
    assert len(r.df) == perfil["filas"] and r.df.shape[1] == perfil["columnas"]
    det = detect_types(r.df)
    for esp, real in zip(perfil["columnas_detalle"], r.df.columns, strict=True):
        d = det[real]
        assert d.detected == esp["tipo"], (nombre, real)
        if esp["desglose_tipos"]:
            assert d.breakdown == esp["desglose_tipos"], (nombre, real)
        assert d.empty == esp["vacios"], (nombre, real)


# ---------------------------------------------------------------- escenario 3


def test_escenario_3_ventas(fixtures_generados, fixtures_esperados):
    esp = _json(fixtures_esperados / "escenario_3", "resumen_ventas.json")
    df = read_table(fixtures_generados / "escenario_3" / "ventas.csv").df
    assert len(df) == esp["filas_total"]
    conv = to_number(df["Importe"])
    # el contrato informa solo las celdas no vacías; las vacías se agregan para comparar
    vacias = [(i, "") for i, v in enumerate(df["Importe"]) if v.strip() == ""]
    invalidas = sorted(conv.failed + vacias)
    esperadas = sorted((c["fila"] - 1, c["valor"]) for c in esp["celdas_invalidas"])
    assert invalidas == esperadas

    por_cliente: dict[str, Decimal] = {}
    for cliente, valor in zip(df["ID cliente"], conv.values, strict=True):
        if pd.notna(valor):
            por_cliente[cliente] = por_cliente.get(cliente, Decimal(0)) + Decimal(
                str(round(valor, 2))
            )
    esperado = {c: Decimal(d["suma_importe"]) for c, d in esp["por_cliente"].items()}
    # clientes con solo importes inválidos no figuran en `por_cliente` si suman 0
    assert {c: s for c, s in por_cliente.items() if s != 0} == {
        c: s for c, s in esperado.items() if s != 0
    }
    assert set(df["ID cliente"]) == set(esp["por_cliente"])
    assert set(esp["ids_fuera_de_clientes"]) <= set(df["ID cliente"])


def test_escenario_3_conversiones_de_fechas(fixtures_generados, fixtures_esperados):
    esp = _json(fixtures_esperados / "escenario_3", "resumen_ventas.json")
    df = read_table(fixtures_generados / "escenario_3" / "ventas.csv").df
    fechas = to_date(df["Fecha"])
    assert fechas.failed == []
    for cliente, d in esp["por_cliente"].items():
        ultima = fechas.values[df["ID cliente"] == cliente].max()
        assert ultima.strftime("%d/%m/%Y") == d["ultima_fecha"], cliente


def test_conversiones_reportan_fallidas():
    s = pd.Series(["Sí", "no", "VERDADERO", "talvez", "", "1"])
    r = to_bool(s)
    assert r.failed == [(3, "talvez"), (5, "1")]
    f = to_date(pd.Series(["31/02/2024", "05/03/2024", "x", ""]))
    assert f.failed == [(0, "31/02/2024"), (2, "x")]


# ---------------------------------------------------------------- rendimiento


def _medir(etiqueta: str, limite: float, f):
    t0 = time.perf_counter()
    res = f()
    seg = time.perf_counter() - t0
    print(f"\n[RENDIMIENTO] {etiqueta}: {seg:.1f} s (objetivo {limite} s, falla > {2 * limite} s)")
    assert seg <= 2 * limite, f"{etiqueta} tardó {seg:.1f} s"
    return res


def test_rendimiento_xlsx_500k(fixtures_grandes):
    ruta = fixtures_grandes / "rendimiento_500k.xlsx"
    r = _medir("read_table 500k xlsx", 30, lambda: read_table(ruta))
    assert len(r.df) == 500_000
    hojas = _medir("list_sheets 500k xlsx", 30, lambda: list_sheets(ruta))
    assert hojas[0].rows == 500_000 and hojas[0].columns == r.df.shape[1]


def test_rendimiento_csv_2M(fixtures_grandes):
    ruta = fixtures_grandes / "rendimiento_2M.csv"
    r = _medir("read_table 2M csv", 20, lambda: read_table(ruta))
    assert len(r.df) == 2_000_000
    det = _medir("detect_types 2M csv", 15, lambda: detect_types(r.df))
    assert set(det) == set(r.df.columns)
    hojas = list_sheets(ruta)
    assert hojas[0].rows == 2_000_000
