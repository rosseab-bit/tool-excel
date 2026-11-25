"""Lectura de Excel con calamine: conversión a texto, hojas, encabezado y errores."""

from __future__ import annotations

import datetime as dt
import json
import shutil

import pandas as pd
import pytest

from excel_tool.core.errors import ArchivoBloqueado, ArchivoIlegible
from excel_tool.core.io import ReadOptions, list_sheets, read_table
from excel_tool.core.io.excel import _celda_a_texto


@pytest.fixture(scope="module")
def esp(fixtures_esperados):
    return json.loads((fixtures_esperados / "fase1.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize(
    ("valor", "texto"),
    [
        (100.0, "100"),
        (100, "100"),
        (-5.0, "-5"),
        (1234.56, "1234.56"),
        (0.001, "0.001"),
        (1e-7, "0.0000001"),
        (1.5e20, "150000000000000000000"),
        (True, "VERDADERO"),
        (False, "FALSO"),
        (dt.date(2024, 3, 5), "05/03/2024"),
        (dt.datetime(2024, 3, 5), "05/03/2024"),
        (dt.datetime(2024, 3, 5, 14, 30, 15), "05/03/2024 14:30:15"),
        (dt.time(8, 5, 9), "08:05:09"),
        (dt.timedelta(hours=26, minutes=1), "26:01:00"),
        ("", ""),
        ("00012", "00012"),
        (float("nan"), ""),
    ],
)
def test_celda_a_texto(valor, texto):
    assert _celda_a_texto(valor) == texto


@pytest.mark.parametrize("nombre", ["excel_tipos.xlsx", "excel_tipos.xls"])
def test_valores_esperados(fixtures_generados, esp, nombre):
    e = esp[nombre]
    r = read_table(fixtures_generados / "fase1" / nombre)
    assert list(r.df.columns) == e["columnas"]
    assert r.df.values.tolist() == e["filas"]
    assert r.sheet == e["hoja_principal"]
    assert r.header_row == 0
    assert all(isinstance(v, str) for v in r.df.to_numpy().ravel())


def test_hoja_con_filas_vacias_al_final(fixtures_generados, esp):
    e = esp["excel_tipos.xlsx"]
    r = read_table(
        fixtures_generados / "fase1" / "excel_tipos.xlsx", ReadOptions(sheet="Con vacias")
    )
    assert list(r.df.columns) == e["hoja_con_vacias"]["columnas"]
    assert r.df.values.tolist() == e["hoja_con_vacias"]["filas"]
    assert r.sheets == ["Tipos", "Con vacias"]


@pytest.mark.parametrize("nombre", ["varias_hojas.xlsx", "chico.xls", "chico.xlsb"])
def test_formatos(fixtures_generados, nombre):
    r = read_table(fixtures_generados / "formatos" / nombre)
    assert r.sheet == "Clientes"
    assert r.format == nombre.rsplit(".", 1)[1]
    assert list(r.df.columns) == ["Código", "Nombre", "Localidad", "Saldo"]
    assert r.df.iloc[0].tolist() == ["00123", "Ñandú Pérez", "Córdoba", "1234.56"]
    assert len(r.df) == 12
    # Textos compactos (ADR-025): dtype str con almacenamiento pyarrow y vacíos como "".
    assert all(isinstance(t, pd.StringDtype) and t.storage == "pyarrow" for t in r.df.dtypes)
    assert not r.df.isna().any().any()


def test_hoja_por_nombre_y_inexistente(fixtures_generados):
    ruta = fixtures_generados / "formatos" / "varias_hojas.xlsx"
    r = read_table(ruta, ReadOptions(sheet="Resumen año"))
    assert list(r.df.columns) == ["Concepto", "Valor"]
    assert r.df.values.tolist() == [["Año", "2024"], ["Observación", "Resumen anual ñandú"]]
    with pytest.raises(ArchivoIlegible) as exc:
        read_table(ruta, ReadOptions(sheet="Nada"))
    assert "Clientes" in exc.value.mensaje and "Ventas" in exc.value.mensaje


def test_nrows(fixtures_generados):
    r = read_table(fixtures_generados / "explorar" / "para_explorar.xlsx", ReadOptions(nrows=10))
    assert len(r.df) == 10


def test_encabezado_con_basura(fixtures_generados, fixtures_esperados):
    e = json.loads((fixtures_esperados / "dificiles.json").read_text(encoding="utf-8"))
    e = e["con_basura.xlsx"]
    ruta = fixtures_generados / "dificiles" / "con_basura.xlsx"
    r = read_table(ruta)
    assert r.header_row == e["fila_encabezado"] - 1
    assert list(r.df.columns) == e["columnas"]
    assert len(r.df) == e["filas_de_datos"]
    r2 = read_table(ruta, ReadOptions(header_row=0))
    assert r2.header_row == 0
    assert len(r2.df) == len(read_table(ruta).df) + 3


def test_para_explorar_igual_al_csv(fixtures_generados):
    x = read_table(fixtures_generados / "explorar" / "para_explorar.xlsx")
    c = read_table(fixtures_generados / "explorar" / "para_explorar.csv")
    assert list(x.df.columns) == list(c.df.columns)
    assert x.renamed_columns == c.renamed_columns
    assert len(x.df) == len(c.df)
    # Los textos y las claves con ceros tienen que ser idénticos.
    for col in x.df.columns:
        if col in ("Código", "Observaciones"):
            assert x.df[col].tolist() == c.df[col].tolist()
    assert x.df.iloc[:, 6].tolist() == c.df.iloc[:, 6].tolist()


def test_ceros_a_la_izquierda(fixtures_generados):
    r = read_table(fixtures_generados / "escenario_3" / "cobranzas.xlsx")
    assert "00018" in r.df.iloc[:, 0].tolist()


def test_list_sheets(fixtures_generados):
    s = list_sheets(fixtures_generados / "formatos" / "varias_hojas.xlsx")
    assert [(h.name, h.rows, h.columns) for h in s] == [
        ("Clientes", 12, 4),
        ("Ventas", 25, 6),
        ("Resumen año", 2, 2),
    ]


def test_list_sheets_con_vacias_y_basura(fixtures_generados, esp):
    h = esp["excel_tipos.xlsx"]["hojas"]
    s = {x.name: x for x in list_sheets(fixtures_generados / "fase1" / "excel_tipos.xlsx")}
    assert (s["Tipos"].rows, s["Tipos"].columns) == (h["Tipos"]["filas"], h["Tipos"]["columnas"])
    assert s["Con vacias"].rows == h["Con vacias"]["filas"]
    b = list_sheets(fixtures_generados / "dificiles" / "con_basura.xlsx")
    assert b[0].rows == 6


def test_danado(fixtures_generados):
    with pytest.raises(ArchivoIlegible):
        read_table(fixtures_generados / "fase1" / "danado.xlsx")
    with pytest.raises(ArchivoIlegible):
        list_sheets(fixtures_generados / "fase1" / "danado.xlsx")


def test_bloqueado(fixtures_generados, tmp_path, monkeypatch):
    import excel_tool.core.io.excel as modulo

    def falla(_ruta):
        raise PermissionError("bloqueado")

    monkeypatch.setattr(modulo.CalamineWorkbook, "from_path", staticmethod(falla))
    destino = tmp_path / "a.xlsx"
    shutil.copy(fixtures_generados / "formatos" / "varias_hojas.xlsx", destino)
    with pytest.raises(ArchivoBloqueado):
        read_table(destino)
