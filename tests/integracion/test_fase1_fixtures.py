"""Consistencia de los archivos de la Fase 1 con sus esperados, leídos con pandas/calamine
directo (sin usar el producto)."""

from __future__ import annotations

import datetime as dt
import json
import zipfile

import pandas as pd
import pytest
from python_calamine import CalamineWorkbook


@pytest.fixture(scope="module")
def esp(fixtures_esperados):
    return json.loads((fixtures_esperados / "fase1.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def carpeta(fixtures_generados):
    return fixtures_generados / "fase1"


def _texto(v) -> str:
    """Formatea un valor de calamine según las reglas del contrato (independiente del producto)."""
    if isinstance(v, bool):
        return "VERDADERO" if v else "FALSO"
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else repr(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, dt.datetime):
        return v.strftime("%d/%m/%Y %H:%M:%S")
    if isinstance(v, dt.date):
        return v.strftime("%d/%m/%Y")
    return str(v)


def _hoja(ruta, nombre=None):
    wb = CalamineWorkbook.from_path(str(ruta))
    filas = wb.get_sheet_by_name(nombre or wb.sheet_names[0]).to_python()
    while filas and all(c == "" for c in filas[-1]):
        filas.pop()
    return [[_texto(c) for c in f] for f in filas]


@pytest.mark.parametrize("nombre", ["excel_tipos.xlsx", "excel_tipos.xls"])
def test_excel_valores_esperados(carpeta, esp, nombre):
    e = esp[nombre]
    filas = _hoja(carpeta / nombre, e["hoja_principal"])
    assert filas[0] == e["columnas"]
    assert filas[1:] == e["filas"]


def test_excel_tipos_enteros_son_numericos_en_el_archivo(carpeta):
    crudo = CalamineWorkbook.from_path(str(carpeta / "excel_tipos.xlsx"))
    filas = crudo.get_sheet_by_name("Tipos").to_python()
    assert filas[1][0] == 100 and filas[4][0] == 1048577
    assert isinstance(filas[2][4], bool)
    assert isinstance(filas[1][2], dt.date) and isinstance(filas[1][3], dt.datetime)


def test_xlsx_formula_con_valor_guardado_y_combinadas(carpeta, esp):
    with zipfile.ZipFile(carpeta / "excel_tipos.xlsx") as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert "<f>A2+1</f><v>101</v>" in xml
    assert "<f>A3+1</f><v>1</v>" in xml
    assert 't="e"><f>NA()</f><v>#N/A</v>' in xml
    for rango in esp["excel_tipos.xlsx"]["celdas_combinadas"]:
        assert f'<mergeCell ref="{rango}"' in xml


def test_xlsx_hoja_con_filas_vacias_al_final(carpeta, esp):
    e = esp["excel_tipos.xlsx"]
    with zipfile.ZipFile(carpeta / "excel_tipos.xlsx") as z:
        xml = z.read("xl/worksheets/sheet2.xml").decode("utf-8")
    assert '<dimension ref="A1:B6"' in xml  # el archivo declara 6 filas, 3 sin datos
    h = e["hojas"]["Con vacias"]
    filas = _hoja(carpeta / "excel_tipos.xlsx", "Con vacias")
    assert len(filas) - 1 == h["filas"]
    assert filas[0] == e["hoja_con_vacias"]["columnas"]
    assert filas[1:] == e["hoja_con_vacias"]["filas"]
    assert len(_hoja(carpeta / "excel_tipos.xlsx", "Tipos")) - 1 == e["hojas"]["Tipos"]["filas"]


@pytest.mark.parametrize(
    "nombre", ["csv/pipe.csv", "csv/estilo_punto.csv", "csv/solo_encabezado.csv"]
)
def test_csv_esperados(carpeta, esp, nombre):
    e = esp[nombre]
    df = pd.read_csv(carpeta / nombre, sep=e["separador"], encoding=e["encoding"], dtype=str)
    assert list(df.columns) == e["columnas"]
    assert df.fillna("").values.tolist() == e["filas"]
    assert list(e["tipos"]) == e["columnas"]


def test_csv_encoding_y_pipe(carpeta):
    datos = (carpeta / "csv" / "pipe.csv").read_bytes()
    assert not datos.startswith(b"\xef\xbb\xbf")
    assert "Ñandú".encode() in datos and b"|" in datos and b";" not in datos


def test_csv_estilo_punto_tiene_miles_con_coma(carpeta):
    texto = (carpeta / "csv" / "estilo_punto.csv").read_text(encoding="utf-8")
    assert '"1,234.56"' in texto


def test_vacio_y_danado(carpeta, esp):
    assert (carpeta / "csv" / "vacio.csv").stat().st_size == esp["csv/vacio.csv"]["bytes"] == 0
    assert esp["csv/vacio.csv"]["error"] == "ArchivoIlegible"
    assert esp["danado.xlsx"]["error"] == "ArchivoIlegible"
    ruta = carpeta / "danado.xlsx"
    assert ruta.stat().st_size > 0 and not zipfile.is_zipfile(ruta)
    with pytest.raises(Exception):  # noqa: B017 - cualquier error de lectura de calamine
        CalamineWorkbook.from_path(str(ruta))


def test_generacion_deterministica(generador_fixtures, tmp_path):
    a, b = tmp_path / "a", tmp_path / "b"
    generador_fixtures.grupo_fase1(a)
    generador_fixtures.grupo_fase1(b)
    for rel in ("excel_tipos.xlsx", "excel_tipos.xls", "csv/pipe.csv", "danado.xlsx"):
        assert (a / "fase1" / rel).read_bytes() == (b / "fase1" / rel).read_bytes(), rel
