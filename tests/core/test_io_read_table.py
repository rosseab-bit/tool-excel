"""read_table y list_sheets: despacho por extensión y errores."""

from __future__ import annotations

import shutil

import pytest

from excel_tool.core.errors import ArchivoNoEncontrado, FormatoNoSoportado
from excel_tool.core.io import list_sheets, read_table


def test_csv_y_txt(fixtures_generados):
    assert read_table(fixtures_generados / "csv" / "excel_ar.csv").format == "csv"
    assert read_table(fixtures_generados / "csv" / "tabulado.txt").format == "csv"


def test_extension_en_mayusculas(fixtures_generados, tmp_path):
    destino = tmp_path / "DATOS.XLSX"
    shutil.copy(fixtures_generados / "formatos" / "varias_hojas.xlsx", destino)
    r = read_table(destino)
    assert r.format == "xlsx" and r.sheets == ["Clientes", "Ventas", "Resumen año"]


def test_formato_no_soportado(tmp_path):
    ruta = tmp_path / "a.pdf"
    ruta.write_bytes(b"x")
    with pytest.raises(FormatoNoSoportado):
        read_table(ruta)
    with pytest.raises(FormatoNoSoportado):
        list_sheets(ruta)


@pytest.mark.parametrize("nombre", ["no_existe.csv", "no_existe.xlsx"])
def test_no_encontrado(tmp_path, nombre):
    with pytest.raises(ArchivoNoEncontrado):
        read_table(tmp_path / nombre)
    with pytest.raises(ArchivoNoEncontrado):
        list_sheets(tmp_path / nombre)


def test_list_sheets_csv(fixtures_generados):
    s = list_sheets(fixtures_generados / "csv" / "excel_ar.csv")
    assert [(h.name, h.rows, h.columns) for h in s] == [("excel_ar.csv", 12, 4)]


def test_rendimiento_500k(fixtures_grandes):
    import time

    ruta = fixtures_grandes / "rendimiento_500k.xlsx"
    t = time.perf_counter()
    r = read_table(ruta)
    assert len(r.df) == 500_000
    assert time.perf_counter() - t < 30
    t = time.perf_counter()
    assert list_sheets(ruta)[0].rows == 500_000
    assert time.perf_counter() - t < 30
