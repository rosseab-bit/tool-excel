"""Archivos grandes: solo corren si están generados (si no, se saltean con aviso)."""

from __future__ import annotations

import csv
import json

import pytest
from python_calamine import CalamineWorkbook


@pytest.fixture(scope="module")
def esperado(fixtures_esperados):
    return json.loads((fixtures_esperados / "grandes.json").read_text(encoding="utf-8"))


def _filas(ruta):
    with ruta.open(encoding="utf-8-sig", newline="") as f:
        lector = csv.reader(f, delimiter=";")
        encabezado = next(lector)
        yield encabezado
        yield from lector


def test_ventas_y_clientes_por_region(fixtures_grandes, esperado):
    it = _filas(fixtures_grandes / "clientes_grande.csv")
    assert next(it) == ["ID cliente", "Nombre", "Región", "Vendedor"]
    region_de = {fila[0]: fila[2] for fila in it}
    assert len(region_de) == esperado["filas"]["clientes_grande.csv"]
    por_region_clientes: dict[str, int] = {}
    for r in region_de.values():
        por_region_clientes[r] = por_region_clientes.get(r, 0) + 1
    assert por_region_clientes == esperado["clientes_por_region"]

    it = _filas(fixtures_grandes / "ventas_grande.csv")
    assert next(it) == ["ID cliente", "Fecha", "Artículo", "Cantidad", "Importe"]
    ventas: dict[str, int] = {}
    con_ventas: set[str] = set()
    total = 0
    for fila in it:
        total += 1
        con_ventas.add(fila[0])
        r = region_de[fila[0]]
        ventas[r] = ventas.get(r, 0) + 1
    assert total == esperado["filas"]["ventas_grande.csv"]
    assert ventas == esperado["ventas_por_region"]

    sin_ventas = len(region_de) - len(con_ventas)
    assert sin_ventas == esperado["clientes_sin_ventas"]["total"]
    detalle = {r: n + esperado["clientes_sin_ventas"]["por_region"][r] for r, n in ventas.items()}
    assert detalle == esperado["detalle_left"]["por_region"]
    assert sum(detalle.values()) == esperado["detalle_left"]["total"]

    lim = esperado["limites_excel"]
    assert [r for r, n in detalle.items() if n > lim["max_rows"]] == (
        esperado["regiones_que_superan_el_maximo_de_filas_de_excel"]
    )
    assert sorted(r for r, n in detalle.items() if n > lim["comfortable_rows"]) == sorted(
        esperado["regiones_que_superan_el_limite_comodo_de_excel"]
    )


def test_limites_del_esperado_coinciden_con_el_producto(esperado, generador_fixtures):
    lim = generador_fixtures._limites_producto()["excel"]
    assert esperado["limites_excel"] == {
        "max_rows": lim["max_rows"],
        "comfortable_rows": lim["comfortable_rows"],
    }


def test_rendimiento_csv_2m_filas(fixtures_grandes, esperado):
    it = _filas(fixtures_grandes / "rendimiento_2M.csv")
    encabezado = next(it)
    assert len(encabezado) == 10
    n = 0
    for fila in it:
        n += 1
        if n % 100_000 == 0:
            assert len(fila) == 10
    assert n == esperado["filas"]["rendimiento_2M.csv"]


def test_rendimiento_xlsx_500k_filas(fixtures_grandes, esperado):
    libro = CalamineWorkbook.from_path(str(fixtures_grandes / "rendimiento_500k.xlsx"))
    assert libro.sheet_names == ["Datos"]
    hoja = libro.get_sheet_by_name("Datos")
    # Ojo: total_height de calamine da una fila menos en este archivo (sin <dimension>);
    # se cuenta con to_python, que devuelve encabezado + datos.
    filas = hoja.to_python(skip_empty_area=False)
    assert len(filas) == esperado["filas"]["rendimiento_500k.xlsx"] + 1
    assert filas[0][0] == "ID" and filas[0][5] == "Descripción"
    assert filas[-1][0] == str(esperado["filas"]["rendimiento_500k.xlsx"])
