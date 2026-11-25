"""Integración de la Fase 2 (core): leer -> perfilar -> distintos -> filtrar/ordenar -> páginas ->
exportar -> releer. Solo API pública. Los esperados salen de `perfil.json` y de recontar el CSV con
el módulo `csv` de la biblioteca estándar (sin usar el producto)."""

from __future__ import annotations

import csv
import json
from collections import Counter

import pytest

from excel_tool.core.explore import (
    DistinctOptions,
    FilterCondition,
    filter_rows,
    get_page,
    rows_with_value,
    sort_rows,
    value_counts,
)
from excel_tool.core.io import ReadOptions, list_sheets, read_table
from excel_tool.core.profile import export_distincts, export_profile, profile_table

ARCHIVOS = ["explorar/para_explorar.csv", "explorar/para_explorar.xlsx"]


@pytest.fixture(scope="module")
def esperado(fixtures_esperados) -> dict:
    return json.loads((fixtures_esperados / "perfil.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def filas_csv(fixtures_generados) -> list[dict[str, str]]:
    """El CSV recontado de forma independiente (columna repetida y sin nombre: por posición)."""
    ruta = fixtures_generados / "explorar" / "para_explorar.csv"
    with ruta.open(encoding="utf-8-sig", newline="") as f:
        lector = csv.reader(f, delimiter=";")
        filas = list(lector)
    encabezado = filas[0]
    return [dict(enumerate(f[: len(encabezado)])) for f in filas[1:]]


@pytest.fixture(scope="module", params=ARCHIVOS, ids=["csv", "xlsx"])
def explorado(request, fixtures_generados):
    leido = read_table(fixtures_generados / request.param)
    return leido, profile_table(leido)


# ---------------------------------------------------------------- perfil vs perfil.json
def test_perfil_general_contra_perfil_json(explorado, esperado):
    _, perfil = explorado
    assert perfil.rows == esperado["filas"]
    assert perfil.columns == esperado["columnas"]
    assert perfil.empty_rows == esperado["filas_vacias"]
    assert perfil.duplicate_rows == esperado["filas_repetidas"]


def test_perfil_columnas_contra_perfil_json(explorado, esperado):
    _, perfil = explorado
    assert len(perfil.columns_profile) == len(esperado["columnas_detalle"])
    for col, esp in zip(perfil.columns_profile, esperado["columnas_detalle"], strict=True):
        etiqueta = f"{col.letter} {col.name}"
        assert col.position == esp["posicion"] - 1, etiqueta
        assert col.letter == esp["letra"], etiqueta
        if esp["nombre"]:
            # las repetidas ("Total", "Total") se numeran; el resto queda igual
            assert col.name.startswith(esp["nombre"]), etiqueta
        assert col.detected_type == esp["tipo"], etiqueta
        assert col.empty == esp["vacios"], etiqueta
        assert col.distinct == esp["distintos"], etiqueta
        assert col.non_empty + col.empty == esperado["filas"], etiqueta
        assert col.leading_zeros == esp["ceros_a_la_izquierda"], etiqueta
        assert col.untrimmed == esp["con_espacios_en_los_bordes"], etiqueta
        if esp["desglose_tipos"]:
            assert {str(k): v for k, v in col.type_breakdown.items()} == esp["desglose_tipos"]


def test_perfil_marcas_y_variantes(explorado):
    _, perfil = explorado
    por_nombre = {c.name: c for c in perfil.columns_profile}
    assert not any(c.is_unique for c in perfil.columns_profile)  # ninguna sin vacíos y única
    assert por_nombre["Activo"].few_values
    assert por_nombre["Región"].few_values
    assert por_nombre["Región"].case_variants == 3  # "Norte" / "NORTE" / "Norte "
    assert por_nombre["Mixta"].not_converted == 38
    assert por_nombre["Importe"].summable and por_nombre["Total"].summable
    assert not por_nombre["Descripción"].summable
    assert any("Región" in a for a in perfil.warnings)


def test_perfil_minimos_y_maximos_como_texto_argentino(explorado, filas_csv):
    _, perfil = explorado
    importe = {c.name: c for c in perfil.columns_profile}["Importe"]
    valores = [float(f[1].replace(",", ".")) for f in filas_csv if f[1] != ""]
    assert float(importe.max.replace(",", ".")) == max(valores)
    assert float(importe.min.replace(",", ".")) == min(valores)
    assert "," in importe.max or importe.max.isdigit()
    assert importe.sum == pytest.approx(sum(valores), rel=1e-9)


# ---------------------------------------------------------------- distintos
def test_value_counts_region_contra_recuento_independiente(explorado, filas_csv):
    leido, _ = explorado
    esperado = Counter(f[12] for f in filas_csv)  # "Región" es la columna 13
    resultado = value_counts(leido.df, ["Región"])
    obtenido = {
        ("" if v == "(vacío)" else v): int(n)
        for v, n in zip(resultado.table["Región"], resultado.table["Cantidad"], strict=True)
    }
    assert obtenido == dict(esperado)
    assert resultado.distinct_total == len(esperado)
    assert resultado.rows_total == 200
    assert not resultado.truncated
    assert [sorted(g) for g in resultado.variant_groups] == [["NORTE", "Norte", "Norte "]]
    assert resultado.table["Porcentaje"].sum() == pytest.approx(100.0)


def test_value_counts_juntando_variantes_y_buscando(explorado):
    leido, _ = explorado
    junto = value_counts(leido.df, ["Región"], DistinctOptions(normalize=True))
    textos = [v.strip().lower() for v in junto.table["Región"]]
    assert len(textos) == len(set(textos))
    assert int(junto.table.set_index(junto.table.columns[0]).loc["(vacío)", "Cantidad"]) == 4
    assert junto.variant_groups == []
    sin_vacios = value_counts(leido.df, ["Región"], DistinctOptions(include_empty=False))
    assert "(vacío)" not in set(sin_vacios.table["Región"])
    buscado = value_counts(leido.df, ["Región"], DistinctOptions(search="nor"))
    assert {v.strip().lower() for v in buscado.table["Región"]} == {"norte"}
    alfa = value_counts(leido.df, ["Región"], DistinctOptions(sort="alfabetico"))
    nombres = [v for v in alfa.table["Región"] if v != "(vacío)"]
    assert nombres == sorted(nombres, key=str.lower)


def test_value_counts_de_dos_columnas(explorado, filas_csv):
    leido, _ = explorado
    esperado = Counter((f[4], f[12]) for f in filas_csv)  # Activo + Región
    r = value_counts(leido.df, ["Activo", "Región"])
    assert r.distinct_total == len(esperado)
    assert int(r.table["Cantidad"].sum()) == 200


def test_rows_with_value_son_las_filas_correctas(explorado, filas_csv):
    leido, _ = explorado
    filas = rows_with_value(leido.df, ["Región"], ("Norte",))
    assert list(filas) == [i for i, f in enumerate(filas_csv) if f[12] == "Norte"]
    juntas = rows_with_value(leido.df, ["Región"], ("Norte",), normalize=True)
    assert list(juntas) == [i for i, f in enumerate(filas_csv) if f[12].strip().lower() == "norte"]
    vacias = rows_with_value(leido.df, ["Región"], ("",))
    assert len(vacias) == 4


# ---------------------------------------------------------------- filtros, orden, páginas
def _tipos(perfil):
    return {c.name: c.detected_type for c in perfil.columns_profile}


def test_filtro_contiene_y_mayor(explorado, filas_csv):
    leido, perfil = explorado
    tipos = _tipos(perfil)
    contiene = filter_rows(leido.df, [FilterCondition("Descripción", "contiene", "ÑANDÚ")], tipos)
    assert list(contiene) == [i for i, f in enumerate(filas_csv) if "ñandú" in f[5].lower()]
    mayor = filter_rows(leido.df, [FilterCondition("Importe", "mayor", "5000")], tipos)
    esperado = [i for i, f in enumerate(filas_csv) if f[1] and float(f[1].replace(",", ".")) > 5000]
    assert list(mayor) == esperado
    ambos = filter_rows(
        leido.df,
        [
            FilterCondition("Descripción", "contiene", "ñandú"),
            FilterCondition("Importe", "mayor", "5000"),
        ],
        tipos,
    )
    assert list(ambos) == sorted(set(contiene) & set(mayor))


def test_filtro_vacio_igual_y_valor_invalido(explorado, filas_csv):
    from excel_tool.core.errors import ExcelToolError

    leido, perfil = explorado
    tipos = _tipos(perfil)
    vacios = filter_rows(leido.df, [FilterCondition("Importe", "vacio")], tipos)
    assert len(vacios) == 10
    no_vacios = filter_rows(leido.df, [FilterCondition("Importe", "no_vacio")], tipos)
    assert len(no_vacios) == 190
    with pytest.raises(ExcelToolError) as exc:
        filter_rows(leido.df, [FilterCondition("Importe", "mayor", "abc")], tipos)
    assert exc.value.mensaje


def test_orden_numerico_con_vacios_al_final_y_paginas(explorado, filas_csv):
    leido, perfil = explorado
    tipos = _tipos(perfil)
    todas = filter_rows(leido.df, [], tipos)
    desc = sort_rows(leido.df, todas, "Importe", False, tipos["Importe"])
    valores = [float(filas_csv[i][1].replace(",", ".")) for i in desc[:190]]
    assert valores == sorted(valores, reverse=True)
    assert all(filas_csv[i][1] == "" for i in desc[190:])
    pagina = get_page(leido.df, desc, 0, 50)
    assert len(pagina) == 50
    assert list(pagina["N.º fila"]) == [int(i) + 1 for i in desc[:50]]
    ultima = get_page(leido.df, desc, 3, 50)
    assert len(ultima) == 50
    assert len(get_page(leido.df, desc, 4, 50)) == 0
    por_defecto = get_page(leido.df, desc, 0)
    assert len(por_defecto) == 200  # una sola página: tiene 200 filas


def test_orden_alfabetico_sin_importar_mayusculas(explorado, filas_csv):
    leido, perfil = explorado
    tipos = _tipos(perfil)
    filas = filter_rows(leido.df, [FilterCondition("Descripción", "no_vacio")], tipos)
    asc = sort_rows(leido.df, filas, "Descripción", True, tipos["Descripción"])
    textos = [filas_csv[i][5].lower() for i in asc]
    assert textos == sorted(textos)  # coincide con el orden Unicode (ver nota de F2-03 sobre Ñ)


# ---------------------------------------------------------------- exportar y releer
def test_export_profile_y_releer(explorado, tmp_path, filas_csv):
    leido, perfil = explorado
    region = value_counts(leido.df, ["Región"])
    destino = tmp_path / "informe.xlsx"
    export_profile(perfil, [region], destino)
    hojas = [h.name for h in list_sheets(destino)]
    assert hojas[:2] == ["General", "Columnas"]
    assert "Región" in hojas
    columnas = read_table(destino, ReadOptions(sheet="Columnas")).df
    assert len(columnas) == 13
    assert list(columnas["Nombre"])[:2] == ["ID", "Importe"]
    assert list(columnas["Vacías"])[:2] == ["4", "10"]
    general = read_table(destino, ReadOptions(sheet="General")).df
    datos = dict(zip(general["Dato"], general["Valor"], strict=True))
    assert datos["Filas"] == "200" and datos["Columnas"] == "13"
    assert datos["Filas vacías"] == "4" and datos["Filas repetidas"] == "8"
    relido = read_table(destino, ReadOptions(sheet="Región")).df
    assert sum(int(x) for x in relido["Cantidad"]) == 200


@pytest.mark.parametrize("formato,ext", [("xlsx", ".xlsx"), ("csv_excel", ".csv")])
def test_export_distincts_y_releer(explorado, tmp_path, formato, ext):
    leido, _ = explorado
    region = value_counts(leido.df, ["Región"])
    destino = tmp_path / f"valores{ext}"
    export_distincts(region, destino, formato)
    relido = read_table(destino).df
    assert list(relido.columns)[0] == "Región" and "Cantidad" in relido.columns
    assert len(relido) == region.distinct_total
    assert sum(int(x) for x in relido["Cantidad"]) == 200
    assert "Norte " in set(relido["Región"])  # el espacio del borde se conserva


# ---------------------------------------------------------------- varias hojas y otros formatos
def test_varias_hojas_se_perfila_cada_una(fixtures_generados):
    ruta = fixtures_generados / "formatos" / "varias_hojas.xlsx"
    hojas = list_sheets(ruta)
    assert [h.name for h in hojas] == ["Clientes", "Ventas", "Resumen año"]
    for hoja in hojas:
        leido = read_table(ruta, ReadOptions(sheet=hoja.name))
        perfil = profile_table(leido, sheets=hojas)
        assert perfil.sheet == hoja.name
        assert [h.name for h in perfil.sheets] == [h.name for h in hojas]
        assert perfil.columns == hoja.columns
        assert perfil.rows == hoja.rows
        assert len(perfil.columns_profile) == hoja.columns
        distintos = value_counts(leido.df, [perfil.columns_profile[0].name])
        assert distintos.rows_total == perfil.rows


@pytest.mark.parametrize("nombre", ["chico.xls", "chico.xlsb"])
def test_otros_formatos_dan_el_mismo_perfil_que_xlsx(fixtures_generados, nombre):
    base = profile_table(read_table(fixtures_generados / "formatos" / "varias_hojas.xlsx"))
    otro = profile_table(read_table(fixtures_generados / "formatos" / nombre))
    assert (otro.rows, otro.columns) == (base.rows, base.columns) == (12, 4)
    assert [c.name for c in otro.columns_profile] == [c.name for c in base.columns_profile]
    assert [c.detected_type for c in otro.columns_profile] == [
        c.detected_type for c in base.columns_profile
    ]
    assert [c.distinct for c in otro.columns_profile] == [c.distinct for c in base.columns_profile]


def test_ods_sin_fixture():
    pytest.skip("No hay fixture .ods (el generador no lo crea): ver informe de QA, F2-20.")
