"""Perfil de una tabla (core/profile.py)."""

from __future__ import annotations

import json
import threading
import time
from pathlib import Path

import pytest

from excel_tool.core.errors import Cancelado
from excel_tool.core.io import ReadOptions, SheetSummary, list_sheets, read_table
from excel_tool.core.profile import (
    ColumnProfile,
    TableProfile,
    column_letter,
    profile_table,
)
from excel_tool.core.progress import CancelToken, ProgressEvent

# ---------------------------------------------------------------- contra perfil.json


@pytest.fixture(scope="module")
def esperado(fixtures_esperados) -> dict:
    return json.loads((fixtures_esperados / "perfil.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module", params=["para_explorar.csv", "para_explorar.xlsx"])
def perfil(request, fixtures_generados) -> TableProfile:
    return profile_table(read_table(fixtures_generados / "explorar" / request.param))


def test_totales_coinciden_con_perfil_json(perfil, esperado):
    assert perfil.rows == esperado["filas"]
    assert perfil.columns == esperado["columnas"] == len(perfil.columns_profile)
    assert perfil.empty_rows == esperado["filas_vacias"]
    assert perfil.duplicate_rows == esperado["filas_repetidas"]


def test_columnas_coinciden_con_perfil_json(perfil, esperado):
    for pos, esp in enumerate(esperado["columnas_detalle"]):
        col = perfil.columns_profile[pos]
        ctx = (pos, esp["nombre"])
        assert col.position == pos and col.letter == esp["letra"], ctx
        if esp["nombre"]:
            assert col.name == esp["nombre"] or col.name.startswith(esp["nombre"]), ctx
        assert col.detected_type == esp["tipo"], ctx
        assert col.empty == esp["vacios"], ctx
        assert col.non_empty == perfil.rows - esp["vacios"], ctx
        assert col.distinct == esp["distintos"], ctx
        assert col.leading_zeros == esp["ceros_a_la_izquierda"], ctx
        assert col.untrimmed == esp["con_espacios_en_los_bordes"], ctx
        if esp["desglose_tipos"]:
            assert col.type_breakdown == esp["desglose_tipos"], ctx


def test_columna_mixta(perfil):
    mixta = perfil.columns_profile[7]
    assert mixta.type_breakdown == {"entero": 158, "texto": 38}
    assert mixta.not_converted == 38
    assert not mixta.summable  # 158 / 196 = 80 %
    assert mixta.sum is None
    assert any("38" in n for n in mixta.notes)
    assert any("Mixta" in w and "38" in w for w in perfil.warnings)


def test_numeros_y_fechas_como_texto_argentino(perfil):
    importe = perfil.columns_profile[1]
    assert importe.summable and importe.sum and importe.mean
    assert importe.min == "89,18" and importe.max == "9856,22"  # coma decimal, sin miles
    fecha = perfil.columns_profile[2]
    assert fecha.min == "03/01/2024" and fecha.max == "31/12/2024"
    assert perfil.columns_profile[3].min == "03/01/2024 09:53"
    assert perfil.columns_profile[0].min == "1" and perfil.columns_profile[0].max == "188"


def test_variantes_de_region(perfil):
    region = perfil.columns_profile[12]
    assert region.case_variants == 3
    assert region.few_values and region.distinct == 6
    assert any('"Norte"' in n for n in region.notes)
    assert any("Región" in w and "mayúsculas" in w for w in perfil.warnings)
    assert all(c.case_variants == 0 for c in perfil.columns_profile[:12])


def test_columna_vacia_y_pocos_valores(perfil):
    vacia = perfil.columns_profile[8]
    assert vacia.detected_type == "vacia" and not vacia.few_values and not vacia.is_unique
    assert vacia.top_values == [] and vacia.sample_values == []
    activo = perfil.columns_profile[4]
    assert activo.few_values and not activo.summable
    assert sorted(v for v, _ in activo.top_values) == ["No", "Sí"]


def test_archivo_y_hojas(fixtures_generados):
    ruta = fixtures_generados / "explorar" / "para_explorar.xlsx"
    hojas = list_sheets(ruta)
    p = profile_table(read_table(ruta), sheets=hojas)
    assert p.file.name == "para_explorar.xlsx" and p.file.format == "xlsx"
    assert p.file.size == ruta.stat().st_size
    assert p.sheets == hojas and p.sheet == read_table(ruta).sheet
    csv = profile_table(read_table(fixtures_generados / "explorar" / "para_explorar.csv"))
    assert csv.sheets == [] and csv.sheet is None
    assert csv.csv_options is not None and csv.file.format == "csv"


# ---------------------------------------------------------------- casos armados a mano


def _leer(tmp_path: Path, texto: str, ext: str = "csv"):
    ruta = tmp_path / f"t.{ext}"
    ruta.write_bytes((texto).encode("utf-8"))
    return read_table(ruta, ReadOptions())


def test_posible_clave_y_marcas(tmp_path):
    filas = "\n".join(f"{i};{'A' if i % 2 else 'B'};{i * 1000},50" for i in range(1, 11))
    p = profile_table(_leer(tmp_path, "Id;Tipo;Monto\n" + filas + "\n"))
    ident, tipo, monto = p.columns_profile
    assert ident.is_unique and not ident.few_values
    assert tipo.few_values and not tipo.is_unique
    assert tipo.top_values == [("A", 5), ("B", 5)] and tipo.sample_values == ["A", "B"]
    assert monto.min == "1000,5" and monto.max == "10000,5"
    assert monto.sum == pytest.approx(55005.0) and monto.mean == pytest.approx(5500.5)
    assert tipo.min_len == tipo.max_len == 1


def test_no_es_clave_si_hay_vacios(tmp_path):
    p = profile_table(_leer(tmp_path, "Id;X\n1;a\n2;b\n;c\n"))
    assert not p.columns_profile[0].is_unique


def test_marcas_de_calidad(tmp_path):
    texto = "Cod;Zona\n001;Norte\n002;NORTE\n003;Norte \n010; Sur\n11;Sur\n"
    p = profile_table(_leer(tmp_path, texto))
    cod, zona = p.columns_profile
    assert cod.leading_zeros == 4
    assert zona.case_variants == 5  # Norte/NORTE/Norte + " Sur"/Sur
    assert zona.untrimmed == 2
    assert zona.distinct == 5


def test_filas_vacias_y_repetidas(tmp_path):
    texto = "A;B\n1;x\n;\n1;x\n1;x\n;\n2;y\n"
    p = profile_table(_leer(tmp_path, texto))
    assert p.rows == 6
    assert p.empty_rows == 2  # las vacías no cuentan como repetidas
    assert p.duplicate_rows == 2


def test_celdas_con_espacios_no_son_repetidas_de_las_normales(tmp_path):
    p = profile_table(_leer(tmp_path, "A\nNorte\nNorte \n"))
    assert p.duplicate_rows == 0


def test_formato_de_numeros_decimales():
    from excel_tool.core.profile import _fecha_a_texto, _numero_a_texto

    assert _numero_a_texto(1234.5) == "1234,5"
    assert _numero_a_texto(-3.0) == "-3"
    assert _numero_a_texto(0.25) == "0,25"
    import pandas as pd

    assert _fecha_a_texto(pd.Timestamp(2024, 3, 5, 7, 8, 9), True) == "05/03/2024 07:08:09"
    assert _fecha_a_texto(pd.Timestamp(2024, 3, 5), False) == "05/03/2024"


def test_letras_de_columnas():
    assert [column_letter(i) for i in (0, 25, 26, 27, 701, 702)] == [
        "A",
        "Z",
        "AA",
        "AB",
        "ZZ",
        "AAA",
    ]


def test_tabla_sin_filas(tmp_path):
    p = profile_table(_leer(tmp_path, "A;B\n"))
    assert p.rows == 0 and p.empty_rows == 0 and p.duplicate_rows == 0
    assert p.columns_profile[0].detected_type == "vacia"
    assert not p.columns_profile[0].is_unique


# ---------------------------------------------------------------- on_column, progreso y cancelación


def test_on_column_y_progreso(fixtures_generados):
    data = read_table(fixtures_generados / "explorar" / "para_explorar.csv")
    llegadas: list[ColumnProfile] = []
    eventos: list[ProgressEvent] = []
    p = profile_table(data, on_column=llegadas.append, progress=eventos.append)
    assert [c.position for c in llegadas] == list(range(13))
    assert llegadas == p.columns_profile
    por_columna = [e for e in eventos if e.message.startswith("Analizando la columna")]
    assert [e.done for e in por_columna] == list(range(13))  # force: ninguna se descarta
    assert all(e.total == 13 for e in eventos)
    assert "Importe" in por_columna[1].message


def test_cancelar_antes_de_empezar(fixtures_generados):
    data = read_table(fixtures_generados / "explorar" / "para_explorar.csv")
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        profile_table(data, cancel=token)


def test_cancelar_a_mitad_de_las_columnas(fixtures_generados):
    data = read_table(fixtures_generados / "explorar" / "para_explorar.csv")
    token = CancelToken()
    llegadas: list[ColumnProfile] = []

    def al_llegar(col: ColumnProfile) -> None:
        llegadas.append(col)
        if len(llegadas) == 3:
            token.cancel()

    with pytest.raises(Cancelado):
        profile_table(data, on_column=al_llegar, cancel=token)
    assert len(llegadas) == 3


def test_sheets_vacio_por_defecto(tmp_path):
    p = profile_table(_leer(tmp_path, "A\n1\n"), sheets=[SheetSummary("Hoja", 1, 1)])
    assert p.sheets == [SheetSummary("Hoja", 1, 1)]


# ---------------------------------------------------------------- rendimiento (archivos grandes)


def test_rendimiento_csv_2m(fixtures_grandes):
    inicio = time.monotonic()
    data = read_table(fixtures_grandes / "rendimiento_2M.csv")
    p = profile_table(data)
    assert time.monotonic() - inicio <= 60
    assert p.rows == 2_000_000 and p.duplicate_rows == 0 and p.empty_rows == 0
    assert p.columns_profile[0].is_unique


def test_rendimiento_cancelar_corta_rapido(fixtures_grandes):
    data = read_table(fixtures_grandes / "rendimiento_2M.csv")
    peor = 0.0
    for demora in (1.0, 9.0):  # 9 s cae dentro de la detección de tipos de una columna grande
        token = CancelToken()
        threading.Timer(demora, token.cancel).start()
        inicio = time.monotonic()
        with pytest.raises(Cancelado):
            profile_table(data, cancel=token)
        peor = max(peor, time.monotonic() - inicio - demora)
    assert peor <= 2.0


def test_rendimiento_xlsx_500k(fixtures_grandes):
    inicio = time.monotonic()
    p = profile_table(read_table(fixtures_grandes / "rendimiento_500k.xlsx"))
    assert time.monotonic() - inicio <= 60
    assert p.rows == 500_000


# ---------------------------------------------------------------- exportar el informe

from openpyxl import load_workbook  # noqa: E402

from excel_tool.core.explore import value_counts  # noqa: E402
from excel_tool.core.profile import export_distincts, export_profile  # noqa: E402
from excel_tool.core.types import detect_column_type  # noqa: E402

ORIGENES = ["para_explorar.csv", "para_explorar.xlsx"]


def _leer_hoja(ruta: Path, hoja: str):
    return read_table(ruta, ReadOptions(sheet=hoja))


@pytest.fixture(params=ORIGENES)
def origen(request, fixtures_generados):
    data = read_table(fixtures_generados / "explorar" / request.param)
    return data, profile_table(data)


def test_exportar_informe_hojas_y_encabezados(origen, tmp_path):
    data, perfil = origen
    columna = perfil.columns_profile[1].name
    d = value_counts(data.df, [columna])
    ruta = tmp_path / "informe.xlsx"
    export_profile(perfil, [d], ruta)
    nombres = [h.name for h in list_sheets(ruta)]
    esperadas = ["General", "Columnas", columna[:31].replace("/", "_")]
    for n in esperadas:
        assert n in nombres
    assert ("Avisos" in nombres) == bool(perfil.warnings)

    general = _leer_hoja(ruta, "General").df
    assert list(general.columns) == ["Dato", "Valor"]
    datos = dict(zip(general["Dato"], general["Valor"], strict=True))
    assert datos["Filas"] == f"{perfil.rows:,}".replace(",", ".")
    assert datos["Nombre del archivo"] == perfil.file.name

    cols = _leer_hoja(ruta, "Columnas").df
    assert list(cols.columns[:7]) == [
        "Columna",
        "Nombre",
        "Tipo",
        "Con datos",
        "Vacías",
        "Valores distintos",
        "Marcas",
    ]
    assert len(cols) == perfil.columns
    primera = perfil.columns_profile[0]
    assert cols.iloc[0]["Nombre"] == primera.name
    assert cols.iloc[0]["Columna"] == "A (1)"
    assert cols.iloc[0]["Con datos"] == str(primera.non_empty)

    hoja_d = _leer_hoja(ruta, nombres[2]).df
    assert list(hoja_d.columns) == [columna, "Cantidad", "Porcentaje"]
    assert len(hoja_d) == len(d.table)


def test_exportar_informe_numeros_como_numero(origen, tmp_path):
    data, perfil = origen
    d = value_counts(data.df, [perfil.columns_profile[1].name])
    ruta = tmp_path / "informe.xlsx"
    export_profile(perfil, [d], ruta)
    wb = load_workbook(ruta)
    hoja = wb[wb.sheetnames[2]]
    assert isinstance(hoja.cell(2, 2).value, int)
    assert isinstance(hoja.cell(2, 3).value, (int, float))
    columnas = wb["Columnas"]
    assert isinstance(columnas.cell(2, 4).value, int)  # Con datos
    wb.close()


def test_exportar_informe_sin_distintos(origen, tmp_path):
    _, perfil = origen
    ruta = tmp_path / "informe.xlsx"
    export_profile(perfil, [], ruta)
    nombres = [h.name for h in list_sheets(ruta)]
    assert nombres[:2] == ["General", "Columnas"]
    assert len(nombres) == 2 + (1 if perfil.warnings else 0)


def test_exportar_informe_nombres_de_hoja_raros_y_repetidos(tmp_path):
    csv = tmp_path / "a.csv"
    csv.write_bytes(b"Una columna/con:nombre*muy?largo[de]verdad ok;General;b\nx;1;y\nx;2;z\n")
    data = read_table(csv)
    perfil = profile_table(data)
    nombre = data.df.columns[0]
    d1 = value_counts(data.df, [nombre])
    d2 = value_counts(data.df, [nombre])
    d3 = value_counts(data.df, ["General"])
    d4 = value_counts(data.df, [nombre, "b"])
    ruta = tmp_path / "informe.xlsx"
    export_profile(perfil, [d1, d2, d3, d4], ruta)
    nombres = [h.name for h in list_sheets(ruta)]
    assert len(nombres) == len({n.lower() for n in nombres}) == 6
    for n in nombres:
        assert 0 < len(n) <= 31 and not any(c in n for c in "[]:*?/\\")
    assert nombres[:2] == ["General", "Columnas"]
    assert "General (2)" in nombres


def test_exportar_informe_avisa_si_la_lista_quedo_cortada(tmp_path):
    csv = tmp_path / "a.csv"
    csv.write_bytes(("a;b\n" + "".join(f"{i};x\n" for i in range(20))).encode("utf-8"))
    data = read_table(csv)
    from excel_tool.core.explore import DistinctOptions

    d = value_counts(data.df, ["a"], DistinctOptions(limit=5))
    assert d.truncated
    ruta = tmp_path / "i.xlsx"
    export_profile(profile_table(data), [d], ruta)
    avisos = _leer_hoja(ruta, "Avisos").df
    assert any("incompleta" in a for a in avisos["Aviso"])


def _resultado_distintos(origen_):
    data, perfil = origen_
    return value_counts(data.df, [perfil.columns_profile[1].name])


def test_exportar_distintos_xlsx_ida_y_vuelta(origen, tmp_path):
    d = _resultado_distintos(origen)
    ruta = tmp_path / "d.xlsx"
    r = export_distincts(d, ruta, "xlsx")
    assert r.rows == len(d.table) and not r.unconverted
    wb = load_workbook(ruta)
    hoja = wb.active
    assert [c.value for c in hoja[1]] == [*d.columns, "Cantidad", "Porcentaje"]
    assert hoja.cell(2, 2).value == int(d.table.iloc[0]["Cantidad"])
    assert hoja.cell(2, 3).value == pytest.approx(float(d.table.iloc[0]["Porcentaje"]), abs=1e-3)
    wb.close()


def test_exportar_distintos_csv_ida_y_vuelta(origen, tmp_path):
    d = _resultado_distintos(origen)
    ruta = tmp_path / "d.csv"
    r = export_distincts(d, ruta, "csv_excel")
    assert r.rows == len(d.table)
    crudo = ruta.read_bytes()
    assert crudo.startswith(b"\xef\xbb\xbf")
    lineas = crudo.decode("utf-8-sig").splitlines()
    assert lineas[0].split(";")[1:] == ["Cantidad", "Porcentaje"]
    assert len(lineas) == len(d.table) + 1
    cantidad, porcentaje = lineas[1].split(";")[1:]
    assert cantidad == str(int(d.table.iloc[0]["Cantidad"]))
    assert "." not in porcentaje
    leido = read_table(ruta).df
    assert len(leido) == len(d.table)


def test_exportar_distintos_csv_google_usa_punto(tmp_path):
    csv = tmp_path / "a.csv"
    csv.write_bytes(b"a\nx\nx\ny\n")
    d = value_counts(read_table(csv).df, ["a"])
    ruta = tmp_path / "g.csv"
    export_distincts(d, ruta, "csv_google")
    lineas = ruta.read_text(encoding="utf-8").splitlines()
    assert lineas[0] == "a,Cantidad,Porcentaje"
    assert lineas[1].startswith("x,2,66.6")


# ---------------------------------------------------------------- cancelar la detección de tipos


def test_detect_column_type_cancelado_antes():
    import pandas as pd

    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        detect_column_type(pd.Series(["1", "2", "3"]), cancel=token)


def test_detect_column_type_cancela_entre_bloques(monkeypatch):
    import pandas as pd

    from excel_tool.core import types as t

    monkeypatch.setattr(t, "BLOQUE_UNICOS", 10)
    token = CancelToken()
    llamadas = {"n": 0}
    original = token.check

    def contar():
        llamadas["n"] += 1
        if llamadas["n"] == 4:
            token.cancel()
        original()

    token.check = contar  # type: ignore[method-assign]
    with pytest.raises(Cancelado):
        detect_column_type(pd.Series([str(i) for i in range(100)]), cancel=token)
    assert llamadas["n"] == 4


def test_detect_column_type_sin_cancel_igual():
    import pandas as pd

    s = pd.Series(["1", "2", "x"])
    assert detect_column_type(s, cancel=CancelToken()) == detect_column_type(s)
