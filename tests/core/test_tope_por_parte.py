"""F5-05 (F5-A): `SizeCheck.max_rows_per_part` y dividir sin cortar grupos con ese tope."""

from __future__ import annotations

import dataclasses

import pandas as pd
import pytest

from excel_tool.core.io import ReadOptions, WriteOptions, read_table, write_table
from excel_tool.core.limits import SizeEstimate, check, load_limits
from excel_tool.core.output import SheetSpec, SplitSpec, plan_split, plan_workbook

STR = pd.StringDtype("pyarrow")


def _est(filas: int, columnas: int = 4) -> SizeEstimate:
    return SizeEstimate(filas, columnas, (filas + 1) * columnas, 0.0, 0.0)


@pytest.fixture
def lim():
    # límites chicos para probar con pocas filas: lo cómodo es 300 y la hoja admite 1.000 filas
    return dataclasses.replace(
        load_limits(),
        excel_max_rows=1_001,
        excel_comfortable_rows=300,
        gs_comfortable_cells=1_200,
        gs_max_cells=10_000_000,
    )


@pytest.mark.parametrize("formato", ["xlsx", "csv_excel"])
def test_tope_por_destino(lim, formato):
    assert check(_est(5), "excel", formato, limits=lim).max_rows_per_part == 300
    # Google Sheets: celdas cómodas / columnas - encabezado
    assert check(_est(5, 4), "google_sheets", formato, limits=lim).max_rows_per_part == 299
    assert check(_est(5, 12), "google_sheets", formato, limits=lim).max_rows_per_part == 99


def test_tope_solo_guardar(lim):
    assert check(_est(5), "solo_guardar", "xlsx", limits=lim).max_rows_per_part == 1_000
    assert check(_est(5), "solo_guardar", "csv_excel", limits=lim).max_rows_per_part is None


def test_tope_en_el_libro(lim, monkeypatch):
    monkeypatch.setattr("excel_tool.core.limits.load_limits", lambda: lim)
    hojas = [
        SheetSpec("A", pd.DataFrame({"x": pd.Series(["1"] * 3, dtype=STR)})),
        SheetSpec("B", pd.DataFrame({c: pd.Series(["1"] * 3, dtype=STR) for c in "xyzw"})),
    ]
    assert plan_workbook(hojas, "excel").max_rows_per_part == 300
    # Google Sheets: manda la hoja con más columnas (1.200 / 4 - 1)
    assert plan_workbook(hojas, "google_sheets").max_rows_per_part == 299


@pytest.mark.parametrize("formato", ["csv", "xlsx"])
def test_dividir_sin_cortar_grupos_con_el_tope_entra_en_las_partes_sugeridas(
    lim, tmp_path, formato
):
    # 2.519 filas en grupos de 11 (el último, más corto): como el escenario grande, que con filas
    # por parte = suggested (280) desbordaba a una parte más con casi nada.
    ids = [str(i // 11) for i in range(2_519)]
    df = pd.DataFrame(
        {"cliente": pd.Series(ids, dtype=STR), "v": pd.Series(["1"] * 2_519, dtype=STR)}
    )
    if formato == "xlsx":  # pasa por un archivo real, como lo lee la app
        ruta = tmp_path / "t.xlsx"
        write_table(df, ruta, WriteOptions(format="xlsx"))
    else:
        ruta = tmp_path / "t.csv"
        write_table(df, ruta, WriteOptions(format="csv_excel"))
    df = read_table(ruta, ReadOptions()).df

    chk = check(_est(len(df), 2), "excel", "xlsx", limits=lim)
    assert chk.suggested_parts == 9
    assert chk.max_rows_per_part == 300

    def partes(filas):
        spec = SplitSpec("rows_keep_groups", rows_per_part=filas, group_column="cliente")
        return plan_split(df, spec, base_name="t")

    con_sugerido = partes(chk.suggested_rows_per_part)
    assert len(con_sugerido) == chk.suggested_parts + 1  # el problema que había
    con_tope = partes(chk.max_rows_per_part)
    assert chk.suggested_parts <= len(con_tope) <= chk.suggested_parts + 1
    assert all(p.rows <= chk.max_rows_per_part for p in con_tope)
    assert sum(p.rows for p in con_tope) == len(df)
