"""Columnas que se llaman "Cantidad" o "Porcentaje" (F5-04): no chocan con las del resultado."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

from excel_tool.core.explore import value_counts
from excel_tool.core.io import ReadOptions, ReadResult, WriteOptions, read_table, write_table
from excel_tool.core.join import JoinPlan, SecondarySpec, SourceSpec, run_join
from excel_tool.core.limits import suggest_split_columns
from excel_tool.core.normalize import KeyOptions
from excel_tool.core.output import OutputSpec, save_result
from excel_tool.core.profile import export_distincts, export_profile, profile_table
from excel_tool.core.split import SplitSpec

_STR = pd.StringDtype("pyarrow")


def _df(**cols) -> pd.DataFrame:
    return pd.DataFrame({k: pd.Series(v, dtype=_STR) for k, v in cols.items()})


def _tabla() -> pd.DataFrame:
    return _df(
        Region=["N", "S", "N", "N", "S", "E"],
        Cantidad=["1", "2", "1", "3", "2", "1"],
        Porcentaje=["10", "20", "10", "30", "20", "10"],
    )


def test_value_counts_con_columna_cantidad():
    r = value_counts(_tabla(), ["Cantidad"])
    assert (r.count_column, r.percent_column) == ("Cantidad de filas", "Porcentaje")
    assert list(r.table.columns) == ["Cantidad", "Cantidad de filas", "Porcentaje"]
    fila = r.table[r.table["Cantidad"] == "1"].iloc[0]
    assert fila[r.count_column] == 3
    assert r.table[r.percent_column].sum() == pytest.approx(100)


def test_value_counts_con_ambas_y_con_otras():
    r = value_counts(_tabla(), ["Cantidad", "Porcentaje"])
    assert (r.count_column, r.percent_column) == ("Cantidad de filas", "Porcentaje de filas")
    assert int(r.table[r.count_column].max()) == 3
    r = value_counts(_tabla(), ["Region"])
    assert (r.count_column, r.percent_column) == ("Cantidad", "Porcentaje")
    assert list(r.table.columns) == ["Region", "Cantidad", "Porcentaje"]


def test_value_counts_si_igual_chocan():
    df = _df(**{"Cantidad": ["a", "b"], "Cantidad de filas": ["x", "x"]})
    r = value_counts(df, ["Cantidad", "Cantidad de filas"])
    assert r.count_column == "Cantidad de filas (2)"
    assert len(set(r.table.columns)) == 4


@pytest.mark.parametrize("formato", ["xlsx", "csv_excel"])
def test_suggest_split_columns_con_cantidad(formato):
    df = _tabla()
    sugeridas = suggest_split_columns(df, "excel", "xlsx")
    por_nombre = {s.column: s for s in sugeridas}
    assert por_nombre["Cantidad"].parts == 3
    assert por_nombre["Cantidad"].largest_part_rows == 3
    assert por_nombre["Region"].largest_part_rows == 3


@pytest.mark.parametrize("formato", ["xlsx", "csv_excel"])
def test_export_distincts_con_cantidad(tmp_path, formato):
    r = value_counts(_tabla(), ["Cantidad"])
    ext = "xlsx" if formato == "xlsx" else "csv"
    ruta = tmp_path / f"d.{ext}"
    export_distincts(r, ruta, formato)
    leido = read_table(ruta).df
    assert list(leido.columns) == ["Cantidad", "Cantidad de filas", "Porcentaje"]
    assert sorted(leido["Cantidad de filas"]) == ["1", "2"] or "3" in set(
        leido["Cantidad de filas"]
    )
    assert "3" in set(leido["Cantidad de filas"])


def test_export_profile_con_cantidad(tmp_path):
    df = _tabla()
    ruta_csv = tmp_path / "t.csv"
    write_table(df, ruta_csv, WriteOptions(format="csv_excel"))
    perfil = profile_table(read_table(ruta_csv))
    distintos = [value_counts(df, ["Cantidad"]), value_counts(df, ["Region"])]
    salida = tmp_path / "informe.xlsx"
    export_profile(perfil, distintos, salida)
    hoja = read_table(salida, ReadOptions(sheet="Cantidad")).df
    assert list(hoja.columns) == ["Cantidad", "Cantidad de filas", "Porcentaje"]


def test_profile_table_sin_archivo():
    df = _tabla()
    datos = ReadResult(df, Path("no_existe/resultado.csv"), "csv", None, [], None, 0, {}, [])
    antes = datetime.now()
    perfil = profile_table(datos, display_name="Resultado de la unión")
    assert perfil.file.name == "Resultado de la unión"
    assert perfil.file.size == 0
    assert perfil.file.modified >= antes
    assert perfil.rows == 6
    with pytest.raises(FileNotFoundError):
        profile_table(datos)  # compatible hacia atrás: sin nombre sigue mirando el archivo


@pytest.fixture(params=["csv", "xlsx"])
def escenario(request, fixtures_generados, tmp_path):
    base = fixtures_generados / "escenario_3"
    clientes = read_table(base / "clientes.xlsx").df
    ventas = read_table(base / "ventas.csv").df
    assert "Cantidad" in ventas.columns
    if request.param == "xlsx":
        ruta = tmp_path / "ventas.xlsx"
        write_table(ventas, ruta, WriteOptions(format="xlsx"))
        ventas = read_table(ruta).df
    return clientes, ventas


def test_unir_con_cantidad_sugerir_y_dividir(escenario, tmp_path):
    clientes, ventas = escenario
    ro = ReadOptions()
    principal = SourceSpec("clientes.xlsx", Path("clientes.xlsx"), ro, list(clientes.columns))
    sec = SecondarySpec(
        "ventas",
        Path("ventas"),
        ro,
        ["Fecha", "Artículo", "Cantidad", "Importe"],
        [("ID cliente", "ID cliente")],
    )
    res = run_join(
        JoinPlan(principal, KeyOptions(strip_leading_zeros=True), [sec]),
        {"clientes.xlsx": clientes, "ventas": ventas},
    )
    assert "Cantidad" in res.df.columns
    sugeridas = suggest_split_columns(res.df, "excel", "xlsx")
    assert sugeridas  # no explota
    columna = sugeridas[0].column
    spec = OutputSpec(
        tmp_path / "salida.csv",
        "csv_excel",
        "solo_guardar",
        split=SplitSpec("by_value", value_columns=[columna]),
    )
    guardado = save_result(res.df, spec)
    assert sum(p.rows for p in guardado.parts) == len(res.df)
