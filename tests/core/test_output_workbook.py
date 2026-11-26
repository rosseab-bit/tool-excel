"""save_workbook / plan_workbook: varias hojas en un .xlsx (F5-02)."""

from __future__ import annotations

import dataclasses
import datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from excel_tool.core import limits as limits_mod
from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy
from excel_tool.core.errors import Cancelado, ExcelToolError
from excel_tool.core.io import ReadOptions, WriteOptions, list_sheets, read_table, write_table
from excel_tool.core.join import JoinPlan, SecondarySpec, SourceSpec, run_join
from excel_tool.core.normalize import KeyOptions
from excel_tool.core.output import SheetSpec, plan_workbook, save_workbook
from excel_tool.core.progress import CancelToken

STR = pd.StringDtype("pyarrow")


def _t(**cols):
    return pd.DataFrame({k: pd.Series(v, dtype=STR) for k, v in cols.items()})


@pytest.fixture
def limites(monkeypatch):
    def poner(**cambios):
        lim = dataclasses.replace(limits_mod.load_limits(), **cambios)
        monkeypatch.setattr(limits_mod, "load_limits", lambda: lim)
        return lim

    return poner


@pytest.fixture(params=["csv", "xlsx"])
def origen(request, tmp_path):
    """Principal y secundario leídos desde un CSV o un xlsx."""
    p = _t(id=["1", "2", "3"], nombre=["Ana", "Beto", "Cami"])
    v = _t(id=["1", "1", "2", "9"], monto=["10", "20,5", "5", "7"], dia=["2024-01-05"] * 4)
    out = {}
    for nombre, df in (("p", p), ("v", v)):
        f = tmp_path / f"{nombre}.{request.param}"
        write_table(df, f, WriteOptions(format="csv_excel" if request.param == "csv" else "xlsx"))
        out[nombre] = read_table(f).df
    return out


def _spec_union(origen, modo):
    ro = ReadOptions()
    p = origen["p"]
    resumen = AggregationSpec([GroupBy("id")], [Aggregation("monto", "sum", "Total")])
    sec = SecondarySpec(
        "v", Path("v"), ro, ["monto"], [("id", "id")], mode=modo,
        summary=resumen if modo == "summary" else None,
    )  # fmt: skip
    return JoinPlan(SourceSpec("p", Path("p"), ro, list(p.columns)), KeyOptions(), [sec])


def test_detalle_resumen_y_sin_coincidencia(origen, tmp_path):
    detalle = run_join(_spec_union(origen, "detail"), origen)
    resumen = run_join(_spec_union(origen, "summary"), origen)
    sin = detalle.unmatched["v"]
    assert len(sin) == 1
    hojas = [
        SheetSpec("Detalle", detalle.df, detalle.column_types),
        SheetSpec("Resumen", resumen.df, resumen.column_types),
        SheetSpec("Sin coincidencia en v", sin),
    ]
    destino = tmp_path / "libro.xlsx"
    res = save_workbook(hojas, destino, destination="excel")

    assert res.files == [destino] and res.folder is None
    assert [p.sheet for p in res.parts] == ["Detalle", "Resumen", "Sin coincidencia en v"]
    assert [p.rows for p in res.parts] == [len(detalle.df), len(resumen.df), 1]
    assert res.rows_total == len(detalle.df) + len(resumen.df) + 1
    assert res.warnings == []

    assert [h.name for h in list_sheets(destino)] == [
        "Detalle", "Resumen", "Sin coincidencia en v",
    ]  # fmt: skip
    assert len(read_table(destino, ReadOptions(sheet="Detalle")).df) == len(detalle.df)
    assert len(read_table(destino, ReadOptions(sheet="Sin coincidencia en v")).df) == 1

    wb = load_workbook(destino)
    for ws in wb.worksheets:
        assert ws.freeze_panes == "A2"
        assert ws.auto_filter.ref is not None
    ws = wb["Detalle"]
    cab = [c.value for c in ws[1]]
    col = cab.index("monto") + 1
    valores = [ws.cell(row=r, column=col).value for r in range(2, ws.max_row + 1)]
    assert all(v is None or isinstance(v, int | float) for v in valores)
    assert 20.5 in valores


def test_tipos_por_hoja(tmp_path):
    a = _t(f=["01/03/2024", "02/03/2024"], n=["1", "2"])
    b = _t(f=["2024-03-01", "x"])
    res = save_workbook(
        [
            SheetSpec("A", a, {"f": "fecha", "n": "entero"}),
            SheetSpec("B", b),  # sin tipos: texto
        ],
        tmp_path / "t.xlsx",
        destination="solo_guardar",
    )
    wb = load_workbook(res.files[0])
    assert isinstance(wb["A"]["A2"].value, datetime.datetime)
    assert wb["A"]["B2"].value == 1
    assert wb["B"]["A2"].value == "2024-03-01"


def test_no_convertidas_se_informan(tmp_path):
    a = _t(n=["1", "abc"])
    res = save_workbook(
        [SheetSpec("A", a, {"n": "entero"})], tmp_path / "t.xlsx", destination="solo_guardar"
    )
    assert res.unconverted == {"n": 1}


def test_nombres_repetidos_e_invalidos(tmp_path):
    t = _t(a=["1"])
    largo = "x" * 40
    res = save_workbook(
        [
            SheetSpec("Datos", t),
            SheetSpec("datos", t),
            SheetSpec("Ventas/2024: [uno]", t),
            SheetSpec(largo, t),
            SheetSpec("", t),
        ],
        tmp_path / "n.xlsx",
        destination="solo_guardar",
    )
    nombres = [p.sheet for p in res.parts]
    assert nombres[0] == "Datos"
    assert len({n.lower() for n in nombres}) == 5
    assert all(len(n) <= 31 and not set("[]:*?/\\") & set(n) for n in nombres)
    assert [h.name for h in list_sheets(res.files[0])] == nombres
    assert len(res.warnings) == 4  # todas menos la primera
    assert any('"datos"' in w for w in res.warnings)


def test_hoja_roja_da_error_con_su_nombre(limites, tmp_path):
    limites(excel_max_rows=4)
    chica = _t(a=["1", "2"])
    grande = _t(a=[str(i) for i in range(10)])
    destino = tmp_path / "r.xlsx"
    hojas = [SheetSpec("Resumen", chica), SheetSpec("Detalle", grande)]
    with pytest.raises(ExcelToolError) as exc:
        save_workbook(hojas, destino, destination="excel")
    msg = str(exc.value)
    assert "La hoja Detalle tiene 10 filas" in msg and "hasta 3 por hoja" in msg
    assert "dividido en partes" in msg and "Resumen" not in msg
    assert not destino.exists()
    assert plan_workbook(hojas, "excel").level == "obligatorio"


def test_google_sheets_total_del_libro(limites, tmp_path):
    limites(gs_max_cells=20)
    a = _t(x=["1"] * 5, y=["1"] * 5)  # 6 x 2 = 12 celdas: entra sola
    hojas = [SheetSpec("A", a), SheetSpec("B", a)]  # 24 en total
    assert plan_workbook(hojas[:1], "google_sheets").level != "obligatorio"
    chk = plan_workbook(hojas, "google_sheets")
    assert chk.level == "obligatorio" and any("entre todas las hojas" in r for r in chk.reasons)
    destino = tmp_path / "g.xlsx"
    with pytest.raises(ExcelToolError, match="entre todas las hojas"):
        save_workbook(hojas, destino, destination="google_sheets")
    assert not destino.exists()
    res = save_workbook(hojas[:1], destino, destination="google_sheets")
    assert res.rows_total == 5


def test_plan_ok_y_recomendado(limites):
    t = _t(a=["1"] * 5)
    assert plan_workbook([SheetSpec("A", t)], "excel").level == "ok"
    limites(excel_comfortable_rows=2)
    chk = plan_workbook([SheetSpec("A", t)], "excel")
    assert chk.level == "recomendado" and chk.reasons[0].startswith("Hoja A:")
    assert chk.suggested_parts >= 2


def test_sin_hojas(tmp_path):
    with pytest.raises(ExcelToolError):
        save_workbook([], tmp_path / "v.xlsx", destination="excel")


def test_cancelar_no_deja_archivo(tmp_path):
    t = _t(a=[str(i) for i in range(50)])
    destino = tmp_path / "c.xlsx"
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        save_workbook(
            [SheetSpec("A", t), SheetSpec("B", t)], destino, destination="excel", cancel=token
        )
    assert list(tmp_path.iterdir()) == []


def test_cancelar_a_mitad_conserva_archivo_previo(tmp_path):
    t = _t(a=[str(i) for i in range(50)])
    destino = tmp_path / "c.xlsx"
    destino.write_bytes(b"previo")
    token = CancelToken()

    def progreso(ev):
        token.cancel()

    with pytest.raises(Cancelado):
        save_workbook(
            [SheetSpec("A", t), SheetSpec("B", t)],
            destino,
            destination="excel",
            progress=progreso,
            cancel=token,
        )
    assert destino.read_bytes() == b"previo"
    assert [p.name for p in tmp_path.iterdir()] == ["c.xlsx"]


def test_progreso_llega_al_total(tmp_path):
    t = _t(a=["1", "2", "3"])
    eventos = []
    save_workbook(
        [SheetSpec("A", t), SheetSpec("B", t)],
        tmp_path / "p.xlsx",
        destination="excel",
        progress=eventos.append,
    )
    assert eventos and eventos[-1].done <= eventos[-1].total == 6
