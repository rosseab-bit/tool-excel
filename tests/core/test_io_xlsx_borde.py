"""Escritor xlsx (xlsxwriter, ADR-033): casos borde."""

from __future__ import annotations

import datetime as dt
import tempfile
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

from excel_tool.core.errors import ArchivoBloqueado, Cancelado
from excel_tool.core.io import ReadOptions, WriteOptions, read_table, write_sheets, write_table
from excel_tool.core.progress import CancelToken


def _escribir(df, ruta, **kw):
    return write_table(df, ruta, WriteOptions(format="xlsx", **kw))


def test_texto_con_igual_queda_texto(tmp_path):
    ruta = tmp_path / "a.xlsx"
    _escribir(pd.DataFrame({"a": ["=1+1", "=SUMA(A1)", "+5", "@x"]}), ruta)
    ws = load_workbook(ruta)["Datos"]
    assert [ws.cell(i, 1).value for i in range(2, 6)] == ["=1+1", "=SUMA(A1)", "+5", "@x"]
    assert all(ws.cell(i, 1).data_type == "s" for i in range(2, 6))
    assert read_table(ruta, ReadOptions(header_row=0)).df["a"].tolist()[0] == "=1+1"


def test_no_convierte_texto_a_numero_ni_url(tmp_path):
    ruta = tmp_path / "a.xlsx"
    _escribir(pd.DataFrame({"a": ["007", "http://x.com", "1e3"]}), ruta)
    ws = load_workbook(ruta)["Datos"]
    assert [ws.cell(i, 1).value for i in range(2, 5)] == ["007", "http://x.com", "1e3"]
    assert ws["A3"].hyperlink is None


def test_caracteres_de_control_se_sacan(tmp_path):
    ruta = tmp_path / "a.xlsx"
    _escribir(pd.DataFrame({"a\x01b": ["x\x00y", "ok\ttab\nnl", "z\x1f"]}), ruta)
    ws = load_workbook(ruta)["Datos"]
    assert ws["A1"].value == "ab"
    assert [ws.cell(i, 1).value for i in range(2, 5)] == ["xy", "ok\ttab\nnl", "z"]


def test_varias_hojas_con_tipos_distintos(tmp_path):
    ruta = tmp_path / "h.xlsx"
    h1 = pd.DataFrame({"n": ["1", "2", "3"]})
    h2 = pd.DataFrame({"f": ["01/02/2024", "", "x"], "t": ["a", "b", "c"]})
    res = write_sheets(
        {"Uno": h1, "Dos": h2}, ruta, {"Uno": {"n": "entero"}, "Dos": {"f": "fecha"}}
    )
    wb = load_workbook(ruta)
    assert wb.sheetnames == ["Uno", "Dos"]
    assert wb["Uno"]["A3"].value == 2
    assert wb["Dos"]["A2"].value == dt.datetime(2024, 2, 1)
    assert wb["Dos"]["A4"].value == "x" and wb["Dos"]["B4"].value == "c"
    assert wb["Uno"].freeze_panes == "A2" and wb["Dos"].auto_filter.ref == "A1:B4"
    assert [r.rows for r in res] == [3, 3] and "f" in res[1].unconverted


def test_sin_congelar_ni_filtrar(tmp_path):
    ruta = tmp_path / "a.xlsx"
    _escribir(pd.DataFrame({"a": ["1"]}), ruta, freeze_header=False, autofilter=False)
    ws = load_workbook(ruta)["Datos"]
    assert ws.freeze_panes is None and not ws.auto_filter.ref


def test_encabezados_vacios_y_repetidos(tmp_path):
    ruta = tmp_path / "a.xlsx"
    df = pd.DataFrame([["1", "2", "3", "4"]], columns=["x", "", "x", "y"])
    _escribir(df, ruta)
    ws = load_workbook(ruta)["Datos"]
    assert [ws.cell(1, j).value for j in range(1, 5)] == ["x", None, "x", "y"]
    assert [ws.cell(2, j).value for j in range(1, 5)] == ["1", "2", "3", "4"]
    assert ws.auto_filter.ref == "A1:D2"


def test_fechas_anteriores_a_1900_van_como_texto_con_aviso(tmp_path):
    ruta = tmp_path / "a.xlsx"
    df = pd.DataFrame({"d": ["01/01/1850", "15/03/2020", "31/12/1899", "01/03/1900"]})
    res = _escribir(df, ruta, column_types={"d": "fecha"})
    ws = load_workbook(ruta)["Datos"]
    assert ws["A2"].value == "01/01/1850" and ws["A2"].data_type == "s"
    assert ws["A3"].value == dt.datetime(2020, 3, 15)
    assert ws["A4"].value == "31/12/1899"
    assert ws["A5"].value == dt.datetime(1900, 3, 1)
    assert len(res.warnings) == 1 and "1900" in res.warnings[0]


def test_fecha_hora_anterior_a_1900_como_texto(tmp_path):
    ruta = tmp_path / "a.xlsx"
    df = pd.DataFrame({"d": ["01/01/1850 10:30:00", "02/02/2020 01:02:03"]})
    res = _escribir(df, ruta, column_types={"d": "fecha_hora"})
    ws = load_workbook(ruta)["Datos"]
    assert ws["A2"].value == "01/01/1850 10:30:00"
    assert ws["A3"].value == dt.datetime(2020, 2, 2, 1, 2, 3)
    assert res.warnings


def test_celda_de_mas_de_32767_caracteres_se_recorta_y_avisa(tmp_path):
    ruta = tmp_path / "a.xlsx"
    df = pd.DataFrame({"a": ["x" * 40_000, "corto", "y" * 32_767]})
    res = _escribir(df, ruta)
    ws = load_workbook(ruta)["Datos"]
    assert len(ws["A2"].value) == 32_767 and ws["A3"].value == "corto"
    assert len(ws["A4"].value) == 32_767
    assert len(res.warnings) == 1 and "1 celdas" in res.warnings[0]


def test_numero_con_texto_largo_no_convertible_se_recorta(tmp_path):
    ruta = tmp_path / "a.xlsx"
    df = pd.DataFrame({"n": ["1", "z" * 33_000]})
    res = _escribir(df, ruta, column_types={"n": "entero"})
    ws = load_workbook(ruta)["Datos"]
    assert ws["A2"].value == 1 and len(ws["A3"].value) == 32_767
    assert res.warnings


def test_numeros_enteros_y_decimales(tmp_path):
    ruta = tmp_path / "a.xlsx"
    df = pd.DataFrame({"i": ["1", "", "3000000000"], "d": ["1,5", "2,25", ""]})
    _escribir(df, ruta, column_types={"i": "entero", "d": "decimal"})
    ws = load_workbook(ruta)["Datos"]
    assert ws["A2"].value == 1 and ws["A3"].value is None and ws["A4"].value == 3_000_000_000
    assert ws["B2"].value == pytest.approx(1.5) and ws["B3"].value == pytest.approx(2.25)


def test_bloqueado_da_archivo_bloqueado_y_no_deja_tmp(tmp_path, monkeypatch):
    def bloqueado(*a, **k):
        raise PermissionError("bloqueado")

    monkeypatch.setattr("xlsxwriter.workbook.ZipFile", bloqueado)
    with pytest.raises(ArchivoBloqueado):
        _escribir(pd.DataFrame({"a": ["1"]}), tmp_path / "a.xlsx")
    assert not list(tmp_path.iterdir())


def test_cancelar_no_deja_tmp_ni_destino_ni_temporales(tmp_path, monkeypatch):
    carpetas = []
    real = tempfile.TemporaryDirectory

    def espiar(*a, **k):
        t = real(*a, **k)
        carpetas.append(Path(t.name))
        return t

    monkeypatch.setattr(tempfile, "TemporaryDirectory", espiar)
    token = CancelToken()
    df = pd.DataFrame({"a": [str(i) for i in range(20_000)]})
    destino = tmp_path / "a.xlsx"

    def avance(ev):
        if ev.done >= 0:
            token.cancel()

    with pytest.raises(Cancelado):
        _escribir_con(df, destino, avance, token)
    assert not list(tmp_path.iterdir())
    assert carpetas and not any(c.exists() for c in carpetas)


def _escribir_con(df, destino, avance, token):
    return write_table(df, destino, WriteOptions(format="xlsx"), progress=avance, cancel=token)


def test_cancelar_respeta_archivo_previo(tmp_path):
    destino = tmp_path / "a.xlsx"
    _escribir(pd.DataFrame({"a": ["viejo"]}), destino)
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        _escribir_con(pd.DataFrame({"a": ["nuevo"]}), destino, None, token)
    assert load_workbook(destino)["Datos"]["A2"].value == "viejo"
    assert [p.name for p in tmp_path.iterdir()] == ["a.xlsx"]
