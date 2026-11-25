"""write_table / write_sheets: ida y vuelta, tipos, hojas y errores."""

from __future__ import annotations

import datetime as dt

import pandas as pd
import pytest
from openpyxl import load_workbook

from excel_tool.core.errors import ArchivoBloqueado, ExcelToolError
from excel_tool.core.io import ReadOptions, WriteOptions, read_table, write_sheets, write_table

FORMATOS = {
    "csv_excel": ".csv",
    "csv_google": ".csv",
    "csv_standard": ".csv",
    "xlsx": ".xlsx",
}


def _origenes(generados):
    rutas = sorted((generados / "csv").iterdir())
    rutas.append(generados / "escenario_3" / "ventas.csv")
    rutas.append(generados / "explorar" / "para_explorar.xlsx")
    return rutas


def _df_texto():
    return pd.DataFrame(
        {
            "nombre": ["Ñandú", "acción", "", "=SUMA(A1)"],
            "codigo": ["00123", "00124", "", "7"],
            "monto": ["1234,56", "10", "abc", ""],
            "dia": ["01/02/2024", "31/12/2023", "", "zzz"],
            "hora": ["01/02/2024 10:30:00", "", "05/06/2022 00:00:01", "x"],
            "cant": ["5", "1.000", "", "q"],
        }
    )


@pytest.mark.parametrize("formato", list(FORMATOS))
def test_ida_y_vuelta_fixtures(fixtures_generados, tmp_path, formato):
    for origen in _origenes(fixtures_generados):
        leido = read_table(origen).df
        destino = tmp_path / f"{origen.stem}_{formato}{FORMATOS[formato]}"
        res = write_table(leido, destino, WriteOptions(format=formato))
        assert res.rows == len(leido) and res.unconverted == {}
        # header_row=0: los datos de ida ya traen el encabezado correcto
        vuelta = read_table(destino, ReadOptions(header_row=0)).df
        pd.testing.assert_frame_equal(vuelta, leido, check_dtype=False)


@pytest.mark.parametrize("formato", list(FORMATOS))
def test_texto_dificil(tmp_path, formato):
    df = pd.DataFrame(
        {
            "a": ["Ñandú", "00123", "", "con;punto y coma", 'con "comillas"', "a,b"],
            "b": ["x", "=1+1", "línea\nnueva", "", "z", "w"],
        }
    )
    destino = tmp_path / f"t{FORMATOS[formato]}"
    write_table(df, destino, WriteOptions(format=formato))
    vuelta = read_table(destino, ReadOptions(header_row=0)).df
    if formato == "xlsx":
        assert load_workbook(destino)["Datos"]["B3"].data_type == "s"  # no es fórmula
    pd.testing.assert_frame_equal(vuelta, df, check_dtype=False)


def test_csv_estilos(tmp_path):
    df = pd.DataFrame({"n": ["Ñ"], "m": ["1234,56"], "e": ["7"]})
    tipos = {"m": "decimal", "e": "entero"}
    esperado = {
        "csv_excel": ("utf-8-sig", "n;m;e\r\nÑ;1234,56;7\r\n"),
        "csv_google": ("utf-8", "n,m,e\r\nÑ,1234.56,7\r\n"),
        "csv_standard": ("utf-8-sig", "n,m,e\r\nÑ,1234.56,7\r\n"),
    }
    for formato, (enc, texto) in esperado.items():
        ruta = tmp_path / f"{formato}.csv"
        write_table(df, ruta, WriteOptions(format=formato, column_types=tipos))
        crudo = ruta.read_bytes()
        assert crudo.startswith(b"\xef\xbb\xbf") == (enc == "utf-8-sig")
        assert crudo.decode(enc) == texto


def test_csv_tipos_y_no_convertibles(tmp_path):
    df = _df_texto()
    tipos = {
        "monto": "decimal",
        "dia": "fecha",
        "hora": "fecha_hora",
        "cant": "entero",
        "codigo": "texto",
    }
    ruta = tmp_path / "t.csv"
    res = write_table(df, ruta, WriteOptions(format="csv_excel", column_types=tipos))
    assert res.unconverted == {
        "monto": [(2, "abc")],
        "dia": [(3, "zzz")],
        "hora": [(3, "x")],
        "cant": [(3, "q")],
    }
    vuelta = read_table(ruta, ReadOptions(header_row=0)).df
    assert vuelta["monto"].tolist() == ["1234,56", "10", "abc", ""]
    assert vuelta["dia"].tolist() == ["01/02/2024", "31/12/2023", "", "zzz"]
    assert vuelta["hora"].tolist() == ["01/02/2024 10:30:00", "", "05/06/2022 00:00:01", "x"]
    assert vuelta["cant"].tolist() == ["5", "1000", "", "q"]
    assert vuelta["codigo"].tolist() == ["00123", "00124", "", "7"]


def test_xlsx_tipos_celdas(tmp_path):
    df = _df_texto()
    tipos = {"monto": "decimal", "dia": "fecha", "hora": "fecha_hora", "cant": "entero"}
    ruta = tmp_path / "t.xlsx"
    res = write_table(df, ruta, WriteOptions(format="xlsx", column_types=tipos, sheet_name="X"))
    assert set(res.unconverted) == {"monto", "dia", "hora", "cant"}
    ws = load_workbook(ruta)["X"]
    assert ws["C2"].value == 1234.56 and ws["C2"].data_type == "n"
    assert ws["C4"].value == "abc"
    assert ws["F3"].value == 1000 and isinstance(ws["F3"].value, int)
    assert ws["D2"].value == dt.datetime(2024, 2, 1) and ws["D2"].number_format == "dd/mm/yyyy"
    assert ws["E2"].value == dt.datetime(2024, 2, 1, 10, 30)
    assert ws["E2"].number_format == "dd/mm/yyyy hh:mm:ss"
    assert ws["B2"].value == "00123" and ws["B2"].data_type == "s"
    assert ws["D3"].value == dt.datetime(2023, 12, 31)
    assert ws["D4"].value is None
    # lo leído por read_table vuelve en formato argentino
    vuelta = read_table(ruta, ReadOptions(header_row=0)).df
    assert vuelta["dia"].tolist()[:2] == ["01/02/2024", "31/12/2023"]
    assert vuelta["monto"].tolist()[:2] == ["1234.56", "10"]


def test_xlsx_texto_sin_tipos_son_celdas_de_texto(tmp_path):
    ruta = tmp_path / "t.xlsx"
    write_table(pd.DataFrame({"a": ["1", "2.5"]}), ruta, WriteOptions(format="xlsx"))
    ws = load_workbook(ruta)["Datos"]
    assert [ws["A2"].data_type, ws["A3"].data_type] == ["s", "s"]


def test_xlsx_congelado_y_filtro(tmp_path):
    ruta = tmp_path / "t.xlsx"
    write_table(_df_texto(), ruta, WriteOptions(format="xlsx"))
    ws = load_workbook(ruta)["Datos"]
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref == "A1:F5"


def test_xlsx_sin_filas(tmp_path):
    ruta = tmp_path / "t.xlsx"
    write_table(pd.DataFrame({"a": [], "b": []}, dtype=str), ruta, WriteOptions(format="xlsx"))
    r = read_table(ruta, ReadOptions(header_row=0))
    assert list(r.df.columns) == ["a", "b"] and len(r.df) == 0


def test_nombre_de_hoja_invalido(tmp_path, caplog):
    ruta = tmp_path / "t.xlsx"
    largo = "x" * 40
    write_table(pd.DataFrame({"a": ["1"]}), ruta, WriteOptions(format="xlsx", sheet_name=largo))
    assert load_workbook(ruta).sheetnames == ["x" * 31]
    write_table(pd.DataFrame({"a": ["1"]}), ruta, WriteOptions(format="xlsx", sheet_name="a/b:c?"))
    assert load_workbook(ruta).sheetnames == ["a_b_c_"]
    assert "ajustado" in caplog.text


def test_write_sheets(tmp_path):
    ruta = tmp_path / "t.xlsx"
    a = pd.DataFrame({"a": ["1", "2"]})
    b = pd.DataFrame({"f": ["01/01/2020"], "g": ["x"]})
    res = write_sheets(
        {"Ventas": a, "ventas": a, "Año/2024": b, "z" * 40: a, "z" * 41: a},
        ruta,
        column_types={"Año/2024": {"f": "fecha"}},
    )
    assert [r.rows for r in res] == [2, 2, 1, 2, 2]
    wb = load_workbook(ruta)
    assert wb.sheetnames == ["Ventas", "ventas (2)", "Año_2024", "z" * 31, "z" * 27 + " (2)"]
    for ws in wb.worksheets:
        assert ws.freeze_panes == "A2" and ws.auto_filter.ref
    assert wb["Año_2024"]["A2"].value == dt.datetime(2020, 1, 1)


def test_errores(tmp_path, monkeypatch):
    df = pd.DataFrame({"a": ["1"]})
    with pytest.raises(ExcelToolError) as e:
        write_table(df, tmp_path / "no_existe" / "a.csv", WriteOptions(format="csv_excel"))
    assert not isinstance(e.value, ArchivoBloqueado)
    with pytest.raises(ExcelToolError):
        write_table(df, tmp_path / "no_existe" / "a.xlsx", WriteOptions(format="xlsx"))

    def bloqueado(*a, **k):
        raise PermissionError("bloqueado")

    monkeypatch.setattr("pathlib.Path.open", bloqueado)
    with pytest.raises(ArchivoBloqueado):
        write_table(df, tmp_path / "a.csv", WriteOptions(format="csv_excel"))
    monkeypatch.undo()
    # xlsxwriter envuelve el error de disco en FileCreateError: se sigue viendo como bloqueado
    monkeypatch.setattr("xlsxwriter.workbook.ZipFile", bloqueado)
    with pytest.raises(ArchivoBloqueado):
        write_table(df, tmp_path / "a.xlsx", WriteOptions(format="xlsx"))


def test_csv_bytes_mide_sin_escribir(tmp_path):
    from excel_tool.core.io import csv_bytes

    df = _df_texto()
    crudo = csv_bytes(df)
    destino = tmp_path / "a.csv"
    write_table(df, destino, WriteOptions(format="csv_excel"))
    escrito = destino.read_bytes().removeprefix(b"\xef\xbb\xbf")
    encabezado = escrito.split(b"\r\n", 1)[0] + b"\r\n"
    assert crudo == escrito[len(encabezado) :]  # sin encabezado, mismo contenido
    assert csv_bytes(df, header=True) == escrito
    assert list(tmp_path.iterdir()) == [destino]


def test_write_sheets_acepta_generador_y_avisa_renombre(tmp_path):
    df = _df_texto()

    def hojas():
        yield "Ventas/2024", df
        yield "Ventas/2024", df

    destino = tmp_path / "h.xlsx"
    res = write_sheets(hojas(), destino)
    assert load_workbook(destino).sheetnames == ["Ventas_2024", "Ventas_2024 (2)"]
    assert all(r.warnings for r in res)
    assert not list(tmp_path.glob("*.tmp"))


@pytest.mark.parametrize("fmt", ["csv_excel", "xlsx"])
def test_write_table_progreso_por_filas(tmp_path, fmt):
    df = pd.DataFrame({"A": pd.array([str(i) for i in range(60_000)], dtype="string[pyarrow]")})
    eventos = []
    write_table(df, tmp_path / "p.out", WriteOptions(format=fmt), progress=eventos.append)
    assert eventos[-1].done == eventos[-1].total == 60_000
