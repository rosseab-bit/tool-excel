"""save_result: un archivo, partes en carpeta, varias hojas, límites, cancelación y nombres."""

from __future__ import annotations

import dataclasses

import pandas as pd
import pytest
from openpyxl import load_workbook

from excel_tool.core import limits as limits_mod
from excel_tool.core.errors import Cancelado, ExcelToolError
from excel_tool.core.io import WriteOptions, read_table, write_table
from excel_tool.core.output import OutputSpec, save_result
from excel_tool.core.progress import CancelToken, ProgressEvent
from excel_tool.core.split import SplitSpec

STR = pd.StringDtype("pyarrow")


@pytest.fixture(params=["csv", "xlsx"])
def tabla(request, fixtures_generados):
    return read_table(fixtures_generados / "explorar" / f"para_explorar.{request.param}").df


@pytest.fixture
def limites(monkeypatch):
    """Cambia los límites para probar con tablas chicas."""

    def poner(**cambios):
        lim = dataclasses.replace(limits_mod.load_limits(), **cambios)
        monkeypatch.setattr(limits_mod, "load_limits", lambda: lim)
        return lim

    return poner


def _spec(tmp_path, nombre="salida", fmt="xlsx", dest="solo_guardar", **kw):
    ext = ".xlsx" if fmt == "xlsx" else ".csv"
    return OutputSpec(path=tmp_path / (nombre + ext), format=fmt, destination=dest, **kw)


def _leer(path):
    return read_table(path).df


# ---------------------------------------------------------------------------- un archivo


def test_un_archivo_xlsx_congelado_y_filtro(tabla, tmp_path):
    res = save_result(tabla, _spec(tmp_path))
    assert res.files == [tmp_path / "salida.xlsx"] and res.folder is None
    assert res.rows_total == len(tabla) == res.parts[0].rows
    ws = load_workbook(res.files[0]).active
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref == f"A1:{chr(64 + len(tabla.columns))}{len(tabla) + 1}"
    assert ws.max_row == len(tabla) + 1


def test_sin_congelar_ni_filtro(tabla, tmp_path):
    destino = tmp_path / "x.xlsx"
    write_table(tabla, destino, WriteOptions(format="xlsx", freeze_header=False, autofilter=False))
    ws = load_workbook(destino).active
    assert ws.freeze_panes is None and not ws.auto_filter.ref


@pytest.mark.parametrize("fmt", ["csv_excel", "csv_google", "xlsx"])
def test_un_archivo_formatos(tabla, tmp_path, fmt):
    res = save_result(tabla, _spec(tmp_path, fmt=fmt))
    leido = _leer(res.files[0])
    assert len(leido) == len(tabla) and list(leido.columns) == list(tabla.columns)


def test_column_types_se_aplican(tabla, tmp_path):
    spec = _spec(tmp_path, column_types={"ID": "entero", "Fecha": "fecha"})
    res = save_result(tabla, spec)
    ws = load_workbook(res.files[0]).active
    assert isinstance(ws["A2"].value, int)
    assert ws["C2"].number_format == "dd/mm/yyyy"
    assert res.unconverted == {}


def test_no_convertidas_se_informan(tmp_path):
    df = pd.DataFrame({"A": pd.array(["1", "x", "3", "y"], dtype=STR)})
    res = save_result(df, _spec(tmp_path, column_types={"A": "entero"}))
    assert res.unconverted == {"A": 2}


def test_cero_filas_solo_encabezado(tabla, tmp_path):
    vacia = tabla.iloc[0:0]
    spec = _spec(tmp_path, split=SplitSpec("rows", rows_per_part=10))
    res = save_result(vacia, spec)
    assert res.folder is None and res.rows_total == 0
    assert list(tmp_path.iterdir()) == [tmp_path / "salida.xlsx"]
    ws = load_workbook(res.files[0]).active
    assert [c.value for c in ws[1]] == list(tabla.columns)
    res = save_result(vacia, _spec(tmp_path, "otro", fmt="csv_excel"))
    assert (tmp_path / "otro.csv").read_text(encoding="utf-8-sig").count("\n") == 1


def test_progreso_y_cancelacion_en_un_archivo(tmp_path):
    df = pd.DataFrame({"A": pd.array([str(i) for i in range(12_000)], dtype=STR)})
    eventos: list[ProgressEvent] = []
    save_result(df, _spec(tmp_path, "a"), progress=eventos.append)
    assert eventos[-1].done == eventos[-1].total == 12_000
    assert any(0 < e.done < 12_000 for e in eventos) or len(eventos) >= 2

    token = CancelToken()
    token.cancel()
    for fmt in ("xlsx", "csv_excel"):
        with pytest.raises(Cancelado):
            save_result(df, _spec(tmp_path, "c", fmt=fmt), cancel=token)
    assert not [p for p in tmp_path.iterdir() if p.name.startswith("c")]


def test_cancelar_no_toca_archivo_existente(tmp_path):
    df = pd.DataFrame({"A": pd.array(["1", "2"], dtype=STR)})
    destino = tmp_path / "salida.csv"
    destino.write_text("viejo", encoding="utf-8")
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        save_result(df, _spec(tmp_path, fmt="csv_excel"), cancel=token)
    assert destino.read_text(encoding="utf-8") == "viejo"
    assert [p.name for p in tmp_path.iterdir()] == ["salida.csv"]


# ------------------------------------------------------------------------------ límites


def test_rojo_sin_division_da_error(tabla, tmp_path, limites):
    limites(excel_max_rows=101)
    with pytest.raises(ExcelToolError, match="no entra"):
        save_result(tabla, _spec(tmp_path, dest="excel"))
    assert list(tmp_path.iterdir()) == []


def test_rojo_con_division_respeta_el_limite(tabla, tmp_path, limites):
    limites(excel_max_rows=101)  # 100 filas de datos por hoja
    spec = _spec(tmp_path, dest="excel", split=SplitSpec("rows", rows_per_part=10_000))
    res = save_result(tabla, spec)
    assert [p.rows for p in res.parts] == [100, 100]
    assert res.rows_total == len(tabla)


def test_google_sheets_con_hojas_da_error(tabla, tmp_path):
    spec = _spec(tmp_path, dest="google_sheets", split=SplitSpec("sheets", rows_per_part=50))
    with pytest.raises(ExcelToolError, match="Google Sheets"):
        save_result(tabla, spec)
    assert list(tmp_path.iterdir()) == []


def test_google_sheets_parte_pesada_avisa(tabla, tmp_path, limites):
    limites(gs_max_import_mb=0.0005)
    spec = _spec(
        tmp_path, fmt="csv_google", dest="google_sheets", split=SplitSpec("rows", rows_per_part=100)
    )
    res = save_result(tabla, spec)
    assert any("Google Sheets puede no importarla" in w for w in res.warnings)


# ------------------------------------------------------------------------------ partes


def _indice(carpeta):
    ws = load_workbook(carpeta / "_indice.xlsx").active
    return [list(f) for f in ws.iter_rows(values_only=True)]


@pytest.mark.parametrize("fmt", ["xlsx", "csv_excel"])
def test_partes_por_filas_con_indice(tabla, tmp_path, fmt):
    spec = _spec(tmp_path, fmt=fmt, split=SplitSpec("rows", rows_per_part=60))
    res = save_result(tabla, spec)
    assert res.folder == tmp_path / "salida"
    assert [p.rows for p in res.parts] == [60, 60, 60, 20]
    assert sum(p.rows for p in res.parts) == res.rows_total == len(tabla)
    ext = ".xlsx" if fmt == "xlsx" else ".csv"
    assert res.files[-1].name == "_indice.xlsx" and len(res.files) == 5
    assert {f.suffix for f in res.files[:-1]} == {ext}
    escritas = [len(_leer(p.path)) for p in res.parts]
    assert escritas == [p.rows for p in res.parts]
    # El índice coincide con lo escrito
    filas = _indice(res.folder)
    assert filas[0] == ["Archivo", "Hoja", "Filas", "Desde fila", "Hasta fila", "Valor"]
    for fila, parte in zip(filas[1:], res.parts, strict=True):
        assert fila[0] == parte.path.name
        assert fila[2] == parte.rows and fila[3] == parte.first_row and fila[4] == parte.last_row
        assert fila[1] == ("Datos" if fmt == "xlsx" else None) or fila[1] is None
    assert filas[1][3] == 1 and filas[-1][4] == len(tabla)


def test_sin_indice(tabla, tmp_path):
    spec = _spec(tmp_path, write_index=False, split=SplitSpec("rows", rows_per_part=100))
    res = save_result(tabla, spec)
    assert not (res.folder / "_indice.xlsx").exists() and len(res.files) == 2


def test_partes_por_valor_y_subdivision(tabla, tmp_path):
    split = SplitSpec("by_value", value_columns=["Región"], rows_per_part=20)
    res = save_result(tabla, _spec(tmp_path, fmt="csv_excel", split=split))
    assert sum(p.rows for p in res.parts) == len(tabla)
    assert max(p.rows for p in res.parts) <= 20
    assert any("parte_1_de_" in p.path.name for p in res.parts)  # subdividió la región grande
    por_valor: dict[str, int] = {}
    for p in res.parts:
        por_valor[p.value[0]] = por_valor.get(p.value[0], 0) + p.rows
    assert por_valor == tabla["Región"].value_counts().to_dict()
    filas = _indice(res.folder)
    assert {f[5] for f in filas[1:]} == {v or "(vacío)" for v in por_valor}


def test_partes_sin_cortar_grupos(tabla, tmp_path):
    split = SplitSpec("rows_keep_groups", rows_per_part=50, group_column="Región")
    res = save_result(tabla, _spec(tmp_path, split=split))
    vistos = {}
    for i, p in enumerate(res.parts):
        for v in _leer(p.path)["Región"].unique():
            assert vistos.setdefault(v, i) == i  # un valor no aparece en dos partes
    assert sum(p.rows for p in res.parts) == len(tabla)


def test_grupo_mas_grande_que_el_limite_avisa(tabla, tmp_path):
    split = SplitSpec("rows_keep_groups", rows_per_part=5, group_column="Activo")
    res = save_result(tabla, _spec(tmp_path, split=split))
    assert any("no se puede cortar" in w for w in res.warnings)


def test_varias_hojas(tabla, tmp_path):
    spec = _spec(
        tmp_path, split=SplitSpec("sheets", rows_per_part=70), column_types={"ID": "entero"}
    )
    res = save_result(tabla, spec)
    assert res.folder is None and res.files == [tmp_path / "salida.xlsx"]
    wb = load_workbook(res.files[0])
    assert len(wb.sheetnames) == 3 == len(res.parts)
    for ws, parte in zip(wb.worksheets, res.parts, strict=True):
        assert ws.title == parte.sheet
        assert ws.max_row == parte.rows + 1
        assert ws.freeze_panes == "A2" and ws.auto_filter.ref
        assert isinstance(ws["A2"].value, int)
    assert sum(p.rows for p in res.parts) == len(tabla)


def test_carpeta_existente_usa_numero(tabla, tmp_path):
    spec = _spec(tmp_path, split=SplitSpec("rows", rows_per_part=100))
    (tmp_path / "salida").mkdir()
    (tmp_path / "salida" / "mio.txt").write_text("no tocar", encoding="utf-8")
    res = save_result(tabla, spec)
    assert res.folder == tmp_path / "salida (2)"
    res2 = save_result(tabla, spec)
    assert res2.folder == tmp_path / "salida (3)"
    assert [p.name for p in (tmp_path / "salida").iterdir()] == ["mio.txt"]


def test_cancelar_a_mitad_deja_todo_como_estaba(tabla, tmp_path):
    spec = _spec(tmp_path, fmt="csv_excel", split=SplitSpec("rows", rows_per_part=40))
    token = CancelToken()
    partes_vistas = []

    def progreso(ev: ProgressEvent) -> None:
        partes_vistas.append(ev.message)
        if "parte 3" in ev.message:
            token.cancel()

    with pytest.raises(Cancelado):
        save_result(tabla, spec, progress=progreso, cancel=token)
    assert any("parte 3" in m for m in partes_vistas)
    assert list(tmp_path.iterdir()) == []  # ni carpeta ni archivos


def test_cancelar_respeta_lo_que_ya_habia(tabla, tmp_path):
    (tmp_path / "salida").mkdir()
    (tmp_path / "salida" / "mio.txt").write_text("x", encoding="utf-8")
    spec = _spec(tmp_path, split=SplitSpec("rows", rows_per_part=40))
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        save_result(tabla, spec, cancel=token)
    assert sorted(p.name for p in tmp_path.iterdir()) == ["salida"]
    assert [p.name for p in (tmp_path / "salida").iterdir()] == ["mio.txt"]


def test_cancelar_en_hojas_no_deja_archivo(tabla, tmp_path):
    spec = _spec(tmp_path, split=SplitSpec("sheets", rows_per_part=40))
    token = CancelToken()

    def progreso(ev: ProgressEvent) -> None:
        if "parte 2" in ev.message:
            token.cancel()

    with pytest.raises(Cancelado):
        save_result(tabla, spec, progress=progreso, cancel=token)
    assert list(tmp_path.iterdir()) == []


def test_error_a_mitad_limpia(tabla, tmp_path, monkeypatch):
    from excel_tool.core import output

    original = output.write_table
    llamadas = []

    def falla(*a, **k):
        llamadas.append(1)
        if len(llamadas) == 3:
            raise ExcelToolError("boom")
        return original(*a, **k)

    monkeypatch.setattr(output, "write_table", falla)
    spec = _spec(tmp_path, split=SplitSpec("rows", rows_per_part=40))
    with pytest.raises(ExcelToolError, match="boom"):
        save_result(tabla, spec)
    assert list(tmp_path.iterdir()) == []


def test_suma_de_filas_incorrecta_da_error(tabla, tmp_path, monkeypatch):
    from excel_tool.core import output

    original = output.write_table

    def mal(df, path, options, **k):
        res = original(df, path, options, **k)
        res.rows += 1
        return res

    monkeypatch.setattr(output, "write_table", mal)
    spec = _spec(tmp_path, split=SplitSpec("rows", rows_per_part=100))
    with pytest.raises(ExcelToolError, match="no coincide"):
        save_result(tabla, spec)
    assert list(tmp_path.iterdir()) == []


# ------------------------------------------------------------------------------- nombres


def test_nombres_raros_validos_en_windows(tmp_path):
    df = pd.DataFrame(
        {
            "Zona": pd.array(["Norte/Sur", "CON", "NUL", "a:b?", "", "Norte/Sur"], dtype=STR),
            "N": pd.array(list("123456"), dtype=STR),
        }
    )
    split = SplitSpec("by_value", value_columns=["Zona"])
    res = save_result(df, _spec(tmp_path, fmt="csv_excel", split=split))
    assert sum(p.rows for p in res.parts) == 6
    for p in res.parts:
        assert p.path.exists()
        assert not any(c in p.path.name for c in '<>:"/\\|?*')
        assert p.path.stem.upper() not in {"CON", "NUL"}
    assert len({p.path.name.lower() for p in res.parts}) == len(res.parts)
    filas = _indice(res.folder)
    assert {f[5] for f in filas[1:]} == {"Norte/Sur", "CON", "NUL", "a:b?", "(vacío)"}


def test_ruta_larga_acorta_nombres_y_avisa(tabla, tmp_path):
    larga = tmp_path
    while len(str(larga)) < 170:
        larga = larga / ("c" * 40)
    larga.mkdir(parents=True)
    split = SplitSpec(
        "by_value", value_columns=["Descripción"], name_pattern="{nombre}_{valor}_" + "x" * 60
    )
    spec = OutputSpec(
        path=larga / "datos.csv", format="csv_excel", destination="solo_guardar", split=split
    )
    res = save_result(tabla, spec)
    assert any("acortaron" in w for w in res.warnings)
    assert sum(p.rows for p in res.parts) == len(tabla)
    assert all(len(str(p.path)) <= 250 and p.path.exists() for p in res.parts)
    assert len({p.path.name.lower() for p in res.parts}) == len(res.parts)
    assert len(_indice(res.folder)) == len(res.parts) + 1


def test_ruta_de_carpeta_imposible_da_error(tabla, tmp_path):
    larga = tmp_path
    while len(str(larga)) < 240:
        larga = larga / ("c" * 40)
    larga.mkdir(parents=True)
    spec = OutputSpec(
        path=larga / "d.csv",
        format="csv_excel",
        destination="solo_guardar",
        split=SplitSpec("rows", rows_per_part=100),
    )
    with pytest.raises(ExcelToolError, match="demasiado larga"):
        save_result(tabla, spec)
    assert not (larga / "d").exists()


def test_carpeta_destino_inexistente(tabla, tmp_path):
    spec = OutputSpec(
        path=tmp_path / "no_existe" / "d.csv",
        format="csv_excel",
        destination="solo_guardar",
        split=SplitSpec("rows", rows_per_part=100),
    )
    with pytest.raises(ExcelToolError, match="no existe"):
        save_result(tabla, spec)


def test_hoja_renombrada_va_a_warnings(tabla, tmp_path):
    spec = _spec(tmp_path, sheet_name="Ventas/2024: resumen largo largo largo")
    res = save_result(tabla, spec)
    assert any("nombre de hoja" in w for w in res.warnings)
    assert load_workbook(res.files[0]).sheetnames[0] == res.parts[0].sheet
