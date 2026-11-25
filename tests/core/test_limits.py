import json
import re
from dataclasses import replace
from pathlib import Path

import pandas as pd
import pytest

from excel_tool import platform_utils
from excel_tool.core import limits
from excel_tool.core.errors import Cancelado
from excel_tool.core.io import WriteOptions, read_table, write_table
from excel_tool.core.limits import (
    ColumnSuggestion,
    Limits,
    SizeEstimate,
    check,
    estimate,
    load_limits,
    suggest_split_columns,
)
from excel_tool.core.progress import CancelToken, ProgressEvent

STR = pd.StringDtype("pyarrow")
SRC = Path(__file__).resolve().parents[2] / "src" / "excel_tool"


@pytest.fixture(autouse=True)
def _aislar(monkeypatch, tmp_path):
    """Datos del usuario y ejecutable en carpetas vacías de prueba."""
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "xdg"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "local"))
    monkeypatch.setattr(platform_utils, "app_dir", lambda: None)


def _est(filas: int, columnas: int = 10, mb_xlsx: float = 10, mb_csv: float = 10) -> SizeEstimate:
    return SizeEstimate(filas, columnas, (filas + 1) * columnas, mb_xlsx, mb_csv)


# ----------------------------------------------------------------- carga


def test_defaults():
    lim = load_limits()
    assert lim.excel_max_rows == 1_048_576
    assert lim.excel_max_columns == 16_384
    assert lim.excel_comfortable_rows == 300_000
    assert lim.gs_max_cells == 20_000_000
    assert lim.gs_max_import_mb == 100
    assert lim.gs_comfortable_cells == 5_000_000


def test_pisa_junto_al_ejecutable(monkeypatch, tmp_path):
    carpeta = tmp_path / "exe"
    carpeta.mkdir()
    (carpeta / "limites.json").write_text(
        json.dumps({"excel": {"comfortable_rows": 100_000}}), encoding="utf-8"
    )
    monkeypatch.setattr(platform_utils, "app_dir", lambda: carpeta)
    lim = load_limits()
    assert lim.excel_comfortable_rows == 100_000
    assert lim.excel_max_rows == 1_048_576  # lo que no se pisa queda
    # y cambia la recomendación: 150.000 filas ya no es verde
    assert check(_est(150_000), "excel", "xlsx", limits=lim).level == "recomendado"
    assert (
        check(
            _est(150_000), "excel", "xlsx", limits=replace(lim, excel_comfortable_rows=300_000)
        ).level
        == "ok"
    )


def test_pisa_en_datos_de_usuario_y_gana_al_del_ejecutable(monkeypatch, tmp_path):
    carpeta = tmp_path / "exe"
    carpeta.mkdir()
    (carpeta / "limites.json").write_text(
        json.dumps({"excel": {"comfortable_rows": 100_000, "comfortable_mb": 5}}), encoding="utf-8"
    )
    monkeypatch.setattr(platform_utils, "app_dir", lambda: carpeta)
    (platform_utils.user_data_dir() / "limites.json").write_text(
        json.dumps(
            {"excel": {"comfortable_rows": 200_000}, "google_sheets": {"max_import_mb": 10}}
        ),
        encoding="utf-8",
    )
    lim = load_limits()
    assert lim.excel_comfortable_rows == 200_000
    assert lim.excel_comfortable_mb == 5
    assert lim.gs_max_import_mb == 10
    assert check(_est(250_000, mb_xlsx=1), "excel", "xlsx", limits=lim).level == "recomendado"


def test_cambia_la_recomendacion_con_limites_json():
    assert check(_est(400_000), "excel", "xlsx", limits=load_limits()).level == "recomendado"
    (platform_utils.user_data_dir() / "limites.json").write_text(
        json.dumps({"excel": {"comfortable_rows": 500_000, "comfortable_mb": 500}}),
        encoding="utf-8",
    )
    assert check(_est(400_000), "excel", "xlsx", limits=load_limits()).level == "ok"


@pytest.mark.parametrize(
    "contenido",
    [b"{no es json", b"\xff\xfe\x00basura", b"[1, 2]", b"", b'"hola"'],
)
def test_json_danado_usa_defaults(contenido, caplog):
    (platform_utils.user_data_dir() / "limites.json").write_bytes(contenido)
    with caplog.at_level("WARNING"):
        assert load_limits() == _defaults()
    assert caplog.records


def _defaults() -> Limits:
    # con la carpeta de datos todavía sin limites.json
    carpeta = platform_utils.user_data_dir() / "limites.json"
    previo = carpeta.read_bytes() if carpeta.exists() else None
    carpeta.unlink(missing_ok=True)
    try:
        return load_limits()
    finally:
        if previo is not None:
            carpeta.write_bytes(previo)


def test_claves_invalidas_usan_default_de_esa_clave(caplog):
    (platform_utils.user_data_dir() / "limites.json").write_text(
        json.dumps(
            {
                "excel": {
                    "max_rows": "mucho",
                    "max_columns": -5,
                    "comfortable_rows": True,
                    "comfortable_mb": None,
                },
                "google_sheets": {"max_cells": 1.5, "max_import_mb": 20, "otra": 1},
                "otra_seccion": {},
            }
        ),
        encoding="utf-8",
    )
    with caplog.at_level("WARNING"):
        lim = load_limits()
    base = _defaults()
    assert lim.excel_max_rows == base.excel_max_rows
    assert lim.excel_max_columns == base.excel_max_columns
    assert lim.excel_comfortable_rows == base.excel_comfortable_rows
    assert lim.excel_comfortable_mb == base.excel_comfortable_mb
    assert lim.gs_max_cells == base.gs_max_cells
    assert lim.gs_max_import_mb == 20  # la válida sí se aplica
    assert len(caplog.records) >= 5


def test_comodo_no_supera_el_maximo():
    (platform_utils.user_data_dir() / "limites.json").write_text(
        json.dumps({"excel": {"max_rows": 1000}}), encoding="utf-8"
    )
    lim = load_limits()
    assert lim.excel_max_rows == 1000
    assert lim.excel_comfortable_rows == 999


# ----------------------------------------------------------------- check


def test_400k_excel_amarillo_y_solo_guardar_verde():
    lim = load_limits()
    est = _est(400_000, mb_xlsx=20, mb_csv=20)
    r = check(est, "excel", "xlsx", limits=lim)
    assert r.level == "recomendado"
    assert r.suggested_parts == 2
    assert r.suggested_rows_per_part == 200_000
    assert any("300.000" in m and "400.000" in m for m in r.reasons)
    s = check(est, "solo_guardar", "csv_excel", limits=lim)
    assert s.level == "ok"
    assert s.suggested_parts == 1
    assert s.reasons == []


def test_2_3m_excel_rojo_con_partes():
    r = check(_est(2_300_000, mb_xlsx=40), "excel", "xlsx", limits=load_limits())
    assert r.level == "obligatorio"
    assert r.suggested_parts == 8  # ceil(2.300.000 / 300.000)
    assert r.suggested_rows_per_part == 287_500
    assert r.reasons[0] == (
        "Excel acepta hasta 1.048.575 filas de datos por hoja (más el encabezado); "
        "el resultado tiene 2.300.000."
    )


def test_justo_en_el_limite_de_excel():
    lim = load_limits()
    assert check(_est(1_048_575, mb_xlsx=1), "excel", "xlsx", limits=lim).level == "recomendado"
    assert check(_est(1_048_576, mb_xlsx=1), "excel", "xlsx", limits=lim).level == "obligatorio"


def test_partes_por_mb_en_excel():
    r = check(_est(100_000, mb_xlsx=120), "excel", "xlsx", limits=load_limits())
    assert r.level == "recomendado"
    assert r.suggested_parts == 3  # ceil(120 / 50) > ceil(100.000 / 300.000)
    assert r.suggested_rows_per_part == 33_334
    assert any("MB" in m and "120,0" in m for m in r.reasons)


def test_excel_usa_mb_del_formato():
    est = _est(100_000, mb_xlsx=10, mb_csv=200)
    assert check(est, "excel", "xlsx", limits=load_limits()).level == "ok"
    assert check(est, "excel", "csv_excel", limits=load_limits()).level == "recomendado"


def test_excel_demasiadas_columnas():
    r = check(_est(10, columnas=20_000), "excel", "xlsx", limits=load_limits())
    assert r.level == "obligatorio"
    assert any("16.384 columnas" in m for m in r.reasons)


def test_solo_guardar_xlsx_pasa_el_limite_de_hoja():
    lim = load_limits()
    r = check(_est(1_100_000, mb_xlsx=30), "solo_guardar", "xlsx", limits=lim)
    assert r.level == "obligatorio"
    assert r.suggested_parts == 2
    assert check(_est(1_000_000, mb_xlsx=30), "solo_guardar", "xlsx", limits=lim).level == "ok"
    assert check(_est(5_000_000), "solo_guardar", "csv_google", limits=lim).level == "ok"


def test_google_sheets_celdas_comodas_y_duras():
    lim = load_limits()
    verde = check(_est(400_000, columnas=10, mb_csv=10), "google_sheets", "csv_google", limits=lim)
    assert verde.level == "ok"
    amarillo = check(
        _est(600_000, columnas=10, mb_csv=10), "google_sheets", "csv_google", limits=lim
    )
    assert amarillo.level == "recomendado"
    assert amarillo.suggested_parts == 2  # 6.000.010 / 5.000.000
    rojo = check(_est(2_500_000, columnas=10, mb_csv=10), "google_sheets", "csv_google", limits=lim)
    assert rojo.level == "obligatorio"
    assert rojo.suggested_parts == 6  # 25.000.010 / 5.000.000
    assert "20.000.000 celdas" in rojo.reasons[0]
    assert "25.000.010" in rojo.reasons[0]


def test_google_sheets_por_mb():
    lim = load_limits()
    amarillo = check(_est(1000, mb_csv=70), "google_sheets", "csv_google", limits=lim)
    assert amarillo.level == "recomendado"
    assert amarillo.suggested_parts == 2
    rojo = check(_est(1000, mb_csv=230), "google_sheets", "csv_google", limits=lim)
    assert rojo.level == "obligatorio"
    assert rojo.suggested_parts == 5  # ceil(230 / 50)
    assert "100,0 MB" in rojo.reasons[0]


def test_sugerencias_pasan_al_resultado():
    sug = [ColumnSuggestion("Región", 3, 100, True)]
    r = check(_est(400_000), "excel", "xlsx", limits=load_limits(), suggestions=sug)
    assert r.suggested_columns == sug
    assert check(_est(10), "excel", "xlsx", limits=load_limits()).suggested_columns == []


def test_check_carga_limites_si_no_se_pasan():
    assert check(_est(400_000), "excel", "xlsx").level == "recomendado"


# ----------------------------------------------------------------- estimate


@pytest.fixture(scope="module", params=["csv", "xlsx"])
def explorar(request, fixtures_generados):
    return read_table(fixtures_generados / "explorar" / f"para_explorar.{request.param}").df


def test_estimate_conteos(explorar):
    est = estimate(explorar)
    assert est.rows == len(explorar)
    assert est.columns == len(explorar.columns)
    assert est.cells == (est.rows + 1) * est.columns
    assert est.approx_mb_csv > 0
    assert est.approx_mb_xlsx > 0


def test_estimate_csv_cerca_del_real(explorar, tmp_path):
    est = estimate(explorar)
    destino = tmp_path / "a.csv"
    write_table(explorar, destino, WriteOptions("csv_excel"))
    real = destino.stat().st_size / 1024**2
    assert abs(est.approx_mb_csv - real) / real < 0.1


def test_estimate_xlsx_cerca_del_real(tmp_path):
    n = 30_000
    df = pd.DataFrame(
        {
            "a": pd.array([str(i * 7919 % 100003) for i in range(n)], dtype=STR),
            "b": pd.array([f"Cliente número {i % 500}" for i in range(n)], dtype=STR),
            "c": pd.array([["Norte", "Sur", "Este"][i % 3] for i in range(n)], dtype=STR),
        }
    )
    est = estimate(df)
    destino = tmp_path / "a.xlsx"
    write_table(df, destino, WriteOptions("xlsx"))
    real = destino.stat().st_size / 1024**2
    assert abs(est.approx_mb_xlsx - real) / real < 0.35


def test_estimate_usa_muestra_y_no_cambia_la_tabla():
    n = 50_000
    df = pd.DataFrame({"a": pd.array([f"v{i}" for i in range(n)], dtype=STR)})
    chica = estimate(df, sample_rows=500)
    grande = estimate(df, sample_rows=10_000)
    assert chica.rows == grande.rows == n
    assert abs(chica.approx_mb_csv - grande.approx_mb_csv) / grande.approx_mb_csv < 0.1
    assert len(df) == n


def test_estimate_vacio():
    est = estimate(pd.DataFrame({"a": pd.array([], dtype=STR)}))
    assert est.rows == 0
    assert est.approx_mb_csv == 0


# ----------------------------------------------------------------- suggest_split_columns


def test_suggest_explorar(explorar):
    sug = suggest_split_columns(explorar, "excel", "xlsx", limits=load_limits())
    por_nombre = {s.column: s for s in sug}
    assert "Región" in por_nombre
    region = por_nombre["Región"]
    assert region.parts == explorar["Región"].nunique(dropna=False)
    assert region.largest_part_rows == int(explorar["Región"].value_counts(dropna=False).max())
    assert region.fits
    assert "ID" not in por_nombre  # tantos valores como filas
    assert all(2 <= s.parts <= 50 for s in sug)
    # orden: primero las que entran, y entre ellas las de menos partes
    clave = [(not s.fits, s.parts) for s in sug]
    assert clave == sorted(clave)


def test_suggest_max_parts_y_fits():
    df = pd.DataFrame(
        {
            "grande": pd.array(["a"] * 6 + ["b"] * 3, dtype=STR),
            "pareja": pd.array(["x", "y", "z"] * 3, dtype=STR),
            "unica": pd.array(["u"] * 9, dtype=STR),
            "ids": pd.array([str(i) for i in range(9)], dtype=STR),
        }
    )
    lim = Limits(1_048_576, 16_384, 5, 50, 20_000_000, 100, 5_000_000, 50)
    sug = suggest_split_columns(df, "excel", "xlsx", limits=lim, max_parts=5)
    assert [(s.column, s.parts, s.largest_part_rows, s.fits) for s in sug] == [
        ("pareja", 3, 3, True),
        ("grande", 2, 6, False),
    ]


def test_suggest_progreso_y_cancelacion(explorar):
    eventos: list[ProgressEvent] = []
    suggest_split_columns(explorar, "excel", "xlsx", limits=load_limits(), progress=eventos.append)
    assert eventos
    assert eventos[-1].done == eventos[-1].total
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        suggest_split_columns(explorar, "excel", "xlsx", limits=load_limits(), cancel=token)


def test_suggest_google_sheets_usa_celdas():
    df = pd.DataFrame(
        {"r": pd.array(["a", "b"] * 5, dtype=STR), "v": pd.array(list("0123456789"), dtype=STR)}
    )
    lim = Limits(1_048_576, 16_384, 300_000, 50, 20_000_000, 100, 8, 50)
    sug = suggest_split_columns(df, "google_sheets", "csv_google", limits=lim)
    # 8 celdas cómodas / 2 columnas - 1 encabezado = 3 filas por parte
    assert [(s.column, s.fits) for s in sug] == [("v", True), ("r", False)]


# ----------------------------------------------------------------- reglas del proyecto


def test_limites_solo_en_limits_py():
    patron = re.compile(r"1048576|1_048_576|1\.048\.576|20_000_000|20000000|16384|16_384")
    for archivo in SRC.rglob("*.py"):
        if archivo.name == "limits.py":
            continue
        assert not patron.search(archivo.read_text(encoding="utf-8")), archivo
    assert not patron.search((Path(limits.__file__)).read_text(encoding="utf-8"))
