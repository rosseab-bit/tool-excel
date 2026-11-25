"""read_table / list_sheets con progress y cancel (ADR-027)."""

from __future__ import annotations

import time
from pathlib import Path

import pandas as pd
import pytest

from excel_tool.core.errors import Cancelado
from excel_tool.core.io import ReadOptions, list_sheets, read_table
from excel_tool.core.progress import CancelToken, ProgressEvent

RENDIMIENTO = (
    Path(__file__).parents[1] / "fixtures" / "generados" / "grandes" / "rendimiento_2M.csv"
)


@pytest.fixture
def csv_grande(tmp_path: Path) -> Path:
    if RENDIMIENTO.exists():
        return RENDIMIENTO
    ruta = tmp_path / "grande.csv"
    n = 450_000
    pd.DataFrame({"id": range(n), "texto": ["abc"] * n, "monto": ["1,5"] * n}).to_csv(
        ruta, sep=";", index=False
    )
    return ruta


def test_cancelar_a_mitad_de_csv_grande(csv_grande: Path) -> None:
    token = CancelToken()
    eventos: list[ProgressEvent] = []
    marca: list[float] = []

    def cb(e: ProgressEvent) -> None:
        eventos.append(e)
        if e.done > 0 and not marca:
            marca.append(time.monotonic())
            token.cancel()

    with pytest.raises(Cancelado):
        read_table(csv_grande, progress=cb, cancel=token)
    assert time.monotonic() - marca[0] <= 2.0


def test_progreso_monotono_y_termina(tmp_path: Path) -> None:
    ruta = tmp_path / "p.csv"
    n = 250_000
    pd.DataFrame({"a": range(n), "b": ["x"] * n}).to_csv(ruta, sep=";", index=False)
    eventos: list[ProgressEvent] = []
    r = read_table(ruta, progress=eventos.append)
    assert len(r.df) == n
    hechos = [e.done for e in eventos]
    assert hechos == sorted(hechos)
    assert all(e.message == "Leyendo el archivo…" for e in eventos)
    assert eventos[-1].done == eventos[-1].total
    assert eventos[-1].total == ruta.stat().st_size


def test_sin_progress_ni_cancel_es_identico(tmp_path: Path) -> None:
    ruta = tmp_path / "i.csv"
    n = 120_000
    pd.DataFrame({"a": range(n), "b": ["ñ"] * n}).to_csv(ruta, sep=";", index=False)
    base = read_table(ruta)
    con = read_table(ruta, progress=lambda e: None, cancel=CancelToken())
    pd.testing.assert_frame_equal(base.df, con.df)
    assert base.header_row == con.header_row


def test_csv_cancelado_antes_de_leer(tmp_path: Path) -> None:
    ruta = tmp_path / "c.csv"
    ruta.write_bytes(b"a;b\n1;2\n")
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        read_table(ruta, cancel=token)
    with pytest.raises(Cancelado):
        list_sheets(ruta, cancel=token)


def test_csv_pequeno_con_nrows_acepta_parametros(tmp_path: Path) -> None:
    ruta = tmp_path / "n.csv"
    ruta.write_bytes(b"a;b\n1;2\n3;4\n")
    r = read_table(ruta, ReadOptions(nrows=1), progress=lambda e: None, cancel=CancelToken())
    assert len(r.df) == 1


@pytest.fixture
def xlsx(tmp_path: Path) -> Path:
    ruta = tmp_path / "x.xlsx"
    pd.DataFrame({"a": ["1", "2"], "b": ["x", "y"]}).to_excel(ruta, index=False)
    return ruta


def test_excel_cancelado_antes_de_leer(xlsx: Path) -> None:
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        read_table(xlsx, cancel=token)
    with pytest.raises(Cancelado):
        list_sheets(xlsx, cancel=token)


def test_excel_cancelado_despues_de_leer(xlsx: Path) -> None:
    token = CancelToken()
    # El evento inicial sale antes de leer; el segundo chequeo es el de después de leer.
    with pytest.raises(Cancelado):
        read_table(xlsx, progress=lambda e: token.cancel(), cancel=token)


def test_excel_emite_evento_indeterminado(xlsx: Path) -> None:
    eventos: list[ProgressEvent] = []
    r = read_table(xlsx, progress=eventos.append, cancel=CancelToken())
    assert len(r.df) == 2
    assert eventos[0] == ProgressEvent("Leyendo la hoja…", 0, None)
    assert all(e.total is None for e in eventos)
    assert [h.name for h in list_sheets(xlsx, progress=eventos.append)] == ["Sheet1"]
