"""split_file: dividir un CSV por partes sin cargarlo entero (F3-05)."""

from __future__ import annotations

import dataclasses
import subprocess
import sys
import textwrap
from pathlib import Path

import pandas as pd
import pytest

from excel_tool.core import limits as limits_mod
from excel_tool.core.errors import ArchivoIlegible, Cancelado, ExcelToolError
from excel_tool.core.io import (
    CsvBlockReader,
    CsvOptions,
    CsvPartWriter,
    ReadOptions,
    WriteOptions,
    read_csv_exact,
    read_table,
    write_table,
)
from excel_tool.core.output import (
    MAX_ARCHIVOS_ABIERTOS,
    NOMBRE_INDICE,
    OutputSpec,
    split_file,
)
from excel_tool.core.progress import CancelToken
from excel_tool.core.split import SplitSpec

RAIZ = Path(__file__).resolve().parents[2]


@pytest.fixture
def raros(fixtures_generados):
    return fixtures_generados / "fase3" / "nombres_raros.csv"


def _spec(tmp_path, split, fmt="csv_excel", dest="solo_guardar", **kw):
    ext = ".xlsx" if fmt == "xlsx" else ".csv"
    return OutputSpec(
        path=tmp_path / ("salida" + ext), format=fmt, destination=dest, split=split, **kw
    )


def _partes_csv(res):
    return [p for p in res.files if p.suffix == ".csv"]


def _juntar(res):
    return pd.concat([read_table(p.path).df for p in res.parts], ignore_index=True)


def _escribir(path: Path, texto: str, encoding="utf-8-sig") -> Path:
    with path.open("w", encoding=encoding, newline="") as f:
        f.write(texto)
    return path


# -------------------------------------------------------------------------------- rows


def test_rows_csv_suma_encabezado_e_indice(raros, tmp_path):
    original = read_table(raros).df
    res = split_file(raros, ReadOptions(), _spec(tmp_path, SplitSpec("rows", rows_per_part=40)))
    assert res.rows_total == len(original) == 99
    assert [p.rows for p in res.parts] == [40, 40, 19]
    assert sum(p.rows for p in res.parts) == res.rows_total
    assert res.folder is not None and res.folder.parent == tmp_path
    assert [p.path.name for p in res.parts] == [
        "salida_parte_01_de_03.csv",
        "salida_parte_02_de_03.csv",
        "salida_parte_03_de_03.csv",
    ]
    assert [(p.first_row, p.last_row) for p in res.parts] == [(1, 40), (41, 80), (81, 99)]
    for p in res.parts:  # el encabezado está en cada parte
        primera = p.path.read_text(encoding="utf-8-sig").splitlines()[0]
        assert primera == "ID;Sucursal;Importe"
    pd.testing.assert_frame_equal(_juntar(res), original, check_dtype=False)
    assert (res.folder / NOMBRE_INDICE).exists() and res.files[-1].name == NOMBRE_INDICE
    indice = read_table(res.folder / NOMBRE_INDICE).df
    assert indice["Filas"].tolist() == ["40", "40", "19"]
    assert sorted(p.name for p in res.folder.iterdir()) == sorted(f.name for f in res.files)


def test_rows_sin_indice_y_carpeta_existente(raros, tmp_path):
    (tmp_path / "salida").mkdir()
    spec = _spec(tmp_path, SplitSpec("rows", rows_per_part=50), write_index=False)
    res = split_file(raros, ReadOptions(), spec)
    assert res.folder == tmp_path / "salida (2)"
    assert not (res.folder / NOMBRE_INDICE).exists()
    assert list((tmp_path / "salida").iterdir()) == []


@pytest.mark.parametrize("fmt", ["csv_google", "csv_standard"])
def test_rows_formatos_csv(raros, tmp_path, fmt):
    res = split_file(
        raros, ReadOptions(), _spec(tmp_path, SplitSpec("rows", rows_per_part=60), fmt)
    )
    sep = ";" if fmt == "csv_excel" else ","
    bruto = res.parts[0].path.read_bytes()
    assert bruto.startswith(b"\xef\xbb\xbf") == (fmt == "csv_standard")
    assert f"ID{sep}Sucursal{sep}Importe".encode() in bruto
    pd.testing.assert_frame_equal(_juntar(res), read_table(raros).df, check_dtype=False)


def test_rows_xlsx(raros, tmp_path):
    spec = _spec(tmp_path, SplitSpec("rows", rows_per_part=40), "xlsx")
    res = split_file(raros, ReadOptions(), spec)
    assert [p.rows for p in res.parts] == [40, 40, 19]
    assert all(p.path.suffix == ".xlsx" and p.sheet == "Datos" for p in res.parts)
    pd.testing.assert_frame_equal(_juntar(res), read_table(raros).df, check_dtype=False)
    assert not list(res.folder.glob("*.part"))


def test_rows_tope_de_excel_manda(raros, tmp_path, monkeypatch):
    lim = dataclasses.replace(limits_mod.load_limits(), excel_max_rows=31)  # 30 filas de datos
    monkeypatch.setattr(limits_mod, "load_limits", lambda: lim)
    spec = _spec(tmp_path, SplitSpec("rows", rows_per_part=1000), "xlsx", dest="excel")
    res = split_file(raros, ReadOptions(), spec)
    assert [p.rows for p in res.parts] == [30, 30, 30, 9]


def test_rows_requiere_cantidad(raros, tmp_path):
    with pytest.raises(ExcelToolError, match="cuántas filas"):
        split_file(raros, ReadOptions(), _spec(tmp_path, SplitSpec("rows")))
    assert list(tmp_path.iterdir()) == []


def test_filas_vacias_del_medio_se_conservan_y_las_del_final_no(tmp_path):
    origen = _escribir(tmp_path / "o.csv", "A;B\r\n1;x\r\n;\r\n3;z\r\n4;w\r\n;\r\n;\r\n")
    sub = tmp_path / "out"
    sub.mkdir()
    res = split_file(
        origen,
        ReadOptions(),
        OutputSpec(sub / "r.csv", "csv_excel", "solo_guardar", SplitSpec("rows", rows_per_part=2)),
    )
    assert res.rows_total == 4 and [p.rows for p in res.parts] == [2, 2]
    exacto = read_csv_exact(res.parts[0].path, CsvOptions(";", "utf-8-sig"))
    assert exacto.shape == (2, 2) and exacto.iloc[1].tolist() == ["", ""]


def test_sin_filas_de_datos_deja_solo_el_encabezado(tmp_path):
    origen = _escribir(tmp_path / "o.csv", "A;B\r\n;\r\n")
    sub = tmp_path / "out"
    sub.mkdir()
    res = split_file(
        origen,
        ReadOptions(),
        OutputSpec(sub / "r.csv", "csv_excel", "solo_guardar", SplitSpec("rows", rows_per_part=2)),
    )
    assert res.rows_total == 0 and res.folder is None
    assert res.files == [sub / "r.csv"]
    assert (sub / "r.csv").read_text(encoding="utf-8-sig").strip() == "A;B"
    assert [p.name for p in sub.iterdir()] == ["r.csv"]


def test_encoding_que_falla_mas_adelante_se_reintenta_con_cp1252(tmp_path):
    cuerpo = "".join(f"{i};abc\r\n" for i in range(60000))
    origen = _escribir(tmp_path / "o.csv", "ID;Nombre\r\n" + cuerpo + "1;Peña\r\n", "cp1252")
    assert origen.stat().st_size > 256 * 1024
    res = split_file(origen, ReadOptions(), _spec(tmp_path, SplitSpec("rows", rows_per_part=40000)))
    assert res.rows_total == 60001
    assert "Peña" in res.parts[-1].path.read_text(encoding="utf-8-sig")


def test_columnas_de_mas_mas_adelante_da_mensaje(tmp_path):
    filas = "".join(f"{i};a\r\n" for i in range(5000))
    origen = _escribir(tmp_path / "o.csv", "A;B\r\n" + filas + "1;2;3\r\n")
    assert origen.stat().st_size > 0
    # Más allá de las primeras 2000 filas: no se puede seguir sin perder columnas.
    with pytest.raises(ArchivoIlegible, match="más columnas"):
        split_file(origen, ReadOptions(), _spec(tmp_path, SplitSpec("rows", rows_per_part=100)))
    assert [p.name for p in tmp_path.iterdir()] == ["o.csv"]


# ------------------------------------------------------------------------------ by_value


def _esperados(fixtures_esperados_dir):
    import json

    return json.loads((fixtures_esperados_dir / "fase3_nombres.json").read_text(encoding="utf-8"))


def test_by_value_nombres_raros(raros, tmp_path):
    esperado = _esperados(raros.parents[2] / "esperados")["filas_por_valor"]
    res = split_file(
        raros, ReadOptions(), _spec(tmp_path, SplitSpec("by_value", value_columns=["Sucursal"]))
    )
    assert res.rows_total == 99
    assert {p.value[0]: p.rows for p in res.parts} == esperado
    assert len(res.parts) == 20
    nombres = [p.path.name for p in res.parts]
    assert len({n.lower() for n in nombres}) == 20  # sin repetir, aunque difieran en mayúsculas
    assert all(p.path.exists() for p in res.parts)
    assert any("vacio" in n for n in nombres)
    # El vacío va al final y las partes quedan ordenadas sin distinguir mayúsculas.
    valores = [p.value[0] for p in res.parts]
    assert valores[-1] == "" and valores[:-1] == sorted(
        valores[:-1], key=lambda v: (v.casefold(), v)
    )
    # Lo escrito, junto, tiene las mismas filas que el original.
    original = read_table(raros).df
    junto = _juntar(res)
    assert len(junto) == 99
    assert sorted(junto["ID"].astype(int)) == sorted(original["ID"].astype(int))
    for p in res.parts:
        df = read_table(p.path).df
        assert set(df["Sucursal"]) == {p.value[0]}
    indice = read_table(res.folder / NOMBRE_INDICE).df
    assert "(vacío)" in indice["Valor"].tolist()


def test_by_value_xlsx_y_subdivision(raros, tmp_path):
    spec = _spec(
        tmp_path, SplitSpec("by_value", value_columns=["Sucursal"], rows_per_part=5), "xlsx"
    )
    res = split_file(raros, ReadOptions(), spec)
    assert sum(p.rows for p in res.parts) == 99 and max(p.rows for p in res.parts) <= 5
    centro = [p for p in res.parts if p.value == ("Centro",)]
    assert [p.rows for p in centro] == [5, 5, 2]
    assert [p.path.name for p in centro][0].endswith("_parte_1_de_3.xlsx")
    assert (centro[0].first_row, centro[0].last_row) == (1, 28) or centro[0].first_row >= 1
    assert not list(res.folder.glob("*.part"))


def test_by_value_dos_columnas(fixtures_generados, tmp_path):
    origen = _escribir(
        tmp_path / "o.csv",
        "R;T;N\r\nN;a;1\r\nS;a;2\r\nN;b;3\r\nN;a;4\r\n",
    )
    sub = tmp_path / "out"
    sub.mkdir()
    spec = OutputSpec(
        sub / "d.csv", "csv_excel", "solo_guardar", SplitSpec("by_value", value_columns=["R", "T"])
    )
    res = split_file(origen, ReadOptions(), spec)
    assert [(p.value, p.rows, p.first_row, p.last_row) for p in res.parts] == [
        (("N", "a"), 2, 1, 4),
        (("N", "b"), 1, 3, 3),
        (("S", "a"), 1, 2, 2),
    ]


def test_by_value_columna_inexistente(raros, tmp_path):
    with pytest.raises(ExcelToolError, match="No encuentro"):
        split_file(
            raros, ReadOptions(), _spec(tmp_path, SplitSpec("by_value", value_columns=["Nada"]))
        )
    assert list(tmp_path.iterdir()) == []


def test_by_value_demasiados_valores(tmp_path):
    n = MAX_ARCHIVOS_ABIERTOS + 10
    origen = _escribir(tmp_path / "o.csv", "K;V\r\n" + "".join(f"k{i};{i}\r\n" for i in range(n)))
    sub = tmp_path / "out"
    sub.mkdir()
    spec = OutputSpec(
        sub / "d.csv", "csv_excel", "solo_guardar", SplitSpec("by_value", value_columns=["K"])
    )
    with pytest.raises(ExcelToolError, match="demasiados valores distintos"):
        split_file(origen, ReadOptions(), spec)
    assert list(sub.iterdir()) == []  # no queda nada: ni partes ni carpeta


# ---------------------------------------------------------------------- rows_keep_groups


@pytest.mark.parametrize("fmt", ["csv_excel", "xlsx"])
def test_keep_groups_no_corta_grupos(raros, tmp_path, fmt):
    spec = _spec(
        tmp_path, SplitSpec("rows_keep_groups", rows_per_part=25, group_column="Sucursal"), fmt
    )
    res = split_file(raros, ReadOptions(), spec)
    assert res.rows_total == 99 and sum(p.rows for p in res.parts) == 99
    dueno: dict[str, int] = {}
    for i, p in enumerate(res.parts):
        for v in set(read_table(p.path).df["Sucursal"]):
            assert dueno.setdefault(v, i) == i  # cada valor en una sola parte
    assert all(p.rows <= 25 for p in res.parts if len(set(read_table(p.path).df["Sucursal"])) > 1)
    original = read_table(raros).df
    junto = _juntar(res)
    assert sorted(junto["ID"].astype(int)) == sorted(original["ID"].astype(int))


def test_keep_groups_avisa_si_un_grupo_pasa_el_limite(raros, tmp_path):
    spec = _spec(tmp_path, SplitSpec("rows_keep_groups", rows_per_part=5, group_column="Sucursal"))
    res = split_file(raros, ReadOptions(), spec)
    assert any("no se puede cortar" in w for w in res.warnings)
    assert sum(p.rows for p in res.parts) == 99


def test_keep_groups_demasiadas_partes(raros, tmp_path, monkeypatch):
    import excel_tool.core.output as out

    monkeypatch.setattr(out, "MAX_ARCHIVOS_ABIERTOS", 3)
    spec = _spec(tmp_path, SplitSpec("rows_keep_groups", rows_per_part=5, group_column="Sucursal"))
    with pytest.raises(ExcelToolError, match="demasiados"):
        split_file(raros, ReadOptions(), spec)
    assert list(tmp_path.iterdir()) == []


# ---------------------------------------------------------------- errores y cancelación


def test_excel_como_entrada_da_el_mensaje(fixtures_generados, tmp_path):
    xlsx = fixtures_generados / "explorar" / "para_explorar.xlsx"
    with pytest.raises(ExcelToolError) as exc:
        split_file(xlsx, ReadOptions(), _spec(tmp_path, SplitSpec("rows", rows_per_part=10)))
    assert exc.value.mensaje == (
        "Por ahora solo se pueden dividir archivos CSV sin abrirlos enteros; "
        "para Excel, abrilo con 'Ver información' y usá Guardar."
    )
    assert list(tmp_path.iterdir()) == []


def test_modo_hojas_e_inexistente(raros, tmp_path):
    with pytest.raises(ExcelToolError, match="hojas"):
        split_file(raros, ReadOptions(), _spec(tmp_path, SplitSpec("sheets", rows_per_part=10)))
    with pytest.raises(ExcelToolError):
        split_file(
            tmp_path / "no_existe.csv",
            ReadOptions(),
            _spec(tmp_path, SplitSpec("rows", rows_per_part=1)),
        )
    with pytest.raises(ExcelToolError, match="Indicá cómo dividir"):
        split_file(raros, ReadOptions(), _spec(tmp_path, None))


@pytest.mark.parametrize("modo", ["rows", "by_value", "rows_keep_groups"])
def test_cancelar_a_mitad_borra_solo_lo_creado(fixtures_generados, tmp_path, modo):
    origen = fixtures_generados / "grandes" / "fase3_400k.csv"
    if not origen.exists():
        pytest.skip("faltan los archivos grandes")
    previo = tmp_path / "mio.txt"
    previo.write_text("no tocar", encoding="utf-8")
    split = {
        "rows": SplitSpec("rows", rows_per_part=100_000),
        "by_value": SplitSpec("by_value", value_columns=["Región"]),
        "rows_keep_groups": SplitSpec(
            "rows_keep_groups", rows_per_part=150_000, group_column="Región"
        ),
    }[modo]
    token = CancelToken()
    eventos = []

    def progreso(ev):
        eventos.append(ev)
        if len(eventos) >= 3:
            token.cancel()

    with pytest.raises(Cancelado):
        split_file(origen, ReadOptions(), _spec(tmp_path, split), progress=progreso, cancel=token)
    assert [p.name for p in tmp_path.iterdir()] == ["mio.txt"]
    assert previo.read_text(encoding="utf-8") == "no tocar"


def test_cancelado_de_entrada(raros, tmp_path):
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        split_file(
            raros, ReadOptions(), _spec(tmp_path, SplitSpec("rows", rows_per_part=10)), cancel=token
        )
    assert list(tmp_path.iterdir()) == []


def test_progreso_en_bytes_del_origen(raros, tmp_path):
    eventos = []
    split_file(
        raros,
        ReadOptions(),
        _spec(tmp_path, SplitSpec("rows", rows_per_part=40)),
        progress=eventos.append,
    )
    assert eventos and all(e.total == raros.stat().st_size for e in eventos)
    assert all("Dividiendo" in e.message for e in eventos)


# ----------------------------------------------------------- piezas de core/io (lectura/escritura)


def test_csv_part_writer_igual_que_write_table(raros, tmp_path):
    df = read_table(raros).df
    for formato in ("csv_excel", "csv_google", "csv_standard"):
        a, b = tmp_path / f"a_{formato}.csv", tmp_path / f"b_{formato}.csv"
        write_table(df, a, WriteOptions(format=formato))
        with CsvPartWriter(b, formato, list(df.columns)) as w:
            w.write(df.iloc[:30])
            w.write(df.iloc[30:])
        assert w.rows == len(df)
        assert a.read_bytes() == b.read_bytes()


def test_estilos_coinciden_con_el_escritor():
    from excel_tool.core.io.csv import CSV_ESTILOS
    from excel_tool.core.io.writer import _CSV_FORMATOS

    assert {k: v[:2] for k, v in _CSV_FORMATOS.items()} == CSV_ESTILOS


def test_block_reader_igual_que_read_table(raros, fixtures_generados):
    esperado = read_table(raros)
    with CsvBlockReader(raros) as lector:
        assert lector.info.columns == list(esperado.df.columns)
        assert lector.info.csv_options == esperado.csv_options
        bloques = list(lector)
        assert lector.rows == len(esperado.df)
    pd.testing.assert_frame_equal(
        pd.concat(bloques, ignore_index=True), esperado.df, check_dtype=False
    )
    # Con un xlsx no aplica: el lector es solo de CSV.
    assert CsvOptions(";", "utf-8-sig") == esperado.csv_options


def test_block_reader_varios_bloques_y_encabezado_en_otra_fila(tmp_path):
    cuerpo = "".join(f"{i};v{i}\r\n" for i in range(250_000))
    origen = _escribir(tmp_path / "o.csv", "Reporte de prueba\r\n\r\nID;V\r\n" + cuerpo)
    with CsvBlockReader(origen, ReadOptions(header_row=2)) as lector:
        assert lector.info.columns == ["ID", "V"]
        tamanios = [len(b) for b in lector]
    assert sum(tamanios) == 250_000 and len(tamanios) >= 3 and max(tamanios) <= 100_000


# ------------------------------------------------------------------------ archivos grandes


def test_grande_1600k_roundtrip(fixtures_grandes, tmp_path):
    origen = fixtures_grandes / "fase3_dividir_1600k.csv"
    if not origen.exists():
        pytest.skip("faltan los archivos grandes")
    res = split_file(
        origen, ReadOptions(), _spec(tmp_path, SplitSpec("rows", rows_per_part=300_000))
    )
    assert res.rows_total == 1_600_000
    assert [p.rows for p in res.parts] == [300_000] * 5 + [100_000]
    assert sum(p.rows for p in res.parts) == res.rows_total
    original = read_table(origen).df
    pd.testing.assert_frame_equal(_juntar(res), original, check_dtype=False)
    for p in res.parts:
        assert p.path.read_text(encoding="utf-8-sig").splitlines()[0] == "ID;Región;Cliente;Importe"


def test_grande_1600k_por_region(fixtures_grandes, tmp_path):
    origen = fixtures_grandes / "fase3_dividir_1600k.csv"
    if not origen.exists():
        pytest.skip("faltan los archivos grandes")
    esperado = {
        "Norte": 1_200_000,
        "Sur": 100_000,
        "Centro": 100_000,
        "Cuyo": 100_000,
        "Litoral": 100_000,
    }
    spec = _spec(tmp_path, SplitSpec("by_value", value_columns=["Región"], rows_per_part=500_000))
    res = split_file(origen, ReadOptions(), spec)
    assert res.rows_total == 1_600_000
    por = {}
    for p in res.parts:
        por[p.value[0]] = por.get(p.value[0], 0) + p.rows
        assert p.rows <= 500_000
    assert por == esperado
    assert [p.rows for p in res.parts if p.value == ("Norte",)] == [500_000, 500_000, 200_000]


def test_memoria_acotada_al_dividir_2M(fixtures_grandes, tmp_path):
    """El pico del proceso (que ya incluye pandas) queda muy por debajo de cargar el archivo."""
    origen = fixtures_grandes / "rendimiento_2M.csv"
    if not Path("/proc/self/status").exists():
        pytest.skip("solo se mide con /proc (Linux)")
    if not origen.exists():
        pytest.skip("faltan los archivos grandes")
    codigo = textwrap.dedent(
        """
        import sys
        from pathlib import Path
        from excel_tool.core.io import ReadOptions
        from excel_tool.core.output import OutputSpec, split_file
        from excel_tool.core.split import SplitSpec
        out = Path(sys.argv[2])
        res = split_file(Path(sys.argv[1]), ReadOptions(), OutputSpec(
            out / "r.csv", "csv_excel", "solo_guardar", SplitSpec("rows", rows_per_part=300000)))
        pico = 0.0
        for linea in Path("/proc/self/status").read_text().splitlines():
            if linea.startswith("VmHWM:"):  # pico real del proceso (ru_maxrss hereda el del padre)
                pico = int(linea.split()[1]) / 1024
        print(res.rows_total, len(res.parts), pico)
        """
    )
    r = subprocess.run(
        [sys.executable, "-I", "-c", codigo, str(origen), str(tmp_path)],
        capture_output=True,
        text=True,
        timeout=300,
        check=False,
    )
    assert r.returncode == 0, r.stderr
    filas, partes, pico = r.stdout.split()
    assert (int(filas), int(partes)) == (2_000_000, 7)
    assert float(pico) < 300  # cargarlo entero llega a ~570 MB
