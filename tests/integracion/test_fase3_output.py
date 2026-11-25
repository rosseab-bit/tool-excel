"""Fase 3, core de punta a punta: `save_result` y `split_file` con el core real.

Se guarda, se vuelve a leer TODO con `read_table` (nada de mirar internas) y se compara con el
original. Los números esperados salen de los datos (200 filas de `para_explorar.csv`, 99 de
`nombres_raros.csv`, y los JSON de `tests/fixtures/esperados/`), no del código que se prueba.
El formato del xlsx (tipos reales, encabezado congelado, filtro) se verifica con openpyxl
(extra `dev`).
"""

from __future__ import annotations

import datetime as dt
import json
import re
from collections import Counter
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from excel_tool import platform_utils
from excel_tool.core.errors import Cancelado, ExcelToolError
from excel_tool.core.io import ReadOptions, read_table
from excel_tool.core.output import NOMBRE_INDICE, OutputSpec, save_result, split_file
from excel_tool.core.profile import profile_table
from excel_tool.core.progress import CancelToken
from excel_tool.core.split import SplitSpec

FORMATOS = ["xlsx", "csv_excel", "csv_google", "csv_standard"]
EXT = {"xlsx": ".xlsx", "csv_excel": ".csv", "csv_google": ".csv", "csv_standard": ".csv"}
RESERVADOS = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
PROHIBIDOS = set('<>:"/\\|?*')
EXCEL_TOPE_FILAS = 1_048_575  # 1.048.576 contando el encabezado (a mano, PLAN §1.2)


@pytest.fixture(autouse=True)
def _datos_de_usuario_aislados(tmp_path, monkeypatch):
    carpeta = tmp_path / "datos_usuario"
    carpeta.mkdir()
    monkeypatch.setattr(platform_utils, "user_data_dir", lambda: carpeta)
    return carpeta


# ------------------------------------------------------------------------------ datos


@pytest.fixture(scope="module")
def explorar(fixtures_generados):
    lectura = read_table(fixtures_generados / "explorar" / "para_explorar.csv")
    perfil = profile_table(lectura)
    tipos = {c.name: c.detected_type for c in perfil.columns_profile}
    return lectura.df, tipos


@pytest.fixture(scope="module")
def nombres(fixtures_generados):
    return read_table(fixtures_generados / "fase3" / "nombres_raros.csv").df


def _canon(df: pd.DataFrame, tipos: dict[str, str] | None) -> list[tuple]:
    """Filas comparables entre lo original y lo releído.

    Con tipos, el decimal sale con punto en los CSV de Google y xlsx, y la fecha con hora lleva
    segundos (convención de lectura de la Fase 1): se normalizan las dos puntas."""
    cols = list(df.columns)
    filas = []
    for fila in df.itertuples(index=False, name=None):
        f = []
        for col, v in zip(cols, fila, strict=True):
            v = "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
            t = (tipos or {}).get(col)
            if t in ("decimal", "entero") and v:
                v = f"{float(v.replace(',', '.')):.6f}"
            elif t == "fecha_hora" and len(v) == 16:
                v += ":00"
            f.append(v)
        if any(
            f
        ):  # el lector descarta las filas vacías del final de una hoja (contrato de lectura)
            filas.append(tuple(f))
    return filas


def _nombre_valido_windows(nombre: str) -> None:
    assert not (set(nombre) & PROHIBIDOS), nombre
    assert nombre == nombre.rstrip(" ."), nombre
    assert nombre.split(".")[0].upper() not in RESERVADOS, nombre
    assert 0 < len(nombre) <= 255


def _leer_partes(res, formato: str, hojas: bool = False) -> list[pd.DataFrame]:
    """Relee cada parte (no el índice) con `read_table`, en el orden de `res.parts`."""
    out = []
    for p in res.parts:
        opc = ReadOptions(sheet=p.sheet) if (p.sheet and hojas) else None
        out.append(read_table(p.path, opc).df if opc else read_table(p.path).df)
    return out


def _verificar_partes(res, original, tipos, formato, *, ordenado: bool, hojas: bool = False):
    dfs = _leer_partes(res, formato, hojas)
    assert res.rows_total == len(original) == sum(p.rows for p in res.parts)
    for d in dfs:  # encabezado en cada parte
        assert list(d.columns) == list(original.columns)
    juntas = [f for d in dfs for f in _canon(d, tipos)]
    esperadas = _canon(original, tipos)
    if ordenado:
        assert juntas == esperadas
    else:
        assert Counter(juntas) == Counter(esperadas)
    return dfs


def _verificar_indice(res, carpeta: Path):
    idx = read_table(carpeta / NOMBRE_INDICE).df
    assert list(idx.columns) == ["Archivo", "Hoja", "Filas", "Desde fila", "Hasta fila", "Valor"]
    assert idx["Archivo"].tolist() == [p.path.name for p in res.parts]
    assert [int(x) for x in idx["Filas"]] == [p.rows for p in res.parts]
    assert sum(int(x) for x in idx["Filas"]) == res.rows_total
    assert [int(x) for x in idx["Desde fila"]] == [p.first_row for p in res.parts]
    assert [int(x) for x in idx["Hasta fila"]] == [p.last_row for p in res.parts]
    return idx


def _verificar_xlsx_formato(path: Path, tipos: dict[str, str], hoja: str | None = None):
    libro = openpyxl.load_workbook(path)  # no read_only: hace falta freeze_panes / auto_filter
    ws = libro[hoja] if hoja else libro.worksheets[0]
    assert ws.freeze_panes == "A2", f"{path.name}: encabezado no congelado"
    assert ws.auto_filter.ref, f"{path.name}: sin filtro"
    assert ws.auto_filter.ref.startswith("A1:")
    cols = [c.value for c in ws[1]]
    primera = {c: ws.cell(row=2, column=i + 1) for i, c in enumerate(cols)}
    for nombre, t in tipos.items():
        if nombre not in primera:
            continue
        celda = primera[nombre]
        if t in ("entero", "decimal"):
            assert isinstance(celda.value, (int, float)), (path.name, nombre, celda.value)
        elif t in ("fecha", "fecha_hora"):
            assert isinstance(celda.value, (dt.datetime, dt.date)), (path.name, nombre, celda.value)
        elif nombre == "Código":
            assert celda.value == "02055" or isinstance(celda.value, str), (nombre, celda.value)
    libro.close()


def _spec(tmp_path, formato, destino="excel", split=None, tipos=None, nombre="salida"):
    return OutputSpec(
        path=tmp_path / f"{nombre}{EXT[formato]}",
        format=formato,
        destination=destino,
        split=split,
        column_types=tipos,
    )


# ------------------------------------------------------------------------------ un archivo


@pytest.mark.parametrize("formato", FORMATOS)
@pytest.mark.parametrize("con_tipos", [False, True], ids=["texto", "tipos"])
def test_un_archivo(tmp_path, explorar, formato, con_tipos):
    df, tipos = explorar
    t = tipos if con_tipos else None
    res = save_result(df, _spec(tmp_path, formato, tipos=t))
    assert res.folder is None
    assert [p.name for p in tmp_path.iterdir() if p.is_file()] == [f"salida{EXT[formato]}"]
    assert res.files == [res.parts[0].path]
    _verificar_partes(res, df, t, formato, ordenado=True)
    if not con_tipos:  # sin tipos todo vuelve idéntico, letra por letra
        assert read_table(res.files[0]).df.equals(df)
    if formato == "xlsx":
        _verificar_xlsx_formato(res.files[0], t or {})


def test_un_archivo_xlsx_deja_numeros_y_fechas_reales(tmp_path, explorar):
    df, tipos = explorar
    res = save_result(df, _spec(tmp_path, "xlsx", tipos=tipos))
    ws = openpyxl.load_workbook(res.files[0]).worksheets[0]
    cols = [c.value for c in ws[1]]
    fila2 = {c: ws.cell(row=2, column=i + 1) for i, c in enumerate(cols)}
    assert fila2["ID"].value == 1 and isinstance(fila2["ID"].value, int)
    assert fila2["Importe"].value == pytest.approx(3322.83)
    assert fila2["Fecha"].value == dt.datetime(2024, 7, 3)
    assert fila2["Fecha"].is_date
    assert fila2["Fecha y hora"].value == dt.datetime(2024, 7, 3, 13, 5)
    assert fila2["Código"].value == "02055"  # texto: conserva el cero
    assert fila2["Código"].data_type == "s"
    # la fila 2 de Excel es la fila 1 de datos; 200 filas + encabezado
    assert ws.max_row == 201


def test_xlsx_texto_que_empieza_con_igual_queda_texto(tmp_path):
    df = pd.DataFrame(
        {
            "A": ["=1+1", "=SUMA(A1;A2)", "+5", "-3", "@x", "normal"],
            "B": ["1", "2", "3", "4", "5", "6"],
        },
        dtype=object,
    ).astype("string[pyarrow]")
    res = save_result(df, _spec(tmp_path, "xlsx"))
    ws = openpyxl.load_workbook(res.files[0]).worksheets[0]
    assert [ws.cell(row=i, column=1).value for i in range(2, 8)] == df["A"].tolist()
    assert all(ws.cell(row=i, column=1).data_type == "s" for i in range(2, 8))
    assert read_table(res.files[0]).df["A"].tolist() == df["A"].tolist()


def test_destino_existente_se_reemplaza_en_un_archivo_suelto(tmp_path, explorar):
    df, _ = explorar
    destino = tmp_path / "salida.xlsx"
    destino.write_bytes(b"VIEJO")
    save_result(df, _spec(tmp_path, "xlsx"))
    assert read_table(destino).df.equals(df)


# ------------------------------------------------------------------------------ partes por filas


@pytest.mark.parametrize("formato", FORMATOS)
@pytest.mark.parametrize("con_tipos", [False, True], ids=["texto", "tipos"])
def test_partes_por_filas(tmp_path, explorar, formato, con_tipos):
    df, tipos = explorar
    t = tipos if con_tipos else None
    res = save_result(
        df, _spec(tmp_path, formato, split=SplitSpec("rows", rows_per_part=70), tipos=t)
    )
    assert res.folder == tmp_path / "salida"
    assert [p.rows for p in res.parts] == [70, 70, 60]
    assert [(p.first_row, p.last_row) for p in res.parts] == [(1, 70), (71, 140), (141, 200)]
    assert [p.path.name for p in res.parts] == [
        f"salida_parte_{n:02d}_de_03{EXT[formato]}" for n in (1, 2, 3)
    ]
    assert sorted(x.name for x in res.folder.iterdir()) == sorted(
        [p.path.name for p in res.parts] + [NOMBRE_INDICE]
    )
    assert set(res.files) == {*(p.path for p in res.parts), res.folder / NOMBRE_INDICE}
    _verificar_partes(res, df, t, formato, ordenado=True)
    _verificar_indice(res, res.folder)
    for p in res.parts:
        _nombre_valido_windows(p.path.name)
        if formato == "xlsx":
            _verificar_xlsx_formato(p.path, t or {})


def test_partes_por_filas_sin_indice(tmp_path, explorar):
    df, _ = explorar
    spec = OutputSpec(
        tmp_path / "s.csv",
        "csv_excel",
        "excel",
        SplitSpec("rows", rows_per_part=100),
        write_index=False,
    )
    res = save_result(df, spec)
    assert sorted(p.name for p in res.folder.iterdir()) == sorted(p.path.name for p in res.parts)
    assert not (res.folder / NOMBRE_INDICE).exists()


def test_partes_por_filas_cae_en_el_limite_duro_de_excel(tmp_path, explorar):
    """Con un límite de Excel bajo en limites.json, ninguna parte lo supera."""
    df, _ = explorar
    datos = platform_utils.user_data_dir()
    (datos / "limites.json").write_text(
        json.dumps({"excel": {"max_rows": 51, "comfortable_rows": 30}}), encoding="utf-8"
    )
    res = save_result(df, _spec(tmp_path, "xlsx", split=SplitSpec("rows", rows_per_part=1000)))
    assert max(p.rows for p in res.parts) <= 50
    assert sum(p.rows for p in res.parts) == 200
    assert [len(d) for d in _verificar_partes(res, df, None, "xlsx", ordenado=True)] == [
        p.rows for p in res.parts
    ]


# ------------------------------------------------------------------------------ sin cortar grupos


@pytest.mark.parametrize("formato", FORMATOS)
def test_sin_cortar_grupos(tmp_path, explorar, formato):
    df, tipos = explorar
    spec = SplitSpec("rows_keep_groups", rows_per_part=80, group_column="Región")
    res = save_result(df, _spec(tmp_path, formato, split=spec, tipos=tipos))
    dfs = _verificar_partes(res, df, tipos, formato, ordenado=False)
    valores = [set(d["Región"]) for d in dfs]
    for i, a in enumerate(valores):  # un valor de Región nunca está en dos partes
        for b in valores[i + 1 :]:
            assert not (a & b), (a, b)
    assert all(p.rows <= 80 for p in res.parts) or res.warnings
    _verificar_indice(res, res.folder)
    for p in res.parts:
        _nombre_valido_windows(p.path.name)


# ------------------------------------------------------------------------------ por valor


@pytest.mark.parametrize("formato", FORMATOS)
def test_por_valor(tmp_path, explorar, formato):
    df, tipos = explorar
    cuentas = Counter(df["Región"])
    spec = SplitSpec("by_value", value_columns=["Región"])
    res = save_result(df, _spec(tmp_path, formato, split=spec, tipos=tipos))
    assert len(res.parts) == len(cuentas) == 7
    assert {p.value[0]: p.rows for p in res.parts} == dict(cuentas)
    dfs = _verificar_partes(res, df, tipos, formato, ordenado=False)
    for p, d in zip(res.parts, dfs, strict=True):
        assert set(d["Región"]) <= {
            p.value[0]
        }  # la parte del vacío (filas vacías) se relee sin filas
        _nombre_valido_windows(p.path.name)
    idx = _verificar_indice(res, res.folder)
    assert set(idx["Valor"]) == {
        v if v != "" else "(vacío)" for v in cuentas
    }  # el vacío se muestra como "(vacío)"


@pytest.mark.parametrize("formato", FORMATOS)
def test_por_valor_con_subdivision(tmp_path, explorar, formato):
    df, tipos = explorar
    cuentas = Counter(df["Región"])
    spec = SplitSpec("by_value", value_columns=["Región"], rows_per_part=20)
    res = save_result(df, _spec(tmp_path, formato, split=spec, tipos=tipos))
    esperadas = sum(-(-n // 20) for n in cuentas.values())
    assert len(res.parts) == esperadas
    assert max(p.rows for p in res.parts) <= 20
    _verificar_partes(res, df, tipos, formato, ordenado=False)
    _verificar_indice(res, res.folder)
    assert len({p.path.name.lower() for p in res.parts}) == len(res.parts)


@pytest.mark.parametrize("formato", FORMATOS)
def test_por_valor_con_nombres_raros(tmp_path, nombres, formato):
    esperado = json.loads(
        (Path(__file__).parent.parent / "fixtures" / "esperados" / "fase3_nombres.json").read_text(
            encoding="utf-8"
        )
    )
    spec = SplitSpec("by_value", value_columns=["Sucursal"])
    res = save_result(nombres, _spec(tmp_path, formato, split=spec, nombre="raros"))
    assert len(res.parts) == esperado["valores_distintos"]
    assert {p.value[0]: p.rows for p in res.parts} == esperado["filas_por_valor"]
    nombres_arch = [p.path.name for p in res.parts]
    assert len({n.lower() for n in nombres_arch}) == len(nombres_arch)
    for p in res.parts:
        _nombre_valido_windows(p.path.name)
        assert p.path.parent == res.folder
        assert len(str(p.path)) < 260
    _verificar_partes(res, nombres, None, formato, ordenado=False)
    _verificar_indice(res, res.folder)
    on_disk = sorted(x.name for x in res.folder.iterdir())
    assert on_disk == sorted(nombres_arch + [NOMBRE_INDICE])


# ------------------------------------------------------------------------------ varias hojas


@pytest.mark.parametrize("modo", ["rows", "by_value"])
@pytest.mark.parametrize("con_tipos", [False, True], ids=["texto", "tipos"])
def test_varias_hojas(tmp_path, explorar, modo, con_tipos):
    df, tipos = explorar
    t = tipos if con_tipos else None
    spec = (
        SplitSpec("sheets", rows_per_part=80)
        if modo == "rows"
        else SplitSpec("sheets", value_columns=["Región"])
    )
    res = save_result(df, _spec(tmp_path, "xlsx", split=spec, tipos=t))
    assert res.folder is None and res.files == [tmp_path / "salida.xlsx"]
    assert all(p.path == res.files[0] for p in res.parts)
    hojas = [p.sheet for p in res.parts]
    assert len(set(hojas)) == len(hojas)
    assert [h for h in openpyxl.load_workbook(res.files[0]).sheetnames] == hojas
    for h in hojas:
        assert len(h) <= 31 and not (set(h) & set("[]:*?/\\"))
    dfs = _verificar_partes(res, df, t, "xlsx", ordenado=True, hojas=True)
    if modo == "rows":
        assert [len(d) for d in dfs] == [80, 80, 40]
    for h in hojas:
        _verificar_xlsx_formato(res.files[0], t or {}, h)


def test_varias_hojas_con_nombres_raros(tmp_path, nombres):
    res = save_result(
        nombres,
        _spec(tmp_path, "xlsx", split=SplitSpec("sheets", rows_per_part=5), nombre="raros"),
    )
    libro = openpyxl.load_workbook(res.files[0])
    assert len(libro.sheetnames) == 20  # 99 filas de a 5
    assert len({n.lower() for n in libro.sheetnames}) == 20
    for h in libro.sheetnames:
        assert 0 < len(h) <= 31 and not (set(h) & set("[]:*?/\\"))
    _verificar_partes(res, nombres, None, "xlsx", ordenado=False, hojas=True)


@pytest.mark.parametrize("formato", ["csv_excel", "csv_google", "csv_standard"])
def test_varias_hojas_en_csv_es_error_amigable(tmp_path, explorar, formato):
    df, _ = explorar
    with pytest.raises(ExcelToolError) as exc:
        save_result(df, _spec(tmp_path, formato, split=SplitSpec("sheets", rows_per_part=80)))
    assert "hojas" in str(exc.value)
    assert list(tmp_path.iterdir()) == [tmp_path / "datos_usuario"]


def test_varias_hojas_para_google_sheets_es_error_amigable(tmp_path, explorar):
    df, _ = explorar
    with pytest.raises(ExcelToolError) as exc:
        save_result(
            df, _spec(tmp_path, "xlsx", "google_sheets", SplitSpec("sheets", rows_per_part=80))
        )
    assert "Google Sheets" in str(exc.value)
    assert not (tmp_path / "salida.xlsx").exists()


# ------------------------------------------------------------------------------ carpeta existente


@pytest.mark.parametrize("formato", ["xlsx", "csv_excel"])
def test_carpeta_existente_usa_numero_2_y_no_toca_lo_anterior(tmp_path, explorar, formato):
    df, _ = explorar
    spec = _spec(tmp_path, formato, split=SplitSpec("rows", rows_per_part=100))
    primera = save_result(df, spec)
    antes = {p.name: p.read_bytes() for p in primera.folder.iterdir()}
    segunda = save_result(df, spec)
    tercera = save_result(df, spec)
    assert primera.folder == tmp_path / "salida"
    assert segunda.folder == tmp_path / "salida (2)"
    assert tercera.folder == tmp_path / "salida (3)"
    assert {p.name: p.read_bytes() for p in primera.folder.iterdir()} == antes
    _verificar_partes(segunda, df, None, formato, ordenado=True)
    _verificar_indice(segunda, segunda.folder)


def test_carpeta_existente_con_archivos_ajenos_no_se_toca(tmp_path, explorar):
    df, _ = explorar
    carpeta = tmp_path / "salida"
    carpeta.mkdir()
    (carpeta / "mis_notas.txt").write_text("importante", encoding="utf-8")
    res = save_result(df, _spec(tmp_path, "csv_excel", split=SplitSpec("rows", rows_per_part=100)))
    assert res.folder == tmp_path / "salida (2)"
    assert [p.name for p in carpeta.iterdir()] == ["mis_notas.txt"]


# ------------------------------------------------------------------------------ cancelar


class _CancelarAl:
    """Progress que cancela al recibir el evento número `n`."""

    def __init__(self, token: CancelToken, n: int):
        self.token, self.n, self.eventos = token, n, 0

    def __call__(self, ev) -> None:
        self.eventos += 1
        if self.eventos == self.n:
            self.token.cancel()


def _df_mediano(filas: int = 60_000) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "ID": [str(i) for i in range(filas)],
            "Región": [("Norte", "Sur", "Centro")[i % 3] for i in range(filas)],
            "Texto": [f"descripcion numero {i}" for i in range(filas)],
            "Importe": [f"{i % 997},{i % 100:02d}" for i in range(filas)],
        },
        dtype="string[pyarrow]",
    )


def _foto(carpeta: Path) -> dict[str, bytes | None]:
    return {
        str(p.relative_to(carpeta)): (p.read_bytes() if p.is_file() else None)
        for p in sorted(carpeta.rglob("*"))
    }


@pytest.mark.parametrize("formato", ["xlsx", "csv_excel"])
@pytest.mark.parametrize("modo", ["rows", "by_value", "rows_keep_groups"])
def test_cancelar_a_mitad_en_partes_no_deja_nada(tmp_path, formato, modo):
    df = _df_mediano()
    spec = (
        SplitSpec("rows", rows_per_part=10_000)
        if modo == "rows"
        else SplitSpec("by_value", value_columns=["Región"], rows_per_part=10_000)
        if modo == "by_value"
        else SplitSpec("rows_keep_groups", rows_per_part=25_000, group_column="Región")
    )
    ajeno = tmp_path / "ajeno.txt"
    ajeno.write_text("no me toques", encoding="utf-8")
    previo = tmp_path / "salida"
    previo.mkdir()
    (previo / "viejo.csv").write_text("a;b\n1;2\n", encoding="utf-8")
    antes = _foto(tmp_path)

    token = CancelToken()
    # Se cancela cuando ya hay partes escritas: el evento depende del formato (xlsx es más lento).
    cancelador = _CancelarAl(token, 2)

    visto = {"partes": 0}

    def progreso(ev):
        cancelador(ev)
        carpetas = [p for p in tmp_path.iterdir() if p.is_dir() and p.name.startswith("salida (")]
        for c in carpetas:
            visto["partes"] = max(visto["partes"], len(list(c.iterdir())))

    with pytest.raises(Cancelado):
        save_result(
            df,
            _spec(tmp_path, formato, split=spec),
            progress=progreso,
            cancel=token,
        )
    assert cancelador.eventos >= 2
    assert _foto(tmp_path) == antes, "quedó algo o se tocó lo que existía"
    assert not (tmp_path / "salida (2)").exists()


@pytest.mark.parametrize("formato", ["xlsx", "csv_excel"])
def test_cancelar_un_archivo_no_toca_el_que_existia(tmp_path, formato):
    df = _df_mediano()
    destino = tmp_path / f"salida{EXT[formato]}"
    destino.write_bytes(b"ORIGINAL")
    token = CancelToken()
    with pytest.raises(Cancelado):
        save_result(df, _spec(tmp_path, formato), progress=_CancelarAl(token, 1), cancel=token)
    assert destino.read_bytes() == b"ORIGINAL"
    assert sorted(p.name for p in tmp_path.iterdir() if p.name != "datos_usuario") == [destino.name]


def test_cancelar_un_archivo_nuevo_no_deja_ni_el_temporal(tmp_path):
    df = _df_mediano()
    token = CancelToken()
    with pytest.raises(Cancelado):
        save_result(df, _spec(tmp_path, "xlsx"), progress=_CancelarAl(token, 2), cancel=token)
    assert [p.name for p in tmp_path.iterdir()] == ["datos_usuario"]


def test_cancelar_varias_hojas_no_deja_nada(tmp_path):
    df = _df_mediano()
    destino = tmp_path / "salida.xlsx"
    destino.write_bytes(b"ORIGINAL")
    token = CancelToken()
    with pytest.raises(Cancelado):
        save_result(
            df,
            _spec(tmp_path, "xlsx", split=SplitSpec("sheets", rows_per_part=10_000)),
            progress=_CancelarAl(token, 2),
            cancel=token,
        )
    assert destino.read_bytes() == b"ORIGINAL"
    assert sorted(p.name for p in tmp_path.iterdir()) == ["datos_usuario", "salida.xlsx"]


def test_cancelar_antes_de_empezar_no_deja_nada(tmp_path, explorar):
    df, _ = explorar
    token = CancelToken()
    token.cancel()
    with pytest.raises(Cancelado):
        save_result(
            df, _spec(tmp_path, "xlsx", split=SplitSpec("rows", rows_per_part=70)), cancel=token
        )
    assert [p.name for p in tmp_path.iterdir()] == ["datos_usuario"]


# ------------------------------------------------------------------------------ rojo sin división


def test_rojo_sin_division_da_error_amigable_y_no_escribe(tmp_path):
    datos = platform_utils.user_data_dir()
    (datos / "limites.json").write_text(json.dumps({"excel": {"max_rows": 101}}), encoding="utf-8")
    df = _df_mediano(500)
    with pytest.raises(ExcelToolError) as exc:
        save_result(df, _spec(tmp_path, "xlsx"))
    texto = str(exc.value)
    assert "no entra" in texto and "Traceback" not in texto
    assert [p.name for p in tmp_path.iterdir()] == ["datos_usuario"]
    # con división sí se puede, y ninguna parte pasa el tope
    res = save_result(df, _spec(tmp_path, "xlsx", split=SplitSpec("rows", rows_per_part=100)))
    assert max(p.rows for p in res.parts) <= 100


def test_rojo_por_google_sheets_sin_division(tmp_path):
    datos = platform_utils.user_data_dir()
    (datos / "limites.json").write_text(
        json.dumps({"google_sheets": {"max_cells": 1000}}), encoding="utf-8"
    )
    df = _df_mediano(500)  # 4 columnas x 501 = 2004 celdas
    with pytest.raises(ExcelToolError):
        save_result(df, _spec(tmp_path, "csv_google", "google_sheets"))
    res = save_result(
        df, _spec(tmp_path, "csv_google", "google_sheets", SplitSpec("rows", rows_per_part=200))
    )
    assert max(p.rows for p in res.parts) <= 249  # 1000 // 4 - 1
    assert sum(p.rows for p in res.parts) == 500


def test_tabla_vacia_guarda_solo_el_encabezado(tmp_path, explorar):
    df, _ = explorar
    vacio = df.iloc[0:0]
    for fmt in ("xlsx", "csv_excel"):
        res = save_result(
            vacio,
            _spec(tmp_path, fmt, split=SplitSpec("rows", rows_per_part=10), nombre=f"v_{fmt}"),
        )
        assert res.rows_total == 0
        leido = read_table(res.files[0]).df
        assert list(leido.columns) == list(df.columns) and len(leido) == 0


# ------------------------------------------------------------------------------ archivos grandes


@pytest.fixture(scope="module")
def tabla_1600k(request):
    grandes = request.getfixturevalue("fixtures_grandes")
    return grandes / "fase3_dividir_1600k.csv"


@pytest.fixture(scope="module")
def esperado_grandes(fixtures_esperados):
    return json.loads((fixtures_esperados / "fase3_grandes.json").read_text(encoding="utf-8"))


def _contar_regiones(path: Path) -> Counter:
    """Recuento independiente del producto: módulo csv de la biblioteca estándar."""
    import csv

    cuenta: Counter = Counter()
    with path.open(encoding="utf-8-sig", newline="") as f:
        lector = csv.reader(f, delimiter=";")
        next(lector)
        for fila in lector:
            cuenta[fila[1]] += 1
    return cuenta


def _filas_csv(path: Path) -> int:
    n = 0
    with path.open("rb") as f:
        for _ in f:
            n += 1
    return n - 1


@pytest.mark.parametrize("formato", ["csv_excel", "csv_google"])
def test_split_file_por_valor_con_subdivision(tmp_path, tabla_1600k, esperado_grandes, formato):
    spec = OutputSpec(
        tmp_path / f"dividido{EXT[formato]}",
        formato,
        "excel",
        SplitSpec("by_value", value_columns=["Región"], rows_per_part=EXCEL_TOPE_FILAS),
    )
    res = split_file(tabla_1600k, ReadOptions(), spec)
    esp = esperado_grandes["regiones_dividir"]
    assert res.rows_total == 1_600_000 == sum(esp.values()) == sum(p.rows for p in res.parts)
    # Norte (1.200.000) pasa del tope de Excel y se subdivide en 2; las otras 4 van enteras
    assert len(res.parts) == 6
    assert max(p.rows for p in res.parts) <= EXCEL_TOPE_FILAS
    por_region = Counter()
    for p in res.parts:
        por_region[p.value[0]] += p.rows
    assert dict(por_region) == esp
    # encabezado en cada parte, y las filas del archivo = las declaradas
    cabecera = tabla_1600k.read_text(encoding="utf-8-sig").splitlines()[0].split(";")
    for p in res.parts:
        _nombre_valido_windows(p.path.name)
        leido = read_table(p.path, ReadOptions(nrows=5)).df
        assert list(leido.columns) == cabecera
        assert _filas_csv(p.path) == p.rows
    _verificar_indice(res, res.folder)
    assert sorted(x.name for x in res.folder.iterdir()) == sorted(
        [p.path.name for p in res.parts] + [NOMBRE_INDICE]
    )


def test_split_file_por_filas_y_sin_cortar_grupos(tmp_path, tabla_1600k):
    res = split_file(
        tabla_1600k,
        ReadOptions(),
        OutputSpec(
            tmp_path / "filas.csv", "csv_excel", "excel", SplitSpec("rows", rows_per_part=400_000)
        ),
    )
    assert [p.rows for p in res.parts] == [400_000] * 4
    assert [(p.first_row, p.last_row) for p in res.parts][-1] == (1_200_001, 1_600_000)
    assert all(_filas_csv(p.path) == p.rows for p in res.parts)
    # el contenido, en orden, es el original: la primera y la última fila de cada parte
    original = tabla_1600k.read_text(encoding="utf-8-sig").splitlines()
    for p in res.parts:
        lineas = p.path.read_text(encoding="utf-8-sig").splitlines()
        assert lineas[0] == original[0]
        assert lineas[1] == original[p.first_row]
        assert lineas[-1] == original[p.last_row]

    res2 = split_file(
        tabla_1600k,
        ReadOptions(),
        OutputSpec(
            tmp_path / "grupos.csv",
            "csv_excel",
            "excel",
            SplitSpec("rows_keep_groups", rows_per_part=EXCEL_TOPE_FILAS, group_column="Región"),
        ),
    )
    assert sum(p.rows for p in res2.parts) == 1_600_000
    regiones_por_parte = [set(_contar_regiones(p.path)) for p in res2.parts]
    for i, a in enumerate(regiones_por_parte):
        for b in regiones_por_parte[i + 1 :]:
            assert not (a & b)
    assert tmp_path / "filas (2)" not in list(tmp_path.iterdir())


def test_split_file_a_xlsx_con_el_archivo_de_400k(tmp_path, fixtures_grandes, esperado_grandes):
    """400k en xlsx por valor de Región: 5 partes de 80.000, releídas y con encabezado congelado."""
    path = fixtures_grandes / "fase3_400k.csv"
    res = split_file(
        path,
        ReadOptions(),
        OutputSpec(
            tmp_path / "x.xlsx",
            "xlsx",
            "excel",
            SplitSpec("by_value", value_columns=["Región"]),
            column_types={"ID": "entero"},
        ),
    )
    assert {p.value[0]: p.rows for p in res.parts} == esperado_grandes["regiones_400k"]
    assert all(p.path.suffix == ".xlsx" for p in res.parts)
    _verificar_indice(res, res.folder)
    primera = res.parts[0]
    leido = read_table(primera.path).df
    assert len(leido) == primera.rows == 80_000
    assert set(leido["Región"]) == {primera.value[0]}
    libro = openpyxl.load_workbook(primera.path, read_only=False)
    ws = libro.worksheets[0]
    assert ws.freeze_panes == "A2" and ws.auto_filter.ref


def test_split_file_cancelar_no_deja_nada(tmp_path, tabla_1600k):
    previo = tmp_path / "dividido"
    previo.mkdir()
    (previo / "viejo.txt").write_text("viejo", encoding="utf-8")
    antes = _foto(tmp_path)
    token = CancelToken()
    spec = OutputSpec(
        tmp_path / "dividido.csv", "csv_excel", "excel", SplitSpec("rows", rows_per_part=200_000)
    )
    with pytest.raises(Cancelado):
        split_file(tabla_1600k, ReadOptions(), spec, progress=_CancelarAl(token, 6), cancel=token)
    assert _foto(tmp_path) == antes


def test_split_file_cancelar_por_valor_no_deja_nada(tmp_path, tabla_1600k):
    token = CancelToken()
    spec = OutputSpec(
        tmp_path / "pv.csv", "csv_excel", "excel", SplitSpec("by_value", value_columns=["Región"])
    )
    with pytest.raises(Cancelado):
        split_file(tabla_1600k, ReadOptions(), spec, progress=_CancelarAl(token, 6), cancel=token)
    assert [p.name for p in tmp_path.iterdir()] == ["datos_usuario"]


def test_split_file_carpeta_existente_usa_numero_2(tmp_path, tabla_1600k):
    (tmp_path / "pv").mkdir()
    spec = OutputSpec(
        tmp_path / "pv.csv", "csv_excel", "excel", SplitSpec("rows", rows_per_part=800_000)
    )
    res = split_file(tabla_1600k, ReadOptions(), spec)
    assert res.folder == tmp_path / "pv (2)"
    assert list((tmp_path / "pv").iterdir()) == []


def test_1600k_en_excel_sin_division_es_rojo_y_no_escribe(tmp_path, tabla_1600k):
    df = read_table(tabla_1600k).df
    with pytest.raises(ExcelToolError) as exc:
        save_result(df, OutputSpec(tmp_path / "todo.xlsx", "xlsx", "excel"))
    assert "no entra" in str(exc.value)
    assert [p.name for p in tmp_path.iterdir()] == ["datos_usuario"]
    # "solo guardar" en CSV sí entra (verde): se escribe tal cual, 1.600.000 filas
    res = save_result(df, OutputSpec(tmp_path / "todo.csv", "csv_excel", "solo_guardar"))
    assert res.rows_total == 1_600_000 and _filas_csv(res.files[0]) == 1_600_000
    # ...pero en xlsx el límite duro de la hoja de Excel sigue valiendo
    with pytest.raises(ExcelToolError):
        save_result(df, OutputSpec(tmp_path / "todo2.xlsx", "xlsx", "solo_guardar"))


def test_regiones_del_archivo_coinciden_con_el_esperado(tabla_1600k, esperado_grandes):
    assert dict(_contar_regiones(tabla_1600k)) == esperado_grandes["regiones_dividir"]


def test_nombre_de_archivo_largo_no_pasa_de_260(tmp_path):
    """Carpeta profunda: las partes se acortan y la ruta completa no pasa de 260."""
    base = tmp_path
    while len(str(base)) < 190:
        base = base / ("carpeta_larga_" + "x" * 20)
    base.mkdir(parents=True)
    larga = "V" * 120
    df = pd.DataFrame(
        {"K": [larga + "1", larga + "2", "b"], "N": ["1", "2", "3"]}, dtype="string[pyarrow]"
    )
    res = save_result(
        df,
        OutputSpec(
            base / "s.csv", "csv_excel", "excel", SplitSpec("by_value", value_columns=["K"])
        ),
    )
    assert all(len(str(p.path)) <= 260 for p in res.parts)
    assert len({p.path.name.lower() for p in res.parts}) == 3
    assert any("acortaron" in w for w in res.warnings)
    assert sum(len(read_table(p.path).df) for p in res.parts) == 3


def test_regex_nombres_cubre_reservados():
    # control del propio test: el validador de nombres detecta lo que debe
    for malo in ("a:b.csv", "CON.csv", "x ", "x.", "a/b", "nul"):
        with pytest.raises(AssertionError):
            _nombre_valido_windows(malo)
    assert re.fullmatch(r"[^<>:\"/\\|?*]+", "Bahía Blanca_parte_01_de_03.xlsx")
