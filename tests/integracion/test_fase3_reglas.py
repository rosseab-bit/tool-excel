"""Reglas de CLAUDE.md que suma la Fase 3, revisadas leyendo el código de src/ (AST y texto).

Completa a `test_fase2_reglas.py`: límites solo en `core/limits.py`, escritura de tablas solo en
`core/io` (también el `csv_bytes` nuevo), openpyxl fuera de `src/`, nada de tkinter en `core/` y
textos de la UI sin jerga (con las palabras que se sumaron en esta fase).
"""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

SRC = Path(__file__).resolve().parents[2] / "src" / "excel_tool"
RAIZ = SRC.parents[1]
JERGA_FASE3 = re.compile(
    r"\b(join|merge|group\s?by|distinct|split|null|dataframe|encoding|dtype|traceback|"
    r"chunk|spec|output)\b",
    re.IGNORECASE,
)
LIMITES = re.compile(
    r"1_?048_?57[56]|20_?000_?000|16_?384|(?<![\d_])5_?000_?000|(?<![\d_])300_?000"
)


def _py(carpeta: Path = SRC):
    return sorted(carpeta.rglob("*.py"))


def _rel(ruta: Path) -> str:
    return ruta.relative_to(SRC).as_posix()


def _importados(ruta: Path) -> set[str]:
    arbol = ast.parse(ruta.read_text(encoding="utf-8"))
    mods: set[str] = set()
    for n in ast.walk(arbol):
        if isinstance(n, ast.Import):
            mods.update(a.name.split(".")[0] for a in n.names)
        elif isinstance(n, ast.ImportFrom) and n.level == 0:
            mods.add((n.module or "").split(".")[0])
    return mods


def _cadenas_de_texto(ruta: Path):
    """Cadenas que no son docstrings ni argumentos de log (lo que podría ver el usuario)."""
    arbol = ast.parse(ruta.read_text(encoding="utf-8"))
    excluidos: set[int] = set()
    for n in ast.walk(arbol):
        if isinstance(n, ast.Module | ast.ClassDef | ast.FunctionDef | ast.AsyncFunctionDef):
            if n.body and isinstance(n.body[0], ast.Expr):
                v = n.body[0].value
                if isinstance(v, ast.Constant) and isinstance(v.value, str):
                    excluidos.add(id(v))
        if (
            isinstance(n, ast.Call)
            and isinstance(n.func, ast.Attribute)
            and isinstance(n.func.value, ast.Name)
            and n.func.value.id in {"log", "logger", "logging"}
        ):
            excluidos.update(id(h) for h in ast.walk(n))
    for n in ast.walk(arbol):
        if isinstance(n, ast.Constant) and isinstance(n.value, str) and id(n) not in excluidos:
            yield n.lineno, n.value


def test_los_limites_estan_solo_en_core_limits():
    for ruta in _py():
        if _rel(ruta) == "core/limits.py":
            continue
        hallazgo = LIMITES.search(ruta.read_text(encoding="utf-8"))
        assert not hallazgo, f"{_rel(ruta)}: {hallazgo.group(0)}"


def test_el_regex_de_limites_detecta_lo_que_debe():
    for malo in ("1048576", "1_048_576", "1048575", "20_000_000", "5_000_000", "300000"):
        assert LIMITES.search(f"x = {malo}"), malo
    assert not LIMITES.search("x = 1500000") and not LIMITES.search("x = 25_000_000")


def test_json_de_limites_solo_se_lee_en_limits():
    """(`app_entry.py` lo abre solo para el self-test del ejecutable.)"""
    for ruta in _py():
        if _rel(ruta) in {"core/limits.py", "app_entry.py"}:
            continue
        assert "limites_default" not in ruta.read_text(encoding="utf-8"), _rel(ruta)


def test_nada_de_pandas_io_ni_escritores_fuera_de_core_io():
    pandas_io = re.compile(
        r"\.(read_csv|read_excel|to_csv|to_excel|ExcelWriter)\b|\bpd\.ExcelWriter"
    )
    for ruta in _py():
        rel = _rel(ruta)
        if rel.startswith("core/io/"):
            continue
        texto = ruta.read_text(encoding="utf-8")
        assert not pandas_io.search(texto), rel
        assert "xlsxwriter" not in _importados(ruta), rel
        assert "csv" not in _importados(ruta), rel  # el módulo csv también es de core/io
        assert "calamine" not in "".join(_importados(ruta)), rel


def test_csv_bytes_es_de_core_io_y_los_demas_lo_piden_prestado():
    definidas = [
        _rel(r)
        for r in _py()
        if re.search(r"^def csv_bytes\w*\(", r.read_text(encoding="utf-8"), re.MULTILINE)
    ]
    assert sorted(definidas) == ["core/io/__init__.py", "core/io/writer.py"]


def test_openpyxl_no_se_importa_en_src():
    for ruta in _py():
        assert "openpyxl" not in _importados(ruta), _rel(ruta)
        assert not re.search(
            r"^\s*(import|from)\s+openpyxl", ruta.read_text(encoding="utf-8"), re.M
        )


def test_openpyxl_solo_en_el_extra_dev():
    texto = (RAIZ / "pyproject.toml").read_text(encoding="utf-8")
    dev = re.search(r"dev\s*=\s*\[(.*?)\]", texto, re.S)
    assert dev and "openpyxl" in dev.group(1)
    base = re.search(r"^dependencies\s*=\s*\[(.*?)\]", texto, re.S | re.M)
    assert base and "openpyxl" not in base.group(1) and "xlsxwriter" in base.group(1)


def test_core_sin_tkinter_ni_ui():
    for ruta in _py(SRC / "core"):
        mods = _importados(ruta)
        assert not ({"tkinter", "_tkinter"} & mods), _rel(ruta)
        arbol = ast.parse(ruta.read_text(encoding="utf-8"))
        for n in ast.walk(arbol):
            if isinstance(n, ast.ImportFrom) and n.module:
                assert not n.module.startswith("excel_tool.ui"), _rel(ruta)


def test_sys_frozen_y_meipass_solo_en_platform_utils():
    for ruta in _py():
        if ruta.name == "platform_utils.py":
            continue
        assert not re.search(
            r"sys\.frozen|_MEIPASS|getattr\(sys", ruta.read_text(encoding="utf-8")
        ), _rel(ruta)


def test_save_result_es_la_unica_salida_de_resultados():
    """Ninguna pieza de la UI escribe archivos de resultado por su cuenta."""
    for ruta in _py(SRC / "ui"):
        texto = ruta.read_text(encoding="utf-8")
        assert "write_table" not in texto or ruta.name == "exportar.py", _rel(ruta)
        assert "write_sheets" not in texto, _rel(ruta)
    assert "save_result" in (SRC / "ui" / "guardar.py").read_text(encoding="utf-8")


def test_textos_de_la_ui_sin_jerga_ampliada():
    problemas = []
    for ruta in _py(SRC / "ui"):
        for linea, texto in _cadenas_de_texto(ruta):
            if " " in texto.strip() and JERGA_FASE3.search(texto):
                problemas.append(f"{ruta.name}:{linea}: {texto!r}")
    assert problemas == []


@pytest.mark.parametrize("palabra", ["split", "chunk", "spec", "output", "join", "DataFrame"])
def test_el_detector_de_jerga_ve_las_palabras_nuevas(palabra):
    assert JERGA_FASE3.search(f"Se hizo un {palabra} del archivo")
    assert not JERGA_FASE3.search("Se guardaron 3 archivos con 200 filas en total")
