"""Reglas de CLAUDE.md revisadas leyendo el código de src/ (los grep de QA, automáticos)."""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

SRC = Path(__file__).resolve().parents[2] / "src" / "excel_tool"
JERGA = re.compile(
    r"\b(join|merge|group\s?by|distinct|split|null|dataframe|encoding|dtype|traceback)\b",
    re.IGNORECASE,
)
FUERA_DE_IO = re.compile(r"\b(read_csv|read_excel|to_csv|to_excel|ExcelWriter)\b")


def _py(carpeta: Path):
    return sorted(carpeta.rglob("*.py"))


def _textos_visibles(arbol: ast.AST):
    """Cadenas del código que no son docstrings ni argumentos de log.*()."""
    excluidos: set[int] = set()
    for nodo in ast.walk(arbol):
        if isinstance(nodo, ast.Module | ast.ClassDef | ast.FunctionDef | ast.AsyncFunctionDef):
            if nodo.body and isinstance(nodo.body[0], ast.Expr):
                valor = nodo.body[0].value
                if isinstance(valor, ast.Constant) and isinstance(valor.value, str):
                    excluidos.add(id(valor))
        if (
            isinstance(nodo, ast.Call)
            and isinstance(nodo.func, ast.Attribute)
            and isinstance(nodo.func.value, ast.Name)
            and nodo.func.value.id in {"log", "logger", "logging"}
        ):
            for hijo in ast.walk(nodo):
                excluidos.add(id(hijo))
    for nodo in ast.walk(arbol):
        if (
            isinstance(nodo, ast.Constant)
            and isinstance(nodo.value, str)
            and id(nodo) not in excluidos
        ):
            yield nodo.lineno, nodo.value


def test_textos_de_la_ui_sin_jerga():
    problemas = []
    for ruta in _py(SRC / "ui"):
        arbol = ast.parse(ruta.read_text(encoding="utf-8"))
        for linea, texto in _textos_visibles(arbol):
            if " " in texto.strip() and JERGA.search(texto):
                problemas.append(f"{ruta.name}:{linea}: {texto!r}")
    assert problemas == []


def test_core_sin_tkinter():
    for ruta in _py(SRC / "core"):
        arbol = ast.parse(ruta.read_text(encoding="utf-8"))
        for nodo in ast.walk(arbol):
            if isinstance(nodo, ast.Import):
                assert not any(
                    a.name.split(".")[0] in {"tkinter", "_tkinter"} for a in nodo.names
                ), ruta
            if isinstance(nodo, ast.ImportFrom):
                assert (nodo.module or "").split(".")[0] not in {"tkinter", "_tkinter"}, ruta


def test_lectura_y_escritura_de_tablas_solo_en_core_io():
    for ruta in _py(SRC):
        if "core/io" in ruta.as_posix():
            continue
        assert not FUERA_DE_IO.search(ruta.read_text(encoding="utf-8")), ruta


def test_sys_platform_solo_en_platform_utils():
    for ruta in _py(SRC):
        if ruta.name == "platform_utils.py":
            continue
        texto = ruta.read_text(encoding="utf-8")
        assert not re.search(r"sys\.platform|os\.name\b|platform\.system", texto), ruta


def test_limites_no_hardcodeados():
    for ruta in _py(SRC):
        assert not re.search(
            r"1_?048_?576|20_?000_?000|16_?384", ruta.read_text(encoding="utf-8")
        ), ruta


def test_open_siempre_con_encoding_o_binario():
    for ruta in _py(SRC):
        arbol = ast.parse(ruta.read_text(encoding="utf-8"))
        for nodo in ast.walk(arbol):
            if not isinstance(nodo, ast.Call):
                continue
            f = nodo.func
            es_open = (isinstance(f, ast.Name) and f.id == "open") or (
                isinstance(f, ast.Attribute) and f.attr == "open"
            )
            if not es_open:
                continue
            modo = None
            for a in nodo.args[:2]:
                if isinstance(a, ast.Constant) and isinstance(a.value, str):
                    modo = a.value
            binario = modo is not None and "b" in modo
            con_encoding = any(k.arg == "encoding" for k in nodo.keywords)
            assert binario or con_encoding, f"{ruta.name}:{nodo.lineno}"
        for nodo in ast.walk(arbol):
            if isinstance(nodo, ast.Call) and isinstance(nodo.func, ast.Attribute):
                if nodo.func.attr in {"read_text", "write_text"}:
                    assert any(k.arg == "encoding" for k in nodo.keywords), (
                        f"{ruta.name}:{nodo.lineno}"
                    )


def test_nada_a_stdout_salvo_la_consola_del_modo_prueba():
    permitido = {"app_entry.py"}  # --version y --self-test (contrato de la Fase 0)
    for ruta in _py(SRC):
        if ruta.name in permitido:
            continue
        arbol = ast.parse(ruta.read_text(encoding="utf-8"))
        for nodo in ast.walk(arbol):
            if isinstance(nodo, ast.Call) and isinstance(nodo.func, ast.Name):
                assert nodo.func.id != "print", f"{ruta.name}:{nodo.lineno}"
            if isinstance(nodo, ast.Attribute) and isinstance(nodo.value, ast.Name):
                assert not (nodo.value.id == "sys" and nodo.attr in {"stdout", "stderr"}), ruta


def test_subprocess_solo_en_platform_utils():
    for ruta in _py(SRC):
        if ruta.name == "platform_utils.py":
            continue
        assert not re.search(
            r"\bsubprocess\b|os\.system|Popen", ruta.read_text(encoding="utf-8")
        ), ruta


@pytest.mark.parametrize("nombre", ["tmp", "TEMP"])
def test_sin_temporales_fijos(nombre):
    for ruta in _py(SRC):
        assert f'"/{nombre}' not in ruta.read_text(encoding="utf-8"), ruta
