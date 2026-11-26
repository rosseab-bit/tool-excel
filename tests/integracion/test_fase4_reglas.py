"""Reglas de CLAUDE.md que suma la Fase 4, revisadas leyendo el código de src/ (AST y texto).

Completa a `test_fase2_reglas.py` y `test_fase3_reglas.py`: los resúmenes pasan por
`aggregate.summarize`, nada de tkinter en el core, textos de la UI sin jerga (con "group by",
"aggregate", "join", "merge") y límites solo en `core/limits.py`.
"""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

from tests.integracion.test_fase3_reglas import (
    LIMITES_NUMERICOS,
    SRC,
    _cadenas_de_texto,
    _constantes_numericas,
    _importados,
    _py,
    _rel,
)

JERGA_FASE4 = re.compile(
    r"\b(join|joins|merge|group\s?by|aggregate|aggregation|aggregations|distinct|split|null|"
    r"dataframe|encoding|dtype|spec|output|traceback)\b",
    re.IGNORECASE,
)
AGRUPACIONES = re.compile(r"\.(groupby|agg|aggregate|pivot_table|resample|rolling)\s*\(")


def _llamadas(ruta: Path) -> list[str]:
    """Nombres de las funciones y métodos que se llaman en el archivo."""
    nombres = []
    for n in ast.walk(ast.parse(ruta.read_text(encoding="utf-8"))):
        if isinstance(n, ast.Call):
            f = n.func
            nombres.append(f.attr if isinstance(f, ast.Attribute) else getattr(f, "id", ""))
    return nombres


def test_join_no_agrupa_por_su_cuenta_y_usa_summarize():
    ruta = SRC / "core" / "join.py"
    texto = ruta.read_text(encoding="utf-8")
    assert not AGRUPACIONES.search(texto), AGRUPACIONES.search(texto).group(0)
    llamadas = _llamadas(ruta)
    assert (
        llamadas.count("summarize") >= 3
    )  # secundario en resumen, conteo del final y resumen final
    importado = re.search(r"from excel_tool\.core\.aggregate import[^)]*\bsummarize\b", texto, re.S)
    assert importado, "join.py tiene que importar summarize de core/aggregate"
    assert "summarize" not in re.findall(r"^def (\w+)", texto, re.M)  # y no tiene el suyo


def test_agrupar_y_sumar_solo_en_aggregate_y_en_la_division_por_valor():
    """Los resúmenes son de `core/aggregate.py`. `output.py` agrupa solo para repartir filas."""
    con_agrupacion = {_rel(r) for r in _py() if AGRUPACIONES.search(r.read_text(encoding="utf-8"))}
    assert con_agrupacion == {"core/aggregate.py", "core/output.py"}, con_agrupacion


def test_la_ui_de_resumir_llama_al_core_y_no_calcula_nada():
    editor = SRC / "ui" / "editor_resumen.py"
    assert "core.summarize(" in editor.read_text(encoding="utf-8")  # lo pide por `_core()`
    assert not AGRUPACIONES.search(editor.read_text(encoding="utf-8"))
    assert not AGRUPACIONES.search((SRC / "ui" / "datos.py").read_text(encoding="utf-8"))


def test_nada_de_tkinter_en_el_core_de_la_fase_4():
    for nombre in ("aggregate", "join", "normalize", "recipe", "memory", "types"):
        ruta = SRC / "core" / f"{nombre}.py"
        assert ruta.is_file(), ruta
        assert not ({"tkinter", "_tkinter", "ttk"} & _importados(ruta)), nombre
        assert "excel_tool.ui" not in ruta.read_text(encoding="utf-8"), nombre
    for ruta in _py(SRC / "core"):
        assert not ({"tkinter", "_tkinter"} & _importados(ruta)), _rel(ruta)


def test_textos_de_la_ui_sin_jerga_de_la_fase_4():
    """Mira TODA cadena de `ui/` que el usuario podría ver (no docstrings ni logs)."""
    for ruta in _py(SRC / "ui"):
        for linea, texto in _cadenas_de_texto(ruta):
            if " " not in texto.strip():
                continue  # identificadores y claves ("<<ComboboxSelected>>", "left", "readonly")
            assert not JERGA_FASE4.search(texto), f"{_rel(ruta)}:{linea}: {texto!r}"


def test_los_textos_de_los_mensajes_del_core_nuevo_sin_jerga():
    """Frases de `aggregate`, `join`, `normalize` y `recipe` (errores y avisos al usuario)."""
    for nombre in ("aggregate", "join", "normalize", "recipe"):
        ruta = SRC / "core" / f"{nombre}.py"
        for linea, texto in _cadenas_de_texto(ruta):
            if " " not in texto.strip() or not texto[:1].isupper():
                continue  # solo frases (empiezan con mayúscula), no claves ni textos técnicos
            assert not JERGA_FASE4.search(texto), f"{_rel(ruta)}:{linea}: {texto!r}"


@pytest.mark.parametrize(
    "palabra", ["join", "merge", "group by", "groupby", "aggregate", "aggregation", "distinct"]
)
def test_el_detector_de_jerga_ve_las_palabras_de_la_fase_4(palabra):
    assert JERGA_FASE4.search(f"Hacemos un {palabra} de las tablas")
    assert not JERGA_FASE4.search("Se unieron 3 archivos y se resumió por cliente")


def test_los_limites_siguen_solo_en_limits_con_el_codigo_nuevo():
    for nombre in ("aggregate", "join", "normalize", "recipe", "memory"):
        ruta = SRC / "core" / f"{nombre}.py"
        hallazgo = _constantes_numericas(ruta.read_text(encoding="utf-8")) & LIMITES_NUMERICOS
        assert not hallazgo, f"{nombre}: {sorted(hallazgo)}"
    assert (SRC / "core" / "limits.py").is_file()
