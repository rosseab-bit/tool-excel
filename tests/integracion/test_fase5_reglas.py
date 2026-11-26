"""Reglas de CLAUDE.md que suma la Fase 5 (asistente en `ui/unir/`), revisadas leyendo el código.

- Textos sin jerga (ahora también "join", "merge", "left", "inner", "key", "plan").
- La UI no calcula: nada de `groupby`, `merge`, `.sum(` y parecidos en `ui/`.
- Nada de widgets desde hilos: lo que se manda a `run_in_background` no toca tkinter.
- Rutas con `pathlib` y `encoding=` explícito en lo nuevo.
"""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

from tests.integracion.test_fase3_reglas import SRC, _cadenas_de_texto, _py, _rel

UNIR = SRC / "ui" / "unir"
JERGA_F5 = re.compile(
    r"\b(join|joins|merge|left|inner|key|keys|plan|group\s?by|aggregate|aggregation|distinct|"
    r"split|null|dataframe|encoding|dtype|chunk|spec|output|traceback)\b",
    re.IGNORECASE,
)
# "plan" también es castellano ("un plan"): se mira solo como palabra técnica suelta en inglés
CALCULOS = re.compile(
    r"\.(groupby|merge|agg|aggregate|pivot_table|resample|rolling|sum|mean|cumsum|value_counts"
    r"|nunique|apply|concat)\s*\(|\bpd\.(merge|concat)\b"
)


def _archivos_unir() -> list[Path]:
    archivos = _py(UNIR)
    assert len(archivos) >= 15, archivos  # estado, asistente, widgets, tamano, listas, 7 pasos...
    return archivos


def test_textos_de_ui_unir_sin_jerga_ni_las_palabras_en_ingles():
    for ruta in _archivos_unir():
        for linea, texto in _cadenas_de_texto(ruta):
            if " " not in texto.strip():
                continue  # identificadores y claves ("left", "readonly", "<<ComboboxSelected>>")
            assert not JERGA_F5.search(texto), f"{_rel(ruta)}:{linea}: {texto!r}"


def test_textos_de_toda_la_ui_sin_join_merge_left_inner_key():
    for ruta in _py(SRC / "ui"):
        for linea, texto in _cadenas_de_texto(ruta):
            if " " not in texto.strip():
                continue
            assert not re.search(r"\b(join|merge|left|inner|key)\b", texto, re.I), (
                f"{_rel(ruta)}:{linea}: {texto!r}"
            )


def test_las_palabras_sueltas_que_ve_el_usuario_no_son_jerga():
    """Una sola palabra puede ser un identificador o un texto (un botón "Plan"): se listan las
    cadenas de una palabra que coinciden con la jerga, para que no se cuele ninguna."""
    sueltas = []
    for ruta in _archivos_unir():
        for linea, texto in _cadenas_de_texto(ruta):
            if " " not in texto.strip() and JERGA_F5.fullmatch(texto.strip()):
                sueltas.append((_rel(ruta), linea, texto))
    # solo valores internos del core ("left", "inner") usados como claves de diccionario
    assert {t for _, _, t in sueltas} <= {"left", "inner"}, sueltas


def test_el_detector_de_jerga_f5_ve_las_palabras_nuevas():
    for palabra in ("join", "merge", "left", "inner", "key", "plan"):
        assert JERGA_F5.search(f"Hacemos un {palabra} de las tablas"), palabra
    assert not JERGA_F5.search("Se unieron 3 archivos por la columna clave")


def test_la_ui_no_calcula_nada():
    """Nada de groupby / merge / sum en `ui/`: unir, sumar y resumir son del core."""
    for ruta in _py(SRC / "ui"):
        texto = ruta.read_text(encoding="utf-8")
        sin_comentarios = re.sub(r'"""[\s\S]*?"""|#.*', "", texto)
        hallazgo = CALCULOS.search(sin_comentarios)
        assert not hallazgo, f"{_rel(ruta)}: {hallazgo.group(0)}"


def test_ui_unir_pide_todo_al_core():
    pedidos = {
        "paso3_sumar.py": "diagnose_join",
        "paso4_modo.py": "diagnose_join",
        "paso7_guardar.py": "run_join",
        "tamano.py": "suggest_reductions",
    }
    for nombre, funcion in pedidos.items():
        ruta = next(r for r in _archivos_unir() if r.name == nombre)
        assert funcion in ruta.read_text(encoding="utf-8"), (nombre, funcion)
    assert "save_workbook" in (UNIR / "pasos" / "paso7_guardar.py").read_text(encoding="utf-8")


# ---------------------------------------------------------------- nada de widgets desde hilos

_TOCA_WIDGETS = {
    "configure", "config", "insert", "delete", "state", "pack", "pack_forget", "grid",
    "grid_forget", "place", "after", "after_cancel", "update", "update_idletasks", "focus_set",
    "invoke", "set", "select", "selection_set", "winfo_exists", "destroy", "lift", "title",
    "showinfo", "showwarning", "showerror", "askyesno", "askokcancel",
}  # fmt: skip


def _funciones_enviadas_al_hilo(arbol: ast.AST) -> list[str]:
    """Nombres de las funciones que se pasan como tarea (2.º argumento) a `run_in_background`."""
    nombres = []
    for n in ast.walk(arbol):
        if isinstance(n, ast.Call):
            f = n.func
            nombre = f.attr if isinstance(f, ast.Attribute) else getattr(f, "id", "")
            if nombre == "run_in_background" and len(n.args) >= 2:
                tarea = n.args[1]
                if isinstance(tarea, ast.Name):
                    nombres.append(tarea.id)
                elif isinstance(tarea, ast.Lambda):
                    nombres.append("<lambda>")
    return nombres


def _llamadas_prohibidas(nodo: ast.AST) -> list[str]:
    malas = []
    for n in ast.walk(nodo):
        if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute):
            if n.func.attr in _TOCA_WIDGETS:
                malas.append(f"{ast.unparse(n.func)} (línea {n.lineno})")
    return malas


def test_lo_que_corre_en_el_hilo_no_toca_widgets():
    revisadas = 0
    for ruta in _py(SRC / "ui"):
        arbol = ast.parse(ruta.read_text(encoding="utf-8"))
        enviadas = set(_funciones_enviadas_al_hilo(arbol))
        if not enviadas:
            continue
        metodos = {
            f.name: f
            for c in ast.walk(arbol)
            if isinstance(c, ast.ClassDef)
            for f in c.body
            if isinstance(f, ast.FunctionDef)
        }
        for f in ast.walk(arbol):
            if isinstance(f, ast.FunctionDef) and f.name in enviadas:
                revisadas += 1
                assert not _llamadas_prohibidas(f), (_rel(ruta), f.name, _llamadas_prohibidas(f))
                # y los métodos propios que llama (self._hojas(...)) tampoco
                for n in ast.walk(f):
                    if (
                        isinstance(n, ast.Call)
                        and isinstance(n.func, ast.Attribute)
                        and isinstance(n.func.value, ast.Name)
                        and n.func.value.id == "self"
                        and n.func.attr in metodos
                    ):
                        malas = _llamadas_prohibidas(metodos[n.func.attr])
                        assert not malas, (_rel(ruta), f.name, n.func.attr, malas)
    assert revisadas >= 10, revisadas  # el chequeo mira de verdad las tareas del asistente


def test_el_detector_de_widgets_en_hilos_funciona():
    malo = ast.parse("def tarea(p, c):\n    self.etiqueta.configure(text='x')\n")
    assert _llamadas_prohibidas(malo)
    bueno = ast.parse("def tarea(p, c):\n    return run_join(plan, tablas, progress=p)\n")
    assert not _llamadas_prohibidas(bueno)


def test_las_tareas_del_asistente_se_mandan_al_worker():
    """Leer, diagnosticar, unir y guardar van por `run_in_background` (no en el hilo de la UI)."""
    esperados = {
        "paso3_sumar.py": 1,
        "paso4_modo.py": 1,
        "paso5_columnas.py": 1,
        "paso6_resumen.py": 1,
        "paso7_guardar.py": 4,
    }
    for ruta in _archivos_unir():
        if ruta.name in esperados:
            llamadas = ruta.read_text(encoding="utf-8").count("run_in_background(")
            assert llamadas >= esperados[ruta.name], (ruta.name, llamadas)


# ---------------------------------------------------------------- compatibilidad


def test_lo_nuevo_abre_archivos_con_encoding_y_sin_sys_platform():
    for ruta in _archivos_unir():
        texto = ruta.read_text(encoding="utf-8")
        assert "sys.platform" not in texto and "subprocess" not in texto, ruta.name
        for n in ast.walk(ast.parse(texto)):
            if isinstance(n, ast.Call) and getattr(n.func, "id", "") == "open":
                assert any(k.arg == "encoding" for k in n.keywords) or len(n.args) >= 2, ruta.name
        assert "print(" not in texto, ruta.name  # la app no escribe a stdout


@pytest.mark.parametrize("nombre", ["pandas", "openpyxl"])
def test_ui_unir_no_lee_ni_escribe_tablas(nombre):
    for ruta in _archivos_unir():
        texto = ruta.read_text(encoding="utf-8")
        assert not re.search(rf"{nombre}\.(read_|to_)", texto), ruta.name
        assert not re.search(r"\.(to_excel|to_csv|read_csv|read_excel)\(", texto), ruta.name
