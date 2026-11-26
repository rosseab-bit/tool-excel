"""Integración de la Fase 4 en pantalla (xvfb) con el core REAL: explorador -> Datos -> "Resumir…".

Nada de `summarize` ni `save_result` falsos: solo se reemplazan los diálogos del sistema (en
`tests/integracion/conftest.py`) y `platform_utils.open_folder`. Los esperados salen del JSON del
generador o se recuentan con el módulo `csv` y `Decimal`.
"""

from __future__ import annotations

import csv
import json
import re
import tkinter as tk
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from excel_tool import platform_utils
from excel_tool.core.io import CsvOptions, read_table
from excel_tool.ui import editor_resumen, guardar
from excel_tool.ui.explorador import abrir_explorador
from excel_tool.ui.selector import Seleccion

pytestmark = pytest.mark.pantalla

JERGA = re.compile(
    r"\b(join|merge|group\s?by|aggregate|aggregation|distinct|split|null|dataframe|encoding|"
    r"dtype|chunk|spec|output|traceback)\b",
    re.IGNORECASE,
)


def _dec(texto: str) -> Decimal:
    t = texto.strip()
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    return Decimal(t)


def _textos(widget) -> list[str]:
    hallados = []
    for hijo in widget.winfo_children():
        for opcion in ("text", "values"):
            try:
                hallados.append(" ".join(str(x) for x in hijo.cget(opcion)))
            except tk.TclError:
                pass
        hallados.extend(_textos(hijo))
    return hallados


@pytest.fixture
def ventas(fixtures_generados) -> Path:
    return fixtures_generados / "escenario_3" / "ventas.csv"


@pytest.fixture
def esperado(fixtures_esperados) -> dict:
    ruta = fixtures_esperados / "escenario_3" / "resumen_ventas.json"
    return json.loads(ruta.read_text(encoding="utf-8"))


@pytest.fixture
def capturas(monkeypatch):
    """Guarda las ventanas de `abrir_resumen` y de `abrir_guardar`, y las llamadas a open_folder."""
    resumenes: list = []
    guardados: list = []
    carpetas: list[Path] = []
    original_r, original_g = editor_resumen.abrir_resumen, guardar.abrir_guardar

    def espia_resumen(*args, **kwargs):
        v = original_r(*args, **kwargs)
        resumenes.append(v)
        return v

    def espia_guardar(*args, **kwargs):
        v = original_g(*args, **kwargs)
        guardados.append(v)
        return v

    monkeypatch.setattr(editor_resumen, "abrir_resumen", espia_resumen)
    monkeypatch.setattr(guardar, "abrir_guardar", espia_guardar)
    monkeypatch.setattr(platform_utils, "open_folder", lambda p: carpetas.append(Path(p)))
    return resumenes, guardados


def _abrir_explorador(entorno, path: Path):
    ex = abrir_explorador(entorno.raiz, Seleccion(path, None, CsvOptions(";", "cp1252"), 0))
    entorno.esperar(lambda: ex.perfil is not None, 60)
    entorno.esperar(lambda: not ex.analizando, 60)
    return ex


def _abrir_resumir(entorno, ex, resumenes):
    ex.pestanas.select(ex.pestana_datos)
    entorno.esperar(lambda: not ex.panel_datos.boton_resumir.instate(["disabled"]))
    assert ex.panel_datos.boton_resumir.cget("text") == "Resumir…"
    ex.panel_datos.boton_resumir.invoke()
    entorno.esperar(lambda: len(resumenes) >= 1)
    return resumenes[-1]


def _calculo(entorno, editor, operacion: str, columna: str = ""):
    """Agrega un cálculo como lo haría el usuario: elige operación y columna en las listas."""
    fila = editor.agregar_calculo()
    fila.var_operacion.set(operacion)
    fila.combo_operacion.event_generate("<<ComboboxSelected>>")
    if columna:
        fila.var_columna.set(columna)
        fila.combo_columna.event_generate("<<ComboboxSelected>>")
    entorno.raiz.update()
    return fila


def _tildar(entorno, editor, columna: str, periodo: str | None = None):
    editor.checks[columna].invoke()
    if periodo:
        editor.var_periodo[columna].set(periodo)
        editor.combos_periodo[columna].event_generate("<<ComboboxSelected>>")
    entorno.raiz.update()


def _ver_como_queda(entorno, editor):
    editor.resultado = None
    editor.boton_previa.invoke()
    entorno.esperar(lambda: editor.resultado is not None, 60)
    return editor.resultado


def test_resumir_por_cliente_totales_aviso_y_guardar(
    entorno_ui, ventas, esperado, capturas, tmp_path
):
    resumenes, guardados = capturas
    ex = _abrir_explorador(entorno_ui, ventas)
    v = _abrir_resumir(entorno_ui, ex, resumenes)
    assert v is not None
    editor = v.editor
    assert v.ventana.title() == "Resumir"
    assert not editor.hay_algo_que_resumir()

    _tildar(entorno_ui, editor, "ID cliente")
    _calculo(entorno_ui, editor, "Sumar", "Importe")
    _calculo(entorno_ui, editor, "Contar filas")
    assert editor.etiqueta_frase.cget("text") == "Una fila por cada ID cliente"
    resultado = _ver_como_queda(entorno_ui, editor)

    # 29 IDs en ventas.csv: 27 clientes y los IDs 98 y 99 (que no están en clientes.xlsx).
    df = resultado.df
    assert list(df.columns) == ["ID cliente", "Suma de Importe", "Cantidad de filas"]
    assert len(df) == len(esperado["por_cliente"]) == 29
    por = esperado["por_cliente"]
    for _, f in df.iterrows():
        e = por[f["ID cliente"]]
        assert _dec(f["Suma de Importe"]) == Decimal(e["suma_importe"]), f["ID cliente"]
        assert int(f["Cantidad de filas"]) == e["filas"]
    assert set(df["ID cliente"]) == set(por) >= set(esperado["ids_fuera_de_clientes"])
    # Aviso de los 3 importes inválidos (y el vacío, aparte) en la pantalla y en el resultado.
    estado = editor.etiqueta_estado.cget("text")
    assert "29 filas en el resumen" in estado
    assert "3 celdas de Importe no son números y no se sumaron" in estado
    assert resultado.not_numeric == {"Importe": 3}
    assert any("Hay 1 celda vacía de Importe: no se suma" in w for w in resultado.warnings)
    # La grilla muestra el resumen (no el archivo original).
    entorno_ui.esperar(lambda: editor.grilla.columnas_visibles is not None)
    assert "Suma de Importe" in editor.grilla.columnas_visibles
    # Ni jerga ni carteles de error.
    for texto in _textos(v.ventana):
        assert not JERGA.search(texto), texto
    assert entorno_ui.carteles == []

    # "Guardar el resumen…" -> paso Guardar -> relee.
    v.boton_guardar.invoke()
    entorno_ui.esperar(lambda: len(guardados) == 1)
    g = guardados[0]
    assert g.nombre == "ventas_resumen" and len(g.df) == 29
    destino = tmp_path / "resumen_por_cliente.xlsx"
    entorno_ui.destino = destino
    entorno_ui.esperar(lambda: g.puede_guardar(), 60)
    g._guardar()
    entorno_ui.esperar(lambda: g._resultado is not None or entorno_ui.carteles, 120)
    assert entorno_ui.carteles == []
    assert g._resultado.files == [destino] and g._resultado.rows_total == 29

    releido = read_table(destino).df
    assert releido["ID cliente"].tolist() == df["ID cliente"].tolist()
    for a, b in zip(
        df["Suma de Importe"].fillna(""), releido["Suma de Importe"].fillna(""), strict=True
    ):
        assert (a == b == "") or _dec(a) == _dec(b)
    ws = openpyxl.load_workbook(destino).worksheets[0]
    cols = [c.value for c in ws[1]]
    assert cols == ["ID cliente", "Suma de Importe", "Cantidad de filas"]
    for fila in ws.iter_rows(min_row=2, values_only=True):
        if fila[0] in por or str(fila[0]) in por:
            e = por[str(int(fila[0]))]
            assert isinstance(fila[1], float) and Decimal(str(fila[1])) == Decimal(
                e["suma_importe"]
            )
            assert fila[2] == e["filas"]
    assert ws.max_row == 30


def test_resumir_por_mes_de_la_fecha(entorno_ui, ventas, capturas):
    resumenes, _ = capturas
    ex = _abrir_explorador(entorno_ui, ventas)
    v = _abrir_resumir(entorno_ui, ex, resumenes)
    editor = v.editor
    assert "Fecha" in editor.combos_periodo  # la columna de fechas ofrece "por mes" / "por año"
    assert list(editor.combos_periodo["Fecha"].cget("values")) == ["por día", "por mes", "por año"]

    _tildar(entorno_ui, editor, "Fecha", "por mes")
    _calculo(entorno_ui, editor, "Contar filas")
    _calculo(entorno_ui, editor, "Sumar", "Importe")
    assert "Fecha (por mes)" in editor.etiqueta_frase.cget("text")
    resultado = _ver_como_queda(entorno_ui, editor)

    # Recuento independiente con el módulo csv
    filas: dict[str, int] = defaultdict(int)
    sumas: dict[str, Decimal] = defaultdict(Decimal)
    with ventas.open(encoding="cp1252", newline="") as f:
        lector = csv.reader(f, delimiter=";")
        next(lector)
        for r in lector:
            d, m, a = r[1].split("/")
            filas[f"{a}-{m}"] += 1
            if re.fullmatch(r"[\d.]*\d,\d{2}", r[4]):
                sumas[f"{a}-{m}"] += _dec(r[4])
    df = resultado.df
    assert list(df.columns) == ["Fecha (mes)", "Cantidad de filas", "Suma de Importe"]
    assert {f["Fecha (mes)"]: int(f["Cantidad de filas"]) for _, f in df.iterrows()} == dict(filas)
    assert {f["Fecha (mes)"]: _dec(f["Suma de Importe"]) for _, f in df.iterrows()} == dict(sumas)
    assert sorted(filas) == [f"2024-0{n}" for n in range(1, 7)]
    assert sum(int(x) for x in df["Cantidad de filas"]) == 87
    assert entorno_ui.carteles == []


def test_resumir_lo_filtrado_y_sin_nada_elegido(entorno_ui, ventas, capturas):
    """Con un filtro en Datos, el resumen usa solo esas filas. Sin nada elegido, no calcula."""
    resumenes, _ = capturas
    ex = _abrir_explorador(entorno_ui, ventas)
    ex.pestanas.select(ex.pestana_datos)
    datos = ex.panel_datos
    datos.agregar_condicion("Artículo", "es igual a", "Tornillo")
    datos.aplicar()
    entorno_ui.esperar(lambda: len(datos.filas) < 87)
    entorno_ui.esperar(lambda: not datos.boton_aplicar.instate(["disabled"]))
    esperadas = len(datos.filas)
    with ventas.open(encoding="cp1252", newline="") as f:
        lector = csv.reader(f, delimiter=";")
        next(lector)
        assert esperadas == sum(1 for r in lector if r[2] == "Tornillo") > 0
    datos.boton_resumir.invoke()
    entorno_ui.esperar(lambda: len(resumenes) == 1)
    editor = resumenes[0].editor
    editor.boton_previa.invoke()  # nada elegido: avisa en vez de calcular
    assert editor.etiqueta_estado.cget("text")
    assert editor.resultado is None
    _calculo(entorno_ui, editor, "Contar filas")
    res = _ver_como_queda(entorno_ui, editor)
    assert len(res.df) == 1 and int(res.df.iloc[0, 0]) == esperadas  # toda la tabla = una fila
    assert entorno_ui.carteles == []


def test_la_ventana_de_resumir_entra_en_una_pantalla_de_768(entorno_ui, ventas, capturas):
    resumenes, _ = capturas
    ex = _abrir_explorador(entorno_ui, ventas)
    v = _abrir_resumir(entorno_ui, ex, resumenes)
    editor = v.editor
    _tildar(entorno_ui, editor, "ID cliente")
    _calculo(entorno_ui, editor, "Sumar", "Importe")
    _calculo(entorno_ui, editor, "Contar filas")
    _ver_como_queda(entorno_ui, editor)
    entorno_ui.raiz.update()
    # El pie con "Guardar el resumen…" tiene que quedar adentro de la pantalla.
    pie = v.boton_guardar.winfo_rooty() - v.ventana.winfo_rooty() + v.boton_guardar.winfo_height()
    assert v.ventana.winfo_reqheight() <= 768 and pie <= 768
