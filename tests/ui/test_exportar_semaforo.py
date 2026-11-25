"""F3-11: exportar lo filtrado, semáforo Excel / Google Sheets y tamaño de la ventana Guardar."""

from __future__ import annotations

import os
import sys
import time
import tkinter as tk
from types import SimpleNamespace

import pandas as pd
import pytest

from excel_tool.core.io import CsvOptions
from excel_tool.ui import datos, explorador, guardar, selector

from .perfiles_falsos import columna, hojas_falsas, perfil_falso
from .test_guardar import CoreFalso, _listo

pytestmark = pytest.mark.pantalla

REGIONES = ["Norte", "NORTE", "Norte ", "Sur", "Sur", "", "Este"]
IMPORTES = ["10", "20", "abc", "5", "100", "7", ""]
CIUDADES = ["A", "A", "B", "B", "B", "A", "C"]


def _esperar(raiz, condicion, segundos=10):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        raiz.update()
        if condicion():
            return
        time.sleep(0.02)
    raise AssertionError("se agotó la espera")


def _textos(widget):
    encontrados = []
    for hijo in widget.winfo_children():
        try:
            encontrados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        encontrados.extend(_textos(hijo))
    return encontrados


def _tabla():
    return pd.DataFrame(
        {"Región": REGIONES, "Importe": IMPORTES, "Ciudad": CIUDADES}, dtype="string[pyarrow]"
    )


def _perfil():
    perfil = perfil_falso()
    perfil.rows = 7
    perfil.columns_profile = [
        columna("Región", 0, "texto"),
        columna("Importe", 1, "entero"),
        columna("Ciudad", 2, "texto", is_unique=True),
    ]
    perfil.columns = 3
    return perfil


@pytest.fixture
def raiz(tmp_path, monkeypatch):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos"))
    monkeypatch.setattr(explorador, "list_sheets", lambda path: hojas_falsas())
    monkeypatch.setattr(explorador, "_limites", lambda: None)
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


def _abrir(raiz, monkeypatch, df):
    perfil = _perfil()
    monkeypatch.setattr(
        explorador, "read_table", lambda path, opciones: SimpleNamespace(sheets=[], df=df)
    )
    monkeypatch.setattr(
        explorador,
        "perfilar",
        lambda leido, *, on_column=None, progress=None, cancel=None, sheets=None: perfil,
    )
    seleccion = selector.Seleccion(perfil.file.path, None, CsvOptions(";", "utf-8-sig"), 0)
    ex = explorador.abrir_explorador(raiz, seleccion)
    _esperar(raiz, lambda: ex.perfil is not None)
    return ex


@pytest.fixture
def llamadas(monkeypatch):
    registro = []
    monkeypatch.setattr(guardar, "abrir_guardar", lambda *a, **k: registro.append((a, k)))
    return registro


# ---------------------------------------------------------------- exportar lo filtrado
def test_exportar_filtrado_y_ordenado_sin_numero_de_fila(raiz, monkeypatch, llamadas):
    df = _tabla()
    ex = _abrir(raiz, monkeypatch, df)
    panel = ex.panel_datos
    panel.agregar_condicion("Ciudad", "es igual a", "B")
    panel.aplicar()
    _esperar(raiz, lambda: len(panel.filas) == 3)
    panel._ordenar("Región", False)
    _esperar(raiz, lambda: list(panel.filas) != [2, 3, 4])
    esperado = df.iloc[list(panel.filas)].reset_index(drop=True)

    assert panel.boton_exportar.cget("text") == "Exportar lo filtrado…"
    panel.boton_exportar.invoke()
    _esperar(raiz, lambda: llamadas)
    (args, kwargs) = llamadas[0]
    tabla = args[1]
    assert "N.º fila" not in tabla.columns
    assert list(tabla.columns) == ["Región", "Importe", "Ciudad"]
    pd.testing.assert_frame_equal(tabla, esperado)
    assert kwargs["nombre_sugerido"] == "ventas_filtrado"
    assert kwargs["column_types"] == {"Importe": "entero"}
    assert kwargs["claves_sugeridas"] == ["Ciudad"]


def test_sin_filtro_pasa_la_tabla_original(raiz, monkeypatch, llamadas):
    df = _tabla()
    ex = _abrir(raiz, monkeypatch, df)
    panel = ex.panel_datos
    assert panel.boton_exportar.cget("text") == "Guardar estas filas…"
    panel.boton_exportar.invoke()
    (args, kwargs) = llamadas[0]
    assert args[1] is df
    assert kwargs["nombre_sugerido"] == "ventas"


def test_columnas_ocultas_y_elige_todas(raiz, monkeypatch, llamadas):
    df = _tabla()
    ex = _abrir(raiz, monkeypatch, df)
    panel = ex.panel_datos
    preguntas = []
    monkeypatch.setattr(datos.messagebox, "askyesno", lambda *a, **k: preguntas.append(a) or True)
    panel.grilla.ocultar_columna("Importe")
    panel.boton_exportar.invoke()
    assert len(preguntas) == 1
    (args, kwargs) = llamadas[0]
    assert args[1] is df  # todas las columnas y todas las filas: sin copia
    assert kwargs["column_types"] == {"Importe": "entero"}


def test_columnas_ocultas_y_elige_solo_las_visibles(raiz, monkeypatch, llamadas):
    df = _tabla()
    ex = _abrir(raiz, monkeypatch, df)
    panel = ex.panel_datos
    monkeypatch.setattr(datos.messagebox, "askyesno", lambda *a, **k: False)
    panel.grilla.ocultar_columna("Importe")
    panel.boton_exportar.invoke()
    _esperar(raiz, lambda: llamadas)
    (args, kwargs) = llamadas[0]
    assert list(args[1].columns) == ["Región", "Ciudad"]
    assert len(args[1]) == 7
    assert kwargs["column_types"] == {}
    assert kwargs["claves_sugeridas"] == ["Ciudad"]


def test_sin_columnas_ocultas_no_pregunta(raiz, monkeypatch, llamadas):
    ex = _abrir(raiz, monkeypatch, _tabla())
    # el doble de `askyesno` de conftest falla si se lo llama
    ex.panel_datos.boton_exportar.invoke()
    assert len(llamadas) == 1


# --------------------------------------------------------------------------- semáforo
def _limites_falsos(niveles):
    registro = []

    def check(est, destino, formato, **kw):
        registro.append((destino, formato))
        nivel = niveles[destino]
        razones = {
            "ok": [],
            "recomendado": ["Conviene dividirlo: tiene muchas filas."],
            "obligatorio": ["No entra: tiene demasiadas celdas.", "Otro motivo."],
        }[nivel]
        return SimpleNamespace(level=nivel, reasons=razones)

    modulo = SimpleNamespace(estimate=lambda df: SimpleNamespace(rows=len(df)), check=check)
    return modulo, registro


def test_semaforo_con_los_tres_colores_y_boton_dividir(raiz, monkeypatch, llamadas):
    modulo, registro = _limites_falsos({"excel": "recomendado", "google_sheets": "obligatorio"})
    monkeypatch.setattr(explorador, "_limites", lambda: modulo)
    df = _tabla()
    ex = _abrir(raiz, monkeypatch, df)
    textos = _textos(ex.contenido_general)
    assert "¿Entra en Excel o en Google Sheets?" in textos
    assert any(t.startswith("●  Excel: CONVIENE DIVIDIR") for t in textos)
    assert any(t.startswith("●  Google Sheets: NO ENTRA") for t in textos)
    assert "Conviene dividirlo: tiene muchas filas." in textos
    assert "No entra: tiene demasiadas celdas." in textos
    assert "Otro motivo." not in textos  # solo el primer motivo
    assert registro == [("excel", "xlsx"), ("google_sheets", "csv_google")]
    colores = [m.cget("bg") for m, _ in ex.lineas_semaforo]
    assert colores == [guardar.NIVELES["recomendado"][1], guardar.NIVELES["obligatorio"][1]]
    assert "Dividir y guardar…" in textos

    ex.boton_dividir.invoke()
    (args, kwargs) = llamadas[0]
    assert args[1] is df
    assert kwargs["column_types"] == {"Importe": "entero"}
    assert kwargs["claves_sugeridas"] == ["Ciudad"]
    assert kwargs["nombre_sugerido"] == "ventas"


def test_semaforo_verde_no_tiene_boton_dividir(raiz, monkeypatch):
    modulo, _ = _limites_falsos({"excel": "ok", "google_sheets": "ok"})
    monkeypatch.setattr(explorador, "_limites", lambda: modulo)
    ex = _abrir(raiz, monkeypatch, _tabla())
    textos = _textos(ex.contenido_general)
    assert any("Excel: TODO BIEN" in t for t in textos)
    assert any("Google Sheets: TODO BIEN" in t for t in textos)
    assert "Dividir y guardar…" not in textos


def test_sin_core_de_limites_no_hay_semaforo(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch, _tabla())  # `_limites` devuelve None en el fixture
    assert "¿Entra en Excel o en Google Sheets?" not in _textos(ex.contenido_general)


def test_semaforo_con_el_core_real(raiz, monkeypatch):
    pytest.importorskip("excel_tool.core.limits")
    import excel_tool.core.limits as limits

    monkeypatch.setattr(explorador, "_limites", lambda: limits)
    res = explorador.calcular_semaforo(_tabla())
    assert [(n, c.level) for n, c in res] == [("Excel", "ok"), ("Google Sheets", "ok")]


# ------------------------------------------------------------- ventana de Guardar
def _abrir_guardar(raiz, monkeypatch, nivel="obligatorio", **kw):
    core = CoreFalso(nivel)
    monkeypatch.setattr(guardar, "_core", core.namespace)
    df = pd.DataFrame({"Cliente": list("abcd"), "Region": ["N", "S", "N", "S"]})
    v = guardar.abrir_guardar(raiz, df, nombre_sugerido="resultado", **kw)
    _esperar(raiz, lambda: _listo(v))
    return v, core


def test_guardar_pasa_column_types_al_output_spec(raiz, monkeypatch, tmp_path):
    tipos = {"Cliente": "texto", "Region": "texto"}
    v, core = _abrir_guardar(raiz, monkeypatch, nivel="ok", column_types=tipos)
    monkeypatch.setattr(
        guardar.filedialog, "asksaveasfilename", lambda **k: str(tmp_path / "salida.xlsx")
    )
    v.btn_guardar.invoke()
    _esperar(raiz, lambda: v._resultado is not None)
    assert core.guardados[0].column_types == tipos


def test_guardar_sin_column_types_manda_none(raiz, monkeypatch, tmp_path):
    v, core = _abrir_guardar(raiz, monkeypatch, nivel="ok")
    monkeypatch.setattr(
        guardar.filedialog, "asksaveasfilename", lambda **k: str(tmp_path / "salida.xlsx")
    )
    v.btn_guardar.invoke()
    _esperar(raiz, lambda: v._resultado is not None)
    assert core.guardados[0].column_types is None


def test_ventana_guardar_entra_en_768_px(raiz, monkeypatch, tmp_path):
    v, core = _abrir_guardar(raiz, monkeypatch, nivel="obligatorio", claves_sugeridas=["Region"])
    v.win.update_idletasks()
    assert v.win.winfo_reqheight() <= 768
    assert v.btn_guardar.winfo_ismapped()
    # con el resultado a la vista (lo más alto que puede llegar a estar) también entra
    monkeypatch.setattr(
        guardar.filedialog, "asksaveasfilename", lambda **k: str(tmp_path / "salida.xlsx")
    )
    v.v_modo.set("rows")
    v._etapa_b()
    _esperar(raiz, lambda: v._partes is not None)
    v.btn_guardar.invoke()
    _esperar(raiz, lambda: v._resultado is not None)
    v.win.update_idletasks()
    assert v.f_resultado.winfo_ismapped()
    assert v.win.winfo_reqheight() <= 768
    assert v.btn_guardar.winfo_ismapped()
