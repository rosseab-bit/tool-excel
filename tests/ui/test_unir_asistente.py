"""Tests con pantalla (xvfb) del asistente y su entrada desde el inicio (F5-10)."""

from __future__ import annotations

import os
import shutil
import sys
import time
import tkinter as tk
from tkinter import ttk

import pytest

from excel_tool.core.recipe import save_recipe
from excel_tool.ui import app
from excel_tool.ui.unir import asistente as mod_asistente
from excel_tool.ui.unir.asistente import ALTO_MAXIMO_VENTANA, Asistente
from excel_tool.ui.unir.estado import EstadoUnion, Fuente
from excel_tool.ui.unir.pasos import PASOS

from .test_unir_estado import _estado

pytestmark = pytest.mark.pantalla


@pytest.fixture
def raiz(tmp_path, monkeypatch):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos_app"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos_app"))
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


def _textos(widget):
    encontrados = []
    for hijo in widget.winfo_children():
        try:
            encontrados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        encontrados.extend(_textos(hijo))
    return encontrados


def _esperar(raiz, condicion, segundos=15):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        raiz.update()
        if condicion():
            return
        time.sleep(0.02)
    raise AssertionError("se agotó la espera")


def test_siete_pasos_con_sus_titulos(raiz):
    asistente = Asistente(raiz)
    raiz.update()
    assert [p.titulo for p in asistente.pasos] == [
        "Archivo principal",
        "Columna clave",
        "Archivos a sumar",
        "¿Detalle o resumen?",
        "Columnas",
        "Resumen final",
        "Revisar y guardar",
    ]
    assert len(PASOS) == 7
    assert asistente.indice == 0
    assert str(asistente.boton_anterior.cget("text")) == "Anterior"


def test_avanzar_y_retroceder_sin_perder_nada(raiz, monkeypatch):
    asistente = Asistente(raiz)
    raiz.update()
    for paso in asistente.pasos:  # acá solo se prueba la navegación, no lo que valida cada paso
        monkeypatch.setattr(paso, "validar", lambda: None)
    for esperado in range(1, 7):
        asistente.boton_siguiente.invoke()
        assert asistente.indice == esperado
    assert not asistente.boton_siguiente.winfo_ismapped()
    marcas = [e.cget("text") for e in asistente._etiquetas]
    assert marcas[0].startswith("✓") and not marcas[6].startswith("✓")
    asistente.pasos[2].dato = "queda"
    for _ in range(4):
        asistente.boton_anterior.invoke()
    assert asistente.indice == 2
    assert asistente.paso_actual.dato == "queda"
    asistente.ir_a(0)
    assert "disabled" in asistente.boton_anterior.state()


def test_validar_bloquea_y_muestra_el_motivo(raiz, monkeypatch):
    asistente = Asistente(raiz)
    raiz.update()
    llamadas = []
    paso = asistente.pasos[0]
    monkeypatch.setattr(paso, "validar", lambda: "Elegí un archivo.")
    monkeypatch.setattr(paso, "al_salir", lambda: llamadas.append("salir"))
    monkeypatch.setattr(asistente.pasos[1], "al_entrar", lambda: llamadas.append("entrar"))
    asistente.boton_siguiente.invoke()
    assert asistente.indice == 0 and llamadas == []
    assert asistente.estado_texto.cget("text") == "Elegí un archivo."
    monkeypatch.setattr(paso, "validar", lambda: None)
    asistente.boton_siguiente.invoke()
    assert asistente.indice == 1 and llamadas == ["salir", "entrar"]
    assert asistente.estado_texto.cget("text") == ""


def test_cancelar_sin_datos_cierra_sin_preguntar(raiz):
    asistente = Asistente(raiz)
    raiz.update()
    asistente.boton_cancelar.invoke()
    assert asistente.cerrada


def test_cancelar_con_datos_pide_confirmacion(raiz, monkeypatch, fixtures_generados):
    base = fixtures_generados / "escenario_3"
    estado = EstadoUnion(principal=Fuente(_estado(base).principal.spec))
    asistente = Asistente(raiz, estado)
    raiz.update()
    preguntas = []
    monkeypatch.setattr(
        mod_asistente.messagebox, "askyesno", lambda *a, **k: preguntas.append(a) or False
    )
    asistente.boton_cancelar.invoke()
    assert preguntas and not asistente.cerrada
    monkeypatch.setattr(mod_asistente.messagebox, "askyesno", lambda *a, **k: True)
    asistente.boton_cancelar.invoke()
    assert asistente.cerrada


def test_la_ventana_entra_en_768(raiz):
    asistente = Asistente(raiz)
    raiz.update()
    assert asistente.winfo_height() <= ALTO_MAXIMO_VENTANA
    assert asistente.winfo_reqheight() <= ALTO_MAXIMO_VENTANA
    # con un paso muy alto, el cuerpo hace scroll y el pie sigue a la vista
    for i in range(60):
        ttk.Label(asistente.pasos[0], text=f"fila {i}").pack()
    raiz.update()
    assert asistente.winfo_reqheight() <= ALTO_MAXIMO_VENTANA
    assert asistente.boton_siguiente.winfo_ismapped()


def test_unir_archivos_desde_el_inicio(raiz):
    pantalla = app.PantallaInicio(raiz)
    raiz.update()
    boton = next(b for b in pantalla.botones if b.cget("text") == "Unir archivos")
    assert "disabled" not in boton.state()
    boton.invoke()
    raiz.update()
    assert isinstance(pantalla.asistente, Asistente)
    assert pantalla.asistente.indice == 0
    assert not any("Próximamente" in t for t in _textos(pantalla)[:6])


def _boton_receta(pantalla):
    return next(b for b in pantalla.botones if b.cget("text") == "Abrir una receta…")


def test_abrir_receta_completa(raiz, monkeypatch, fixtures_generados, tmp_path):
    carpeta = tmp_path / "datos"
    shutil.copytree(fixtures_generados / "escenario_3", carpeta)
    receta = carpeta / "mes.json"
    save_recipe(_estado(carpeta).plan(), receta)
    monkeypatch.setattr(app.filedialog, "askopenfilename", lambda **k: str(receta))
    pantalla = app.PantallaInicio(raiz)
    _boton_receta(pantalla).invoke()
    _esperar(raiz, lambda: pantalla.asistente is not None)
    asistente = pantalla.asistente
    assert asistente.indice == 2
    assert asistente.estado.faltantes == []
    assert [s.mode for s in asistente.estado.secundarios] == ["detail", "summary"]


def test_abrir_receta_con_archivos_faltantes_lista_cuales(
    raiz, monkeypatch, fixtures_generados, tmp_path
):
    carpeta = tmp_path / "datos"
    shutil.copytree(fixtures_generados / "escenario_3", carpeta)
    receta = carpeta / "mes.json"
    save_recipe(_estado(carpeta).plan(), receta)
    (carpeta / "ventas.csv").unlink()
    avisos = []
    monkeypatch.setattr(app.filedialog, "askopenfilename", lambda **k: str(receta))
    monkeypatch.setattr(app.messagebox, "showinfo", lambda *a, **k: avisos.append(a))
    pantalla = app.PantallaInicio(raiz)
    _boton_receta(pantalla).invoke()
    _esperar(raiz, lambda: pantalla.asistente is not None)
    assert pantalla.asistente.estado.faltantes == ["ventas.csv"]
    assert avisos and "ventas.csv" in avisos[0][1]


def test_abrir_receta_cancelando_el_dialogo_no_hace_nada(raiz, monkeypatch):
    monkeypatch.setattr(app.filedialog, "askopenfilename", lambda **k: "")
    pantalla = app.PantallaInicio(raiz)
    _boton_receta(pantalla).invoke()
    raiz.update()
    assert pantalla.asistente is None


def test_receta_rota_muestra_mensaje_amigable(raiz, monkeypatch, tmp_path):
    mala = tmp_path / "rota.json"
    mala.write_text("{ no es json", encoding="utf-8")
    errores = []
    monkeypatch.setattr(app.filedialog, "askopenfilename", lambda **k: str(mala))
    monkeypatch.setattr(app.messagebox, "showerror", lambda *a, **k: errores.append(a))
    pantalla = app.PantallaInicio(raiz)
    _boton_receta(pantalla).invoke()
    _esperar(raiz, lambda: bool(errores))
    assert "dañada" in errores[0][1]
    assert pantalla.asistente is None
