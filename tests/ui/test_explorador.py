import os
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_tool.core.io import CsvOptions
from excel_tool.ui import app, explorador, selector

from .perfiles_falsos import hojas_falsas, perfil_falso

pytestmark = pytest.mark.pantalla

RUTA = Path("/datos/ventas.csv")


def _textos(widget):
    encontrados = []
    for hijo in widget.winfo_children():
        try:
            encontrados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        encontrados.extend(_textos(hijo))
    return encontrados


def _esperar(raiz, condicion, segundos=10):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        raiz.update()
        if condicion():
            return
        time.sleep(0.02)
    raise AssertionError("se agotó la espera")


@pytest.fixture
def raiz(tmp_path, monkeypatch):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos"))

    def sin_cartel(*args, **kwargs):
        raise AssertionError(f"cartel inesperado: {args}")

    for nombre in ("showerror", "showwarning", "showinfo"):
        monkeypatch.setattr(explorador.messagebox, nombre, sin_cartel)
    # lectura falsa: el perfil también es falso
    leido = SimpleNamespace(sheets=[], df=None)
    monkeypatch.setattr(explorador, "read_table", lambda path, opciones: leido)
    monkeypatch.setattr(explorador, "list_sheets", lambda path: hojas_falsas())
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


def _seleccion(sheet=None):
    return selector.Seleccion(RUTA, sheet, CsvOptions(";", "utf-8-sig"), 0)


def _abrir(raiz, monkeypatch, perfil=None):
    perfil = perfil or perfil_falso()

    def falso(leido, *, on_column=None, progress=None, cancel=None, sheets=None):
        for col in perfil.columns_profile:
            if on_column:
                on_column(col)
        return perfil

    monkeypatch.setattr(explorador, "perfilar", falso)
    ex = explorador.abrir_explorador(raiz, _seleccion())
    _esperar(raiz, lambda: ex.perfil is not None)
    return ex


def test_general_con_formato_argentino(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    textos = _textos(ex.pestana_general)
    assert "ventas.csv" in textos
    assert "12,3 MB" in textos
    assert "05/03/2026 14:30" in textos
    assert "1.234" in textos  # filas
    assert "7" in textos  # repetidas
    assert "Punto y coma ( ; )" in textos
    assert "Recomendado (con acentos y ñ)" in textos
    assert any("Importe tiene 5 valores" in t for t in textos)
    prohibidas = ("dtype", "null", "distinct ", "DataFrame", "encoding")
    assert not any(p in t for t in _textos(ex) for p in prohibidas)
    assert ex.pestanas.tab(0, "text") == "General"
    assert ex.pestanas.tab(1, "text") == "Columnas"


def test_columnas_con_marcas(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    filas = [ex.tabla.item(i, "values") for i in ex.tabla.get_children()]
    assert [f[1] for f in filas] == ["ID cliente", "Importe", "Región"]
    assert filas[0][2] == "Número entero"
    assert filas[0][5] == "10.000"
    assert "Posible clave" in filas[0][6]
    assert "Se puede sumar" in filas[0][6]
    assert filas[1][2] == "Mixto"
    assert filas[2][6] == "Pocos valores"


def test_columnas_aparecen_de_a_una(raiz, monkeypatch):
    perfil = perfil_falso()
    liberar = threading.Event()

    def lento(leido, *, on_column=None, progress=None, cancel=None, sheets=None):
        on_column(perfil.columns_profile[0])
        liberar.wait(5)
        return perfil

    monkeypatch.setattr(explorador, "perfilar", lento)
    ex = explorador.abrir_explorador(raiz, _seleccion())
    _esperar(raiz, lambda: len(ex.tabla.get_children()) == 1)
    assert ex.perfil is None
    liberar.set()
    _esperar(raiz, lambda: ex.perfil is not None)
    assert len(ex.tabla.get_children()) == 3


def test_detalle_mixta_y_variantes(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    ex.seleccionar_columna("Importe")
    textos = _textos(ex.contenido_detalle)
    assert any("Número decimal: 99,5 % (995)" in t for t in textos)
    assert "1.234.567,89" in textos
    assert any("Hay 5 valores que no son números" in t for t in textos)
    assert len(ex.tabla_frecuentes.get_children()) == 2
    ex.seleccionar_columna("Región")
    textos = _textos(ex.contenido_detalle)
    assert any("Norte y NORTE" in t for t in textos)
    assert "Norte, Sur" in textos


def test_boton_ver_distintos_llama_al_gancho(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    ex.pestanas.select(ex.pestana_columnas)
    ex.on_ver_distintos = None
    ex.seleccionar_columna("Región")
    ex.update()
    assert not ex.boton_distintos.winfo_ismapped()  # sin gancho no se ofrece
    pedidas = []
    ex.on_ver_distintos = pedidas.append
    ex.seleccionar_columna("Región")
    ex.update()
    assert ex.boton_distintos.winfo_ismapped()
    ex.boton_distintos.invoke()
    assert pedidas == ["Región"]


def test_ordenar_por_encabezado(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    ex._ordenar("nombre")
    nombres = [ex.tabla.item(i, "values")[1] for i in ex.tabla.get_children()]
    assert nombres == ["ID cliente", "Importe", "Región"]
    ex._ordenar("distintos")
    ex._ordenar("distintos")
    nombres = [ex.tabla.item(i, "values")[1] for i in ex.tabla.get_children()]
    assert nombres[0] == "ID cliente"


def test_cuatro_pestanas(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    nombres = [ex.pestanas.tab(i, "text") for i in range(ex.pestanas.index("end"))]
    assert nombres == ["General", "Columnas", "Valores distintos", "Datos"]


def test_cambiar_de_hoja_vuelve_a_analizar(raiz, monkeypatch):
    pedidas = []

    def falso(leido, *, on_column=None, progress=None, cancel=None, sheets=None):
        hoja = pedidas[-1]
        return perfil_falso(Path("/datos/libro.xlsx"), hojas_falsas(), hoja, csv=False)

    def leer(path, opciones):
        pedidas.append(opciones.sheet or "Ventas")
        return SimpleNamespace(sheets=["Ventas", "Clientes"], df=None)

    monkeypatch.setattr(explorador, "perfilar", falso)
    monkeypatch.setattr(explorador, "read_table", leer)
    ex = explorador.abrir_explorador(
        raiz, selector.Seleccion(Path("/datos/libro.xlsx"), "Ventas", None, 0)
    )
    _esperar(raiz, lambda: ex.perfil is not None)
    assert pedidas == ["Ventas"]
    assert ex.combo_hoja.winfo_ismapped()
    assert "Ventas (100 filas, 3 columnas)" in ex.var_hoja.get()
    anterior = ex.perfil
    ex.cambiar_hoja("Clientes")
    _esperar(raiz, lambda: ex.perfil is not anterior)
    assert pedidas == ["Ventas", "Clientes"]
    assert ex.perfil.sheet == "Clientes"
    assert "Clientes (50 filas, 4 columnas)" in ex.var_hoja.get()


def test_cancelar_cierra_el_explorador_sin_error(raiz, monkeypatch):
    def colgado(leido, *, on_column=None, progress=None, cancel=None, sheets=None):
        while True:
            time.sleep(0.01)
            cancel.check()

    monkeypatch.setattr(explorador, "perfilar", colgado)
    cerradas = []
    ex = explorador.abrir_explorador(raiz, _seleccion(), on_cerrar=lambda: cerradas.append(1))
    _esperar(raiz, lambda: ex.analizando)
    time.sleep(0.2)
    ex.cancelar_analisis()
    _esperar(raiz, lambda: ex.cerrada)
    assert cerradas == [1]
    assert not ex.winfo_exists()


def test_error_muestra_mensaje_y_cierra(raiz, monkeypatch):
    from excel_tool.core.errors import ExcelToolError

    def falla(leido, **kw):
        raise ExcelToolError("El archivo está dañado.")

    mensajes = []
    monkeypatch.setattr(
        explorador.messagebox, "showerror", lambda titulo, msg, **kw: mensajes.append(msg)
    )
    monkeypatch.setattr(explorador, "perfilar", falla)
    ex = explorador.abrir_explorador(raiz, _seleccion())
    _esperar(raiz, lambda: ex.cerrada)
    assert mensajes == ["El archivo está dañado."]


def test_inicio_abre_el_explorador(raiz, monkeypatch):
    monkeypatch.setattr(explorador, "perfilar", lambda leido, **kw: perfil_falso())
    pantalla = app.PantallaInicio(raiz)
    pantalla._cargar(_seleccion())
    _esperar(raiz, lambda: pantalla.explorador.perfil is not None)
    assert pantalla.explorador.winfo_exists()
