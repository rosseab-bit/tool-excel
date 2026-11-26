import os
import sys
import time
import tkinter as tk
from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_tool.core.io import CsvOptions
from excel_tool.core.preferences import load_preferences
from excel_tool.ui import app, explorador, selector

from .perfiles_falsos import perfil_falso

pytestmark = pytest.mark.pantalla

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "generados"
CSV = FIXTURES / "csv" / "excel_ar.csv"
XLSX = FIXTURES / "formatos" / "varias_hojas.xlsx"


def _textos(widget):
    encontrados = []
    for hijo in widget.winfo_children():
        try:
            encontrados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        encontrados.extend(_textos(hijo))
    return encontrados


@pytest.fixture
def raiz(tmp_path, monkeypatch):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos"))
    monkeypatch.setattr(selector, "comprobar_memoria", lambda path: (True, ""))

    def sin_cartel(*args, **kwargs):
        raise AssertionError(f"cartel inesperado: {args}")

    for nombre in ("showerror", "showwarning", "showinfo"):
        monkeypatch.setattr(selector.messagebox, nombre, sin_cartel)
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


def _esperar(raiz, condicion, segundos=15):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        raiz.update()
        if condicion():
            return
        time.sleep(0.02)
    raise AssertionError("se agotó la espera")


def _abrir(raiz, path, aceptadas):
    sel = selector.abrir_selector(raiz, aceptadas.append, path)
    _esperar(raiz, lambda: sel.listo)
    return sel


def test_inicio_cinco_botones(raiz):
    pantalla = app.PantallaInicio(raiz)
    raiz.update()
    assert len(pantalla.botones) == 5
    estados = [b.instate(["disabled"]) for b in pantalla.botones]
    assert estados == [False, False, False, True, True]
    textos = _textos(pantalla)
    assert "Ver información de un archivo" in textos
    assert sum(app.PROXIMAMENTE in t for t in textos) == 2


def test_selector_csv_y_preferencias(raiz):
    if not CSV.exists():
        pytest.skip("faltan los archivos de prueba (scripts/make_fixtures.py)")
    aceptadas = []
    sel = _abrir(raiz, CSV, aceptadas)
    assert sel.grilla.filas_mostradas() > 0
    assert sel.var_sep.get() == selector.SEPARADORES[";"]
    assert not sel.fila_hoja.winfo_ismapped()
    sel.aceptar()
    assert len(aceptadas) == 1
    eleccion = aceptadas[0]
    assert eleccion.path == CSV
    assert eleccion.csv_options.separator == ";"
    assert eleccion.header_row == 0
    prefs = load_preferences()
    assert prefs.last_folder == CSV.parent
    assert prefs.last_csv_options == CsvOptions(
        separator=";", encoding=eleccion.csv_options.encoding
    )


def test_cambiar_separador_refresca(raiz):
    if not CSV.exists():
        pytest.skip("faltan los archivos de prueba (scripts/make_fixtures.py)")
    sel = _abrir(raiz, CSV, [])
    antes = sel.grilla.columnas
    sel.seleccionar_separador(",")
    assert sel.ocupado
    _esperar(raiz, lambda: sel.listo)
    assert sel.grilla.columnas != antes
    assert sel.seleccion().csv_options.separator == ","


def test_selector_xlsx_varias_hojas(raiz):
    if not XLSX.exists():
        pytest.skip("faltan los archivos de prueba (scripts/make_fixtures.py)")
    aceptadas = []
    sel = _abrir(raiz, XLSX, aceptadas)
    valores = list(sel.combo_hoja.cget("values"))
    assert len(valores) == 3
    assert "filas" in valores[0] and "columnas" in valores[0]
    assert sel.fila_hoja.winfo_ismapped()
    sel.seleccionar_hoja("Ventas")
    _esperar(raiz, lambda: sel.listo)
    sel.aceptar()
    assert aceptadas[0].sheet == "Ventas"
    assert aceptadas[0].csv_options is None


def test_cartel_de_memoria(raiz, monkeypatch):
    avisos = []
    monkeypatch.setattr(selector, "comprobar_memoria", lambda p: (False, "No entra en memoria"))
    monkeypatch.setattr(selector.messagebox, "showwarning", lambda *a, **k: avisos.append(a[1]))
    leidos = []
    monkeypatch.setattr(selector, "read_table", lambda *a, **k: leidos.append(a))
    sel = selector.abrir_selector(raiz, lambda s: None, Path("grande.csv"))
    _esperar(raiz, lambda: sel.cerrada)
    assert avisos == ["No entra en memoria"]
    assert leidos == []


def test_comprobar_memoria_usa_core(monkeypatch):
    modulo = SimpleNamespace(
        check_memory=lambda p: SimpleNamespace(ok=False, message="poca memoria")
    )
    monkeypatch.setitem(sys.modules, "excel_tool.core.memory", modulo)
    assert selector.comprobar_memoria(Path("x.csv")) == (False, "poca memoria")


def test_pedir_archivo_usa_ultima_carpeta(raiz, tmp_path, monkeypatch):
    from excel_tool.core.preferences import Preferences, save_preferences

    save_preferences(Preferences(last_folder=tmp_path))
    capturado = {}

    def falso(**kw):
        capturado.update(kw)
        return ""

    monkeypatch.setattr(selector.filedialog, "askopenfilename", falso)
    assert selector.pedir_archivo(raiz) is None
    assert capturado["initialdir"] == str(tmp_path)
    nombres = [t[0] for t in capturado["filetypes"]]
    assert "Todos los archivos" in nombres


def test_ver_informacion_carga_archivo(raiz, monkeypatch):
    if not CSV.exists():
        pytest.skip("faltan los archivos de prueba (scripts/make_fixtures.py)")
    monkeypatch.setattr(selector.filedialog, "askopenfilename", lambda **kw: str(CSV))
    monkeypatch.setattr(explorador, "perfilar", lambda leido, **kw: perfil_falso(CSV))
    pantalla = app.PantallaInicio(raiz)
    pantalla.botones[0].invoke()
    _esperar(raiz, lambda: pantalla.selector is not None and pantalla.selector.listo)
    pantalla.selector.aceptar()
    _esperar(raiz, lambda: pantalla.explorador is not None and pantalla.explorador.perfil)
    assert pantalla.explorador.perfil.rows == 1234
