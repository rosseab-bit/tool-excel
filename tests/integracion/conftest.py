"""Entorno para los tests de integración con pantalla (xvfb) y el core real.

Mismo cuidado que `tests/ui/conftest.py` (recolector de basura apagado mientras corre el test,
porque `tk.Variable.__del__` llama a Tcl y no puede correr dentro de un hilo del worker), pero en
vez de fallar ante un cartel, los registra: acá los carteles son parte de lo que se verifica.
"""

from __future__ import annotations

import gc
import os
import sys
import threading
import time
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox

import pytest

_CARTELES_INFO = ("showinfo", "showwarning", "showerror")
_CARTELES_PREGUNTA = (
    "askquestion",
    "askokcancel",
    "askyesno",
    "askyesnocancel",
    "askretrycancel",
)
_DIALOGOS = ("askopenfilename", "askopenfilenames", "asksaveasfilename", "askdirectory")


@dataclass
class EntornoUI:
    raiz: tk.Tk
    carteles: list[tuple[str, str, str]] = field(default_factory=list)  # (tipo, título, texto)
    destino: Path | None = None  # lo que devuelve "Guardar como"

    def esperar(self, condicion: Callable[[], bool], segundos: float = 20.0) -> None:
        limite = time.monotonic() + segundos
        while time.monotonic() < limite:
            self.raiz.update()
            if condicion():
                return
            time.sleep(0.02)
        raise AssertionError("se agotó la espera")

    def bombear(self, segundos: float) -> None:
        limite = time.monotonic() + segundos
        while time.monotonic() < limite:
            self.raiz.update()
            time.sleep(0.02)


TOPE_HILOS_S = 40


@pytest.fixture(scope="session")
def raiz_sesion():
    """Un solo intérprete Tk para toda la sesión (como la app real).

    En Windows crear y destruir muchos intérpretes Tcl en un mismo proceso es frágil (errores de
    init.tcl / tcl_findLibrary). Cada test usa esta raíz oculta como `parent` de sus ventanas.
    """
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    raiz = tk.Tk()
    raiz.withdraw()
    yield raiz
    try:
        raiz.destroy()
    except tk.TclError:
        pass


def _esperar_hilos_worker() -> list[str]:
    limite = time.monotonic() + TOPE_HILOS_S
    for hilo in threading.enumerate():
        if hilo.name == "worker" and hilo is not threading.current_thread():
            hilo.join(max(0.0, limite - time.monotonic()))
    return [h.name for h in threading.enumerate() if h.name == "worker" and h.is_alive()]


def _limpiar_raiz(raiz: tk.Tk) -> None:
    try:
        for id_after in raiz.tk.splitlist(raiz.tk.call("after", "info")):
            raiz.after_cancel(id_after)
        for hija in raiz.winfo_children():
            hija.destroy()
        raiz.update_idletasks()
    except tk.TclError:
        pass


@pytest.fixture
def entorno_ui(tmp_path, monkeypatch, raiz_sesion):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos"))

    entorno_box: dict[str, EntornoUI] = {}

    def cartel(tipo):
        def falso(titulo=None, mensaje=None, *args, **kwargs):
            entorno_box["e"].carteles.append((tipo, str(titulo), str(mensaje)))
            return "ok"

        return falso

    for nombre in _CARTELES_INFO:
        monkeypatch.setattr(messagebox, nombre, cartel(nombre))

    def pregunta(nombre):
        def falso(*args, **kwargs):
            raise AssertionError(f"pregunta inesperada: {nombre}{args}")

        return falso

    for nombre in _CARTELES_PREGUNTA:
        monkeypatch.setattr(messagebox, nombre, pregunta(nombre))

    def guardar_como(*args, **kwargs):
        destino = entorno_box["e"].destino
        assert destino is not None, "se abrió 'Guardar como' sin destino preparado"
        return str(destino)

    def dialogo(nombre):
        def falso(*args, **kwargs):
            raise AssertionError(f"diálogo real inesperado: {nombre}")

        return falso

    for nombre in _DIALOGOS:
        monkeypatch.setattr(filedialog, nombre, dialogo(nombre))
    monkeypatch.setattr(filedialog, "asksaveasfilename", guardar_como)

    gc.collect()
    gc.disable()
    raiz = raiz_sesion
    entorno = EntornoUI(raiz)
    entorno_box["e"] = entorno
    colgados: list[str] = []
    try:
        yield entorno
    finally:
        # Los hilos primero: si siguen vivos tocando Tcl, no se destruye nada debajo de ellos.
        colgados = _esperar_hilos_worker()
        _limpiar_raiz(raiz)
        gc.enable()
        gc.collect()
    if colgados:
        pytest.fail(
            f"el test dejó hilos del worker vivos pasados {TOPE_HILOS_S} s: {colgados}",
            pytrace=False,
        )
