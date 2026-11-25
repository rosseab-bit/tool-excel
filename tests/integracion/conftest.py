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


@pytest.fixture
def entorno_ui(tmp_path, monkeypatch):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos"))

    ventanas: list[tk.Tk] = []
    original = tk.Tk.__init__

    def registrar(self, *args, **kwargs):
        original(self, *args, **kwargs)
        ventanas.append(self)

    monkeypatch.setattr(tk.Tk, "__init__", registrar)

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
    try:
        raiz = tk.Tk()
        entorno = EntornoUI(raiz)
        entorno_box["e"] = entorno
        yield entorno
    finally:
        for ventana in ventanas:
            try:
                for id_after in ventana.tk.splitlist(ventana.tk.call("after", "info")):
                    ventana.after_cancel(id_after)
                for hija in ventana.winfo_children():
                    hija.destroy()
                ventana.destroy()
            except tk.TclError:
                pass
        for hilo in threading.enumerate():
            if hilo.name == "worker" and hilo is not threading.current_thread():
                hilo.join(40)
        gc.enable()
        gc.collect()
