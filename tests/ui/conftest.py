"""Limpieza común de los tests de UI.

Causa del cuelgue de `test_cambiar_de_hoja_vuelve_a_analizar` (60 s) y del core dump del
selector: `tk.Variable.__del__` llama a Tcl. Si el recolector de basura corre dentro de un hilo
del worker, Tcl se queda esperando un `mainloop` que no existe y el test se cuelga o el proceso
aborta. Por eso el recolector queda apagado mientras corre cada test y se ejecuta recién al
terminar, en el hilo principal.
"""

from __future__ import annotations

import gc
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import pytest

_CARTELES = (
    "showinfo",
    "showwarning",
    "showerror",
    "askquestion",
    "askokcancel",
    "askyesno",
    "askyesnocancel",
    "askretrycancel",
)
_DIALOGOS = ("askopenfilename", "askopenfilenames", "asksaveasfilename", "askdirectory")
TOPE_HILOS_S = 5


def _cartel_inesperado(nombre):
    def falso(*args, **kwargs):
        raise AssertionError(f"cartel o diálogo real inesperado: {nombre}{args}")

    return falso


def _destruir(ventana: tk.Tk) -> None:
    try:
        for id_after in ventana.tk.splitlist(ventana.tk.call("after", "info")):
            ventana.after_cancel(id_after)
        for hija in ventana.winfo_children():
            hija.destroy()
        ventana.destroy()
    except tk.TclError:
        pass


@pytest.fixture(autouse=True)
def limpieza_ui(monkeypatch):
    ventanas: list[tk.Tk] = []
    original = tk.Tk.__init__

    def registrar(self, *args, **kwargs):
        original(self, *args, **kwargs)
        ventanas.append(self)

    monkeypatch.setattr(tk.Tk, "__init__", registrar)
    for nombre in _CARTELES:
        monkeypatch.setattr(messagebox, nombre, _cartel_inesperado(nombre))
    for nombre in _DIALOGOS:
        monkeypatch.setattr(filedialog, nombre, _cartel_inesperado(nombre))

    gc.collect()
    gc.disable()
    try:
        yield
    finally:
        for ventana in ventanas:
            _destruir(ventana)
        for hilo in threading.enumerate():
            if hilo.name == "worker" and hilo is not threading.current_thread():
                hilo.join(TOPE_HILOS_S)
        gc.enable()
        gc.collect()
