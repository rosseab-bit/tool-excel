import logging
import os
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox

import pytest

from excel_tool.core.errors import ArchivoIlegible
from excel_tool.core.progress import ProgressEvent
from excel_tool.ui import worker
from excel_tool.ui.worker import MENSAJE_INESPERADO, run_in_background

pytestmark = pytest.mark.pantalla


@pytest.fixture
def raiz():
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    r = tk.Tk()
    yield r
    r.destroy()


@pytest.fixture
def mensajes(monkeypatch):
    lista = []
    monkeypatch.setattr(messagebox, "showerror", lambda *a, **k: lista.append((a, k)))
    return lista


def esperar(raiz, condicion, timeout=5.0):
    fin = time.monotonic() + timeout
    while not condicion() and time.monotonic() < fin:
        raiz.update()
        time.sleep(0.01)
    return condicion()


def ventanas(raiz):
    return [w for w in raiz.winfo_children() if isinstance(w, tk.Toplevel)]


def test_exito_con_resultado_en_hilo_principal(raiz, mensajes):
    res = []
    hilos = []

    def func(progress, cancel):
        hilos.append(threading.current_thread())
        return 42

    def ok(r):
        res.append(r)
        hilos.append(threading.current_thread())

    run_in_background(raiz, func, title="Prueba", on_success=ok)
    assert esperar(raiz, lambda: res)
    assert res == [42]
    assert hilos[0] is not threading.main_thread()
    assert hilos[1] is threading.main_thread()
    assert not mensajes
    assert not ventanas(raiz)


def test_progreso_determinado_e_indeterminado(raiz, mensajes):
    res = []
    paso = threading.Event()
    seguir = threading.Event()

    def func(progress, cancel):
        progress(ProgressEvent("Leyendo", 3, 10))
        paso.set()
        seguir.wait(5)
        progress(ProgressEvent("Pensando", 0, None))
        paso.clear()
        seguir.clear()
        paso.set()
        seguir.wait(5)
        return "listo"

    run_in_background(raiz, func, title="Prueba", on_success=res.append)
    assert esperar(raiz, lambda: ventanas(raiz) and paso.is_set())
    win = ventanas(raiz)[0]

    def barra_modo():
        raiz.update()
        for w in win.winfo_children()[0].winfo_children():
            if w.winfo_class() == "TProgressbar":
                return w, str(w.cget("mode")), w.cget("value"), w.cget("maximum")
        return None

    assert esperar(raiz, lambda: barra_modo() and barra_modo()[1] == "determinate")
    _, modo, valor, maximo = barra_modo()
    assert (float(valor), float(maximo)) == (3.0, 10.0)
    paso.clear()
    seguir.set()
    assert esperar(raiz, lambda: barra_modo() and barra_modo()[1] == "indeterminate")
    seguir.set()
    assert esperar(raiz, lambda: res)
    assert res == ["listo"]


def test_cancelar_termina_rapido(raiz, mensajes):
    res, errs = [], []
    arrancado = threading.Event()

    def func(progress, cancel):
        arrancado.set()
        while True:
            cancel.check()
            time.sleep(0.01)

    token = run_in_background(
        raiz, func, title="Prueba", on_success=res.append, on_error=errs.append
    )
    assert esperar(raiz, arrancado.is_set)
    assert esperar(raiz, lambda: ventanas(raiz))
    inicio = time.monotonic()
    boton = None
    for w in ventanas(raiz)[0].winfo_children()[0].winfo_children():
        if w.winfo_class() == "TButton":
            boton = w
    boton.invoke()
    assert token.cancelled
    assert not ventanas(raiz)
    assert esperar(raiz, lambda: not ventanas(raiz), timeout=2)
    assert time.monotonic() - inicio <= 2
    assert not res and not errs and not mensajes


def test_cerrar_con_la_x_cancela(raiz, mensajes):
    arrancado = threading.Event()

    def func(progress, cancel):
        arrancado.set()
        while True:
            cancel.check()
            time.sleep(0.01)

    token = run_in_background(raiz, func, title="Prueba", on_success=lambda r: None)
    assert esperar(raiz, lambda: ventanas(raiz))
    raiz.tk.call(ventanas(raiz)[0].protocol("WM_DELETE_WINDOW"))
    assert token.cancelled
    assert esperar(raiz, lambda: not ventanas(raiz), timeout=2)


def test_error_del_core_muestra_su_mensaje(raiz, mensajes):
    def func(progress, cancel):
        raise ArchivoIlegible("El archivo está dañado.")

    run_in_background(raiz, func, title="Abrir", on_success=lambda r: None)
    assert esperar(raiz, lambda: mensajes)
    assert mensajes[0][0][1] == "El archivo está dañado."


def test_error_inesperado_mensaje_generico_y_log(raiz, mensajes, caplog):
    def func(progress, cancel):
        raise ValueError("detalle-secreto")

    with caplog.at_level(logging.ERROR, logger=worker.__name__):
        run_in_background(raiz, func, title="Abrir", on_success=lambda r: None)
        assert esperar(raiz, lambda: mensajes)
    assert mensajes[0][0][1] == MENSAJE_INESPERADO
    assert "detalle-secreto" not in mensajes[0][0][1]
    registro = [r for r in caplog.records if r.exc_info]
    assert registro and registro[0].exc_info[0] is ValueError
    assert "Traceback" in caplog.text


def test_on_error_reemplaza_el_mensaje(raiz, mensajes):
    errs = []
    hilos = []

    def on_error(e):
        errs.append(e)
        hilos.append(threading.current_thread())

    def func(progress, cancel):
        raise ArchivoIlegible("x")

    run_in_background(raiz, func, title="Abrir", on_success=lambda r: None, on_error=on_error)
    assert esperar(raiz, lambda: errs)
    assert isinstance(errs[0], ArchivoIlegible)
    assert hilos[0] is threading.main_thread()
    assert not mensajes


def test_operacion_corta_no_muestra_ventana(raiz, mensajes):
    res = []
    run_in_background(raiz, lambda p, c: 1, title="Prueba", on_success=res.append)
    visto = []
    assert esperar(raiz, lambda: (visto.append(bool(ventanas(raiz))), bool(res))[1])
    assert not any(visto)


def test_cerrar_la_raiz_cancela_sin_romper(mensajes):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    r = tk.Tk()
    arrancado = threading.Event()
    res = []

    def func(progress, cancel):
        arrancado.set()
        while True:
            cancel.check()
            time.sleep(0.01)

    token = run_in_background(r, func, title="Prueba", on_success=res.append)
    assert esperar(r, arrancado.is_set)
    r.destroy()
    assert token.cancelled
    time.sleep(0.2)
    assert not res and not mensajes


def test_cancelar_responde_al_instante_aunque_la_funcion_lo_ignore(raiz, mensajes):
    ok, canceladas, errores = [], [], []

    def terca(progress, cancel):
        time.sleep(0.5)  # ignora el cancel
        raise ArchivoIlegible("falla tarde")

    run_in_background(
        raiz,
        terca,
        title="Prueba",
        on_success=ok.append,
        on_error=errores.append,
        on_cancel=lambda: canceladas.append(1),
    )
    assert esperar(raiz, lambda: ventanas(raiz))
    boton = _boton_cancelar(ventanas(raiz)[0])
    t0 = time.monotonic()
    boton.invoke()
    assert esperar(raiz, lambda: not ventanas(raiz), timeout=0.5)
    assert time.monotonic() - t0 < 0.5
    assert raiz.grab_current() is None
    # Esperar a que el hilo termine con su excepción.
    time.sleep(0.7)
    for _ in range(10):
        raiz.update()
        time.sleep(0.02)
    assert canceladas == [1]
    assert ok == [] and errores == []
    assert mensajes == []


def _boton_cancelar(ventana):
    pendientes = [ventana]
    while pendientes:
        w = pendientes.pop()
        if w.winfo_class() == "TButton":
            return w
        pendientes.extend(w.winfo_children())
    raise AssertionError("no hay botón Cancelar")


def test_cancelar_con_la_x_y_exito_tardio_no_llama_a_on_success(raiz, mensajes):
    ok, canceladas = [], []

    def terca(progress, cancel):
        time.sleep(0.9)
        return 1

    run_in_background(
        raiz, terca, title="Prueba", on_success=ok.append, on_cancel=lambda: canceladas.append(1)
    )
    assert esperar(raiz, lambda: ventanas(raiz))
    ventanas(raiz)[0].tk.call(ventanas(raiz)[0].protocol("WM_DELETE_WINDOW"))
    assert esperar(raiz, lambda: not ventanas(raiz), timeout=0.5)
    time.sleep(1.1)
    for _ in range(10):
        raiz.update()
        time.sleep(0.02)
    assert canceladas == [1]
    assert ok == []
