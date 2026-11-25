import os
import sys
import tkinter as tk

import pytest

from excel_tool import APP_NAME, __version__
from excel_tool.ui import app


def _textos(widget):
    encontrados = []
    for hijo in widget.winfo_children():
        try:
            encontrados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        encontrados.extend(_textos(hijo))
    return encontrados


@pytest.mark.pantalla
def test_ventana_principal():
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    ventana = app.crear_ventana()
    try:
        ventana.update()
        assert ventana.title() == APP_NAME
        textos = _textos(ventana)
        assert any(__version__ in t for t in textos)
        assert ventana.winfo_width() >= app.ANCHO_MINIMO
    finally:
        ventana.destroy()
