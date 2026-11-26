"""Medición de la Fase 5 (no es un test): recorrido completo del escenario grande por la UI.

    xvfb-run -a /usr/bin/time -v .venv/bin/python -m tests.integracion.medicion_fase5

Maneja el asistente real (200.000 clientes x 2,5 M de ventas, en detalle: avisos, guardar para Excel
y para Google Sheets, "achicar" y resumen final con muestra) y al final imprime el pico de memoria
(`ru_maxrss`). El presupuesto es 1,5 GB (ADR-023). Las salidas van a una carpeta temporal.
"""

from __future__ import annotations

import json
import os
import resource
import shutil
import sys
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ))

import pytest  # noqa: E402

from tests.integracion.asistente_util import Conductor  # noqa: E402
from tests.integracion.conftest import EntornoUI  # noqa: E402
from tests.integracion.escenario_grande_ui import cargar_esperado, recorrer  # noqa: E402

GRANDES = RAIZ / "tests" / "fixtures" / "generados" / "grandes"
ESPERADOS = RAIZ / "tests" / "fixtures" / "esperados"


def main() -> None:
    datos = Path(tempfile.mkdtemp(prefix="medicion_fase5_datos_"))
    salida = Path(tempfile.mkdtemp(prefix="medicion_fase5_"))
    os.environ["XDG_DATA_HOME"] = str(datos)  # preferencias propias, no las del usuario
    monkeypatch = pytest.MonkeyPatch()
    raiz = tk.Tk()
    raiz.withdraw()
    entorno = EntornoUI(raiz)

    def cartel(tipo):
        def falso(titulo=None, mensaje=None, *a, **k):
            entorno.carteles.append((tipo, str(titulo), str(mensaje)))
            return "ok"

        return falso

    for nombre in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(messagebox, nombre, cartel(nombre))
    monkeypatch.setattr(filedialog, "asksaveasfilename", lambda *a, **k: str(entorno.destino))
    try:
        c = Conductor(entorno, monkeypatch)
        hechos = recorrer(c, GRANDES, salida, cargar_esperado(ESPERADOS))
    finally:
        monkeypatch.undo()
        shutil.rmtree(salida, ignore_errors=True)
        shutil.rmtree(datos, ignore_errors=True)
    pico_mb = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024
    print("PICO_MB", round(pico_mb))
    print("HECHOS " + json.dumps(hechos, ensure_ascii=False, default=str))


if __name__ == "__main__":
    main()
