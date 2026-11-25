import os
import sys
import tkinter as tk

import pandas as pd
import pytest

from excel_tool.ui.grilla import COLUMNA_FILA, Grilla, texto_celda

pytestmark = pytest.mark.pantalla


@pytest.fixture
def raiz():
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    r = tk.Tk()
    yield r
    r.destroy()


def proveedor_falso(total, page_size=10):
    def prov(pagina):
        desde = pagina * page_size
        n = max(0, min(page_size, total - desde))
        df = pd.DataFrame(
            {
                COLUMNA_FILA: range(desde + 1, desde + n + 1),
                "Nombre": [f"x{i}" for i in range(n)],
                "Monto": ["" if i == 0 else str(i) for i in range(n)],
            }
        )
        return df, total

    return prov


def test_navegacion_y_posicion(raiz):
    g = Grilla(raiz, proveedor_falso(25), page_size=10)
    g.pack()
    g.ir_a(0)
    assert g.texto_posicion == "Filas 1 a 10 de 25"
    g.siguiente()
    assert g.texto_posicion == "Filas 11 a 20 de 25"
    g.ultima()
    assert g.texto_posicion == "Filas 21 a 25 de 25"
    assert g.filas_mostradas() == 5
    g.siguiente()
    assert g.pagina == 2
    g.anterior()
    assert g.pagina == 1
    g.primera()
    assert g.pagina == 0


def test_numeros_con_separador_de_miles(raiz):
    g = Grilla(raiz, proveedor_falso(2_000_000, 1000), page_size=1000)
    g.ir_a(1)
    assert g.texto_posicion == "Filas 1.001 a 2.000 de 2.000.000"


def test_nunca_mas_de_una_pagina(raiz):
    g = Grilla(raiz, proveedor_falso(95), page_size=10)
    for p in range(10):
        g.ir_a(p)
        assert g.filas_mostradas() <= 10


def test_n_fila_a_la_izquierda_y_celdas_vacias(raiz):
    def prov(pagina):
        return pd.DataFrame({"A": ["a", None], COLUMNA_FILA: [1, 2]}), 2

    g = Grilla(raiz, prov, page_size=10)
    g.ir_a(0)
    assert g.columnas[0] == COLUMNA_FILA
    assert g.arbol.item(g.arbol.get_children()[1], "values")[1] == ""
    assert texto_celda(float("nan")) == ""
    assert len(texto_celda("x" * 1000)) < 200


def test_on_sort(raiz):
    pedidos = []
    g = Grilla(raiz, proveedor_falso(5), on_sort=lambda c, a: pedidos.append((c, a)))
    g.ir_a(0)
    g._clic_encabezado("Nombre")
    g._clic_encabezado("Nombre")
    g._clic_encabezado("Monto")
    assert pedidos == [("Nombre", True), ("Nombre", False), ("Monto", True)]


def test_ocultar_y_mostrar_columnas(raiz):
    g = Grilla(raiz, proveedor_falso(5))
    g.ir_a(0)
    g.ocultar_columna("Monto")
    g.ocultar_columna(COLUMNA_FILA)  # no se puede
    assert g.columnas_visibles == [COLUMNA_FILA, "Nombre"]
    assert len(g.arbol.cget("displaycolumns")) == 2
    g.siguiente()
    g.mostrar_columna("Monto")
    assert g.columnas_visibles == [COLUMNA_FILA, "Nombre", "Monto"]


def test_cero_filas(raiz):
    g = Grilla(raiz, proveedor_falso(0))
    g.ir_a(0)
    assert g.filas_mostradas() == 0
    assert g.texto_posicion == "Sin filas para mostrar"
    g.siguiente()
    g.ultima()
    assert g.pagina == 0


def test_cambiar_proveedor(raiz):
    g = Grilla(raiz, proveedor_falso(30), page_size=10)
    g.ir_a(2)
    g.set_proveedor(proveedor_falso(5))
    assert g.pagina == 0 and g.total_filas == 5
    g.set_proveedor(None)
    assert g.filas_mostradas() == 0
