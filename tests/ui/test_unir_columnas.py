"""Tests con pantalla (xvfb) del paso 5 del asistente: Columnas (F5-13), con el core real."""

from __future__ import annotations

import os
import shutil
import sys
import time
import tkinter as tk
from tkinter import ttk

import pytest

from excel_tool.core.join import run_join
from excel_tool.ui.unir.asistente import Asistente
from excel_tool.ui.unir.estado import EstadoUnion
from excel_tool.ui.unir.pasos.paso5_columnas import (
    SIN_COLUMNAS,
    PasoColumnas,
    texto_celdas,
    texto_tamano,
)

from .test_unir_estado import SUMA, _estado, _fuente, _sec

pytestmark = pytest.mark.pantalla


@pytest.fixture
def raiz():
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


@pytest.fixture
def carpeta(fixtures_generados, tmp_path):
    destino = tmp_path / "datos"
    shutil.copytree(fixtures_generados / "escenario_3", destino)
    return destino


def _esperar(raiz, condicion, segundos=15):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        raiz.update()
        if condicion():
            return
        time.sleep(0.02)
    raise AssertionError("se agotó la espera")


def _textos(widget):
    encontrados = []
    for hijo in widget.winfo_children():
        try:
            encontrados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        encontrados.extend(_textos(hijo))
    return encontrados


def _abrir(raiz, estado) -> PasoColumnas:
    asistente = Asistente(raiz, estado, paso_inicial=4)
    paso = asistente.pasos[4]
    paso.espera_ms = 10
    return paso


def _listo(raiz, paso):
    _esperar(raiz, lambda: paso.calculo is not None and paso.version_mostrada == paso.version)


def _casillas(paso, archivo):
    seccion = next(s for s in paso.secciones if s.nombre == archivo)
    return seccion.lista


def _estado_todo(carpeta, modo_cobranzas="detail") -> EstadoUnion:
    """Los tres archivos con todas las columnas (sin selección previa)."""
    estado = EstadoUnion(
        principal=_fuente(carpeta, "clientes.xlsx", []),
        claves=["ID cliente"],
        secundarios=[
            _sec(carpeta, "ventas.csv", [("ID cliente", "ID cliente")], [], mode="detail"),
            _sec(
                carpeta,
                "cobranzas.xlsx",
                [("ID cliente", "Nro Cliente")],
                [],
                mode=modo_cobranzas,
                summary=SUMA if modo_cobranzas == "summary" else None,
            ),
        ],
    )
    return estado


def test_una_seccion_por_archivo_y_la_clave_una_sola_vez(raiz, carpeta):
    paso = _abrir(raiz, _estado(carpeta))
    _listo(raiz, paso)
    assert [s.nombre for s in paso.secciones] == ["clientes.xlsx", "ventas.csv", "cobranzas.xlsx"]
    principal = _casillas(paso, "clientes.xlsx")
    assert principal.columnas[0] == "ID cliente"
    assert "ID cliente" in principal.fijas
    assert str(principal.casillas["ID cliente"].cget("text")).endswith("(clave)")
    assert "ID cliente" not in _casillas(paso, "ventas.csv").columnas
    clave_final = [t for t in paso.lista_final.textos() if "ID cliente" in t or "Nro" in t]
    assert clave_final == ["ID cliente"]


def test_el_resumen_muestra_sus_columnas(raiz, carpeta):
    paso = _abrir(raiz, _estado(carpeta))
    _listo(raiz, paso)
    cobranzas = next(s for s in paso.secciones if s.nombre == "cobranzas.xlsx")
    assert cobranzas.lista is None
    assert cobranzas.disponibles == ["Total cobrado"]
    assert "Total cobrado" in paso.lista_final.textos()
    assert "•  Total cobrado" in _textos(paso)


def test_por_defecto_viene_todo_marcado_y_respeta_la_seleccion_previa(raiz, carpeta):
    paso = _abrir(raiz, _estado_todo(carpeta))
    assert _casillas(paso, "ventas.csv").marcadas() == [
        "Fecha",
        "Artículo",
        "Cantidad",
        "Importe",
    ]
    assert len(_casillas(paso, "clientes.xlsx").marcadas()) == 4
    previo = _estado(carpeta)  # su selección viene de las columnas del archivo
    paso2 = _abrir(raiz, previo)
    assert _casillas(paso2, "clientes.xlsx").marcadas() == ["ID cliente", "Nombre", "Región"]
    previo.columnas["ventas.csv"] = ["Importe"]
    paso2.al_entrar()
    assert _casillas(paso2, "ventas.csv").marcadas() == ["Importe"]


def test_nombres_finales_de_columnas_repetidas(raiz, carpeta):
    paso = _abrir(raiz, _estado_todo(carpeta))
    _listo(raiz, paso)
    textos = paso.lista_final.textos()
    assert "Fecha" in textos
    assert "Fecha (cobranzas)" in textos
    assert textos.count("Fecha") == 1


def test_el_indicador_baja_al_desmarcar(raiz, carpeta):
    paso = _abrir(raiz, _estado_todo(carpeta))
    _listo(raiz, paso)
    antes = paso.calculo.diagnostico
    assert paso.etiqueta_tamano.cget("text") == texto_tamano(
        antes.rows_result, antes.columns_result, antes.size.cells
    )
    ventas = _casillas(paso, "ventas.csv")
    ventas.variables["Importe"].set(False)
    ventas.variables["Cantidad"].set(False)
    ventas._cambio()
    raiz.update()
    assert "Calculando" in str(paso.etiqueta_tamano.cget("text"))
    _listo(raiz, paso)
    despues = paso.calculo.diagnostico
    assert despues.columns_result == antes.columns_result - 2
    assert despues.size.cells < antes.size.cells
    assert f"{despues.columns_result} columnas" in str(paso.etiqueta_tamano.cget("text"))
    assert "Importe" not in paso.lista_final.textos()
    assert "●" in str(paso.etiqueta_excel.cget("text"))
    assert "Excel" in str(paso.etiqueta_excel.cget("text"))
    assert "Google Sheets" in str(paso.etiqueta_sheets.cget("text"))


def test_muchos_clics_seguidos_calculan_una_sola_vez(raiz, carpeta, monkeypatch):
    llamadas = []
    from excel_tool.ui.unir.pasos import paso5_columnas as mod

    original = mod.calcular

    def contar(*args, **kwargs):
        llamadas.append(1)
        return original(*args, **kwargs)

    monkeypatch.setattr(mod, "calcular", contar)
    paso = _abrir(raiz, _estado_todo(carpeta))
    paso.espera_ms = 200
    _listo(raiz, paso)
    llamadas.clear()
    ventas = _casillas(paso, "ventas.csv")
    for col in ("Fecha", "Artículo", "Cantidad"):
        ventas.variables[col].set(False)
        ventas._cambio()
    _esperar(raiz, lambda: paso.version_mostrada == paso.version)
    assert len(llamadas) == 1


def test_reordenar_se_guarda_en_el_estado_y_el_core_lo_respeta(raiz, carpeta):
    estado = _estado_todo(carpeta, modo_cobranzas="summary")
    paso = _abrir(raiz, estado)
    _listo(raiz, paso)
    lista = paso.lista_final
    ultimo = len(lista.textos()) - 1
    assert lista.textos()[ultimo] == "Total cobrado"
    lista.seleccionar(ultimo)
    lista.arriba_de_todo()
    assert lista.textos()[0] == "Total cobrado"
    lista.seleccionar(1)
    lista.bajar()
    assert estado.orden == lista.textos()
    lista.abajo_de_todo()
    assert estado.orden[-1] == lista.textos()[-1]
    lista.seleccionar(len(lista.textos()) - 1)
    lista.subir()
    paso.al_salir()
    assert estado.orden == lista.textos()
    resultado = run_join(estado.plan(), estado.tablas())
    assert list(resultado.df.columns) == estado.orden


def test_el_core_respeta_las_columnas_desmarcadas(raiz, carpeta):
    estado = _estado_todo(carpeta)
    paso = _abrir(raiz, estado)
    _listo(raiz, paso)
    ventas = _casillas(paso, "ventas.csv")
    ventas.quitar_todas()
    ventas.variables["Importe"].set(True)
    ventas._cambio()
    _casillas(paso, "clientes.xlsx").variables["Vendedor"].set(False)
    _casillas(paso, "clientes.xlsx")._cambio()
    paso.al_salir()
    assert estado.columnas["ventas.csv"] == ["Importe"]
    assert "Vendedor" not in estado.columnas["clientes.xlsx"]
    assert "ID cliente" in estado.columnas["clientes.xlsx"]
    df = run_join(estado.plan(), estado.tablas()).df
    assert "Importe" in df.columns
    assert "Vendedor" not in df.columns
    assert "Artículo" not in df.columns
    assert list(df.columns) == estado.orden


def test_el_orden_de_una_receta_se_conserva(raiz, carpeta):
    estado = _estado_todo(carpeta, modo_cobranzas="summary")
    estado.orden = ["Total cobrado", "Importe", "Nombre", "ID cliente"]
    paso = _abrir(raiz, estado)
    _listo(raiz, paso)
    assert paso.lista_final.textos()[:4] == ["Total cobrado", "Importe", "Nombre", "ID cliente"]
    assert estado.orden[:4] == ["Total cobrado", "Importe", "Nombre", "ID cliente"]


def test_buscador(raiz, carpeta):
    paso = _abrir(raiz, _estado_todo(carpeta))
    ventas = _casillas(paso, "ventas.csv")
    ventas.busqueda.set("arT")
    raiz.update()
    assert ventas.visibles() == ["Artículo"]
    assert ventas.casillas["Artículo"].winfo_manager() == "pack"
    assert ventas.casillas["Fecha"].winfo_manager() == ""
    ventas.busqueda.set("")
    assert len(ventas.visibles()) == 4


def test_todas_y_ninguna(raiz, carpeta):
    paso = _abrir(raiz, _estado_todo(carpeta))
    _listo(raiz, paso)
    clientes = _casillas(paso, "clientes.xlsx")
    ventas = _casillas(paso, "ventas.csv")
    ventas.boton_ninguna.invoke()
    assert ventas.marcadas() == []
    clientes.boton_ninguna.invoke()
    assert clientes.marcadas() == ["ID cliente"]  # la clave no se puede quitar
    ventas.busqueda.set("a")
    ventas.boton_todas.invoke()
    assert set(ventas.marcadas()) == {"Fecha", "Artículo", "Cantidad"}
    ventas.busqueda.set("")
    ventas.boton_todas.invoke()
    assert len(ventas.marcadas()) == 4
    assert str(ventas.cuenta.cget("text")) == "4 de 4"


def test_validacion_pide_al_menos_una_columna_ademas_de_la_clave(raiz, carpeta):
    estado = _estado_todo(carpeta)
    asistente = Asistente(raiz, estado, paso_inicial=4)
    paso = asistente.pasos[4]
    paso.espera_ms = 10
    assert paso.validar() is None
    for archivo in ("clientes.xlsx", "ventas.csv", "cobranzas.xlsx"):
        _casillas(paso, archivo).quitar_todas()
    assert paso.validar() == SIN_COLUMNAS
    asistente.boton_siguiente.invoke()
    assert asistente.indice == 4
    assert str(asistente.estado_texto.cget("text")) == SIN_COLUMNAS
    _casillas(paso, "ventas.csv").variables["Fecha"].set(True)
    _casillas(paso, "ventas.csv")._cambio()
    assert paso.validar() is None


def test_formato_argentino():
    assert texto_celdas(41_300_000) == "41,3 millones de celdas"
    assert texto_celdas(41_000_000) == "41 millones de celdas"
    assert texto_celdas(1_000_000) == "1 millón de celdas"
    assert texto_celdas(123_456) == "123.456 celdas"
    assert texto_tamano(2_300_000, 18, 41_400_000) == (
        "2.300.000 filas × 18 columnas = 41,4 millones de celdas"
    )


def test_sin_pasos_anteriores_no_falla(raiz):
    paso = _abrir(raiz, EstadoUnion())
    raiz.update()
    assert paso.validar() == SIN_COLUMNAS
    assert isinstance(paso, ttk.Frame)
