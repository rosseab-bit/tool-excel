"""Tests con pantalla (xvfb) del paso 4 del asistente: ¿Detalle o resumen? (F5-12)."""

from __future__ import annotations

import os
import shutil
import sys
import time
import tkinter as tk
from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy
from excel_tool.core.io import ReadOptions, read_table
from excel_tool.core.join import SourceSpec
from excel_tool.ui.unir.estado import EstadoUnion, Fuente, FuenteSecundaria
from excel_tool.ui.unir.pasos import paso4_modo
from excel_tool.ui.unir.pasos.paso4_modo import PasoModo

pytestmark = pytest.mark.pantalla


@pytest.fixture
def raiz():
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


@pytest.fixture
def carpeta(fixtures_generados, tmp_path) -> Path:
    destino = tmp_path / "datos"
    shutil.copytree(fixtures_generados / "escenario_3", destino)
    return destino


def _fuente(carpeta, nombre, columnas):
    spec = SourceSpec(nombre, carpeta / nombre, ReadOptions(), columnas)
    return Fuente(spec, read_table(spec.path).df)


def _sec(carpeta, nombre, pares, **kw):
    base = _fuente(carpeta, nombre, [])
    return FuenteSecundaria(base.spec, base.df, key_pairs=pares, **kw)


@pytest.fixture
def estado(carpeta):
    return EstadoUnion(
        principal=_fuente(carpeta, "clientes.xlsx", ["ID cliente", "Nombre", "Región"]),
        claves=["ID cliente"],
        secundarios=[
            _sec(carpeta, "ventas.csv", [("ID cliente", "ID cliente")]),
            _sec(carpeta, "cobranzas.xlsx", [("ID cliente", "Nro Cliente")]),
        ],
    )


def _textos(widget):
    encontrados = []
    for hijo in widget.winfo_children():
        try:
            encontrados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        encontrados.extend(_textos(hijo))
    return encontrados


def _esperar(raiz, condicion, segundos=20):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        raiz.update()
        if condicion():
            return
        time.sleep(0.02)
    raise AssertionError("se agotó la espera")


def _paso(raiz, estado) -> PasoModo:
    paso = PasoModo(raiz, estado, None)
    paso.pack(fill="both", expand=True)
    paso.al_entrar()
    _esperar(raiz, lambda: not paso.pendiente)
    return paso


def _elegir(raiz, paso, tarjeta, boton):
    getattr(tarjeta, boton).invoke()
    _esperar(raiz, lambda: not paso.pendiente)


def _cerrar_editor(raiz, paso):
    _esperar(raiz, lambda: paso.editor is not None)
    paso.boton_listo.invoke()
    _esperar(raiz, lambda: not paso.pendiente)


def test_sin_modo_elegido_no_se_puede_seguir(raiz, estado):
    paso = _paso(raiz, estado)
    assert [s.mode for s in estado.secundarios] == [None, None]
    assert [t.var.get() for t in paso.tarjetas] == ["", ""]
    assert "ventas.csv" in paso.validar()
    _elegir(raiz, paso, paso.tarjetas[0], "radio_detalle")
    assert "cobranzas.xlsx" in paso.validar()


def test_ejemplo_de_cuantas_veces_aparece(raiz, estado):
    paso = _paso(raiz, estado)
    veces = paso.diagnostico.secondaries[0].example_many
    assert veces is not None
    esperado = f"El cliente {veces[0]} aparece {veces[1]} veces en ventas.csv"
    assert str(paso.tarjetas[0].etiqueta_ejemplo.cget("text")) == esperado


def test_clave_que_no_se_repite_queda_en_primera_fila_sin_preguntar(raiz, estado, carpeta):
    shutil.copy(carpeta / "clientes.xlsx", carpeta / "clientes_copia.xlsx")
    estado.secundarios[0] = _sec(
        carpeta, "clientes_copia.xlsx", [("ID cliente", "ID cliente")]
    )  # una fila por cliente
    paso = _paso(raiz, estado)
    sec = estado.secundarios[0]
    assert sec.mode == "first"
    assert sec.nombre in paso.automaticos
    assert "Cada cliente aparece una sola vez" in str(
        paso.tarjetas[0].etiqueta_ejemplo.cget("text")
    )
    assert not paso.tarjetas[0].opciones.winfo_manager()
    # al volver al paso se vuelve a decidir
    paso.al_entrar()
    _esperar(raiz, lambda: not paso.pendiente)
    assert sec.mode == "first"


def test_ventas_en_detalle_muestra_88_filas(raiz, estado):
    paso = _paso(raiz, estado)
    _elegir(raiz, paso, paso.tarjetas[0], "radio_detalle")
    assert estado.secundarios[0].mode == "detail"
    assert "88 filas" in str(paso.tarjetas[0].etiqueta_filas.cget("text"))
    assert paso.diagnostico.rows_result == 88
    assert not paso.marco_varios.winfo_manager()


def test_dos_detalles_avisan_y_cambiar_a_resumen_en_un_clic(raiz, estado):
    paso = _paso(raiz, estado)
    _elegir(raiz, paso, paso.tarjetas[0], "radio_detalle")
    _elegir(raiz, paso, paso.tarjetas[1], "radio_detalle")
    assert paso.diagnostico.many_details
    assert paso.marco_varios.winfo_manager()
    assert any("se combinan todas con todas" in t for t in _textos(paso.marco_varios))
    assert [str(b.cget("text")) for b in paso.botones_cambiar] == [
        "Cambiar cobranzas.xlsx a resumen"
    ]
    paso.botones_cambiar[0].invoke()
    cobranzas = estado.secundarios[1]
    assert cobranzas.mode == "summary"
    assert cobranzas.summary is not None
    funciones = [a.func for a in cobranzas.summary.aggregations]
    assert funciones[0] == "count" and "sum" in funciones
    assert [g.column for g in cobranzas.summary.group_by] == ["Nro Cliente"]
    assert paso.editor is not None  # se abrió para ajustarlo
    _cerrar_editor(raiz, paso)
    assert not paso.diagnostico.many_details
    assert not paso.marco_varios.winfo_manager()
    assert paso.validar() is None


def test_resumen_con_el_editor_guarda_el_resumen_en_el_estado(raiz, estado):
    paso = _paso(raiz, estado)
    tarjeta = paso.tarjetas[0]
    tarjeta.radio_resumen.invoke()
    _esperar(raiz, lambda: paso.editor is not None)
    assert paso.editor.fixed_group == ["ID cliente"]
    assert estado.secundarios[0].mode is None  # todavía no confirmó
    paso.editor.set_spec(
        AggregationSpec([GroupBy("ID cliente")], [Aggregation("Importe", "sum", "Total comprado")])
    )
    paso.boton_listo.invoke()
    _esperar(raiz, lambda: not paso.pendiente)
    sec = estado.secundarios[0]
    assert sec.mode == "summary"
    assert [g.column for g in sec.summary.group_by] == ["ID cliente"]
    assert sec.summary.aggregations[0].name == "Total comprado"
    assert paso.dialogo is None
    assert "Resumen armado" in str(tarjeta.etiqueta_resumen.cget("text"))


def test_cancelar_el_editor_deja_todo_como_estaba(raiz, estado):
    paso = _paso(raiz, estado)
    paso.tarjetas[0].radio_resumen.invoke()
    _esperar(raiz, lambda: paso.editor is not None)
    for hijo in _botones(paso.dialogo):
        if str(hijo.cget("text")) == "Cancelar":
            hijo.invoke()
            break
    assert estado.secundarios[0].mode is None
    assert paso.tarjetas[0].var.get() == ""
    assert paso.dialogo is None


def _botones(widget):
    encontrados = []
    for hijo in widget.winfo_children():
        if hijo.winfo_class() == "TButton":
            encontrados.append(hijo)
        encontrados.extend(_botones(hijo))
    return encontrados


def test_primera_fila_esta_en_mas_opciones(raiz, estado):
    paso = _paso(raiz, estado)
    tarjeta = paso.tarjetas[0]
    assert not tarjeta.cuerpo_mas.winfo_manager()
    tarjeta.boton_mas.invoke()
    assert tarjeta.cuerpo_mas.winfo_manager()
    _elegir(raiz, paso, tarjeta, "radio_primera")
    assert estado.secundarios[0].mode == "first"
    assert paso.validar() is not None  # falta cobranzas


def test_resumen_vacio_no_deja_seguir(raiz, estado):
    paso = _paso(raiz, estado)
    estado.secundarios[0].mode = "summary"
    estado.secundarios[0].summary = AggregationSpec([], [])
    estado.secundarios[1].mode = "first"
    assert "Falta armar el resumen de «ventas.csv»" in paso.validar()


def test_aviso_adelantado_de_tamano(raiz, estado, monkeypatch):
    pedidos = []

    def falso_check(est, destino, formato, **kw):
        pedidos.append((destino, formato))
        return SimpleNamespace(level="obligatorio")

    monkeypatch.setattr(paso4_modo, "limits", SimpleNamespace(check=falso_check))
    paso = _paso(raiz, estado)
    _elegir(raiz, paso, paso.tarjetas[0], "radio_detalle")
    texto = str(paso.etiqueta_tamano.cget("text"))
    assert texto.startswith("🔴 El resultado va a tener ~88 filas")
    assert "no entra en una sola hoja de Excel" in texto
    assert "Al guardar te vamos a proponer dividirlo" in texto
    assert pedidos and pedidos[-1] == ("excel", "xlsx")


def test_sin_aviso_de_tamano_si_entra(raiz, estado):
    paso = _paso(raiz, estado)
    _elegir(raiz, paso, paso.tarjetas[0], "radio_detalle")
    assert str(paso.etiqueta_tamano.cget("text")) == ""


def test_un_resultado_viejo_no_pisa_al_nuevo(raiz, estado):
    paso = _paso(raiz, estado)
    paso.tarjetas[0].radio_detalle.invoke()  # lanza un cálculo...
    paso.tarjetas[1].radio_primera.invoke()  # ...y otro enseguida
    assert paso.version >= 3
    _esperar(raiz, lambda: not paso.pendiente)
    assert paso.diagnostico is not None
    assert "88 filas" in str(paso.tarjetas[0].etiqueta_filas.cget("text"))
