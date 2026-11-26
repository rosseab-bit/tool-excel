"""Tests con pantalla (xvfb) de los pasos 1 a 3 del asistente (F5-11), con el escenario real."""

from __future__ import annotations

import os
import shutil
import sys
import time
import tkinter as tk
from pathlib import Path

import pytest

from excel_tool.core.join import run_join
from excel_tool.ui.selector import Seleccion
from excel_tool.ui.unir import widgets_archivos
from excel_tool.ui.unir.asistente import Asistente
from excel_tool.ui.unir.estado import Fuente

from .test_unir_asistente import _textos
from .test_unir_estado import _estado

pytestmark = pytest.mark.pantalla


@pytest.fixture
def raiz(tmp_path, monkeypatch):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos_app"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos_app"))
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


@pytest.fixture
def carpeta(fixtures_generados, tmp_path) -> Path:
    destino = tmp_path / "datos"
    shutil.copytree(fixtures_generados / "escenario_3", destino)
    return destino


@pytest.fixture
def explorados(monkeypatch):
    """Reemplaza el explorador: anota qué selecciones se abrieron."""
    abiertos: list[Seleccion] = []
    monkeypatch.setattr(
        widgets_archivos, "abrir_explorador", lambda parent, sel, *a, **k: abiertos.append(sel)
    )
    return abiertos


@pytest.fixture
def elegir(monkeypatch, carpeta):
    """Reemplaza el selector: `elegir("ventas.csv")` hace que el próximo selector devuelva ese
    archivo, sin abrir ninguna ventana."""
    pendiente: list[str] = []

    def falso(parent, on_accept, path=None):
        nombre = pendiente.pop(0)
        on_accept(Seleccion(carpeta / nombre, None, None, 0))

    monkeypatch.setattr(widgets_archivos, "abrir_selector", falso)
    return pendiente.append


def _esperar(raiz, condicion, segundos=20):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        raiz.update()
        if condicion():
            return
        time.sleep(0.02)
    raise AssertionError("se agotó la espera")


def _asentado(raiz, paso):
    _esperar(raiz, lambda: not paso.calculando)


def _cargar_principal(raiz, asistente, elegir):
    elegir("clientes.xlsx")
    paso = asistente.pasos[0]
    paso.boton_elegir.invoke()
    _esperar(raiz, lambda: asistente.estado.principal is not None)
    return paso


def _marcar_clave(raiz, asistente, nombre="ID cliente"):
    paso = asistente.pasos[1]
    asistente.siguiente()
    assert asistente.indice == 1
    paso.casillas[nombre].invoke()
    return paso


def _agregar(raiz, paso3, elegir, nombre):
    antes = len(paso3.estado.secundarios)
    elegir(nombre)
    paso3.boton_agregar.invoke()
    _esperar(raiz, lambda: len(paso3.estado.secundarios) == antes + 1)
    return paso3.estado.secundarios[-1]


def _al_paso3(raiz, asistente, elegir):
    _cargar_principal(raiz, asistente, elegir)
    _marcar_clave(raiz, asistente)
    asistente.siguiente()
    assert asistente.indice == 2
    return asistente.pasos[2]


# ---------------------------------------------------------------- paso 1


def test_paso1_cargar_principal(raiz, elegir, explorados):
    asistente = Asistente(raiz)
    raiz.update()
    paso = asistente.pasos[0]
    assert paso.validar() is not None
    assert "Todavía no elegiste" in str(paso.etiqueta_nombre.cget("text"))
    _cargar_principal(raiz, asistente, elegir)
    assert paso.validar() is None
    assert str(paso.etiqueta_nombre.cget("text")) == "clientes.xlsx"
    assert str(paso.etiqueta_detalle.cget("text")) == "30 filas · 4 columnas"
    textos = _textos(paso)
    assert any("una fila por cada fila de este archivo" in t for t in textos)
    paso.boton_info.invoke()
    assert len(explorados) == 1
    assert explorados[0].path.name == "clientes.xlsx"
    # el archivo y su perfil quedaron cargados para los demás pasos
    assert asistente.estado.principal.perfil is not None


def test_paso1_cambiar_principal_reinicia_la_clave(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    _cargar_principal(raiz, asistente, elegir)
    asistente.estado.claves = ["ID cliente"]
    elegir("cobranzas.xlsx")
    asistente.pasos[0].boton_elegir.invoke()
    _esperar(raiz, lambda: asistente.estado.principal.nombre == "cobranzas.xlsx")
    assert asistente.estado.claves == []


# ---------------------------------------------------------------- paso 2


def test_paso2_claves_con_llave_primero(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    _cargar_principal(raiz, asistente, elegir)
    paso = _marcar_clave(raiz, asistente)
    primera = next(iter(paso.casillas.values()))
    assert str(primera.cget("text")) == "🔑 ID cliente"
    assert list(paso.casillas) == ["ID cliente", "Nombre", "Región", "Vendedor"]
    assert any("identifica a cada cliente" in t for t in _textos(paso))
    assert not str(paso.casillas["Región"].cget("text")).startswith("🔑")
    # las que el perfil marca como posible clave van primero, el resto en su orden
    for col in asistente.estado.principal.perfil.columns_profile:
        col.is_unique = col.name in ("Región", "Vendedor")
    paso._archivo = None
    paso.al_entrar()
    assert list(paso.casillas) == ["Región", "Vendedor", "ID cliente", "Nombre"]
    assert str(paso.casillas["Región"].cget("text")) == "🔑 Región"
    assert str(paso.casillas["Nombre"].cget("text")) == "Nombre"


def test_paso2_validar_y_volcar(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    _cargar_principal(raiz, asistente, elegir)
    asistente.siguiente()
    paso = asistente.pasos[1]
    assert paso.validar() is not None
    asistente.siguiente()
    assert asistente.indice == 1  # no avanza sin clave
    assert "al menos una columna clave" in str(asistente.estado_texto.cget("text"))
    paso.casillas["ID cliente"].invoke()
    paso.casillas["Región"].invoke()
    assert paso.validar() is None
    assert asistente.estado.claves == ["ID cliente", "Región"]
    paso.casillas["Región"].invoke()
    assert asistente.estado.claves == ["ID cliente"]
    asistente.siguiente()
    assert asistente.indice == 2


def test_paso2_sin_principal_no_avanza(raiz):
    asistente = Asistente(raiz)
    raiz.update()
    asistente.siguiente()
    assert asistente.indice == 0
    assert "principal" in str(asistente.estado_texto.cget("text"))


# ---------------------------------------------------------------- paso 3


def test_paso3_ventas_se_empareja_sola(raiz, elegir, explorados):
    asistente = Asistente(raiz)
    raiz.update()
    paso = _al_paso3(raiz, asistente, elegir)
    assert paso.validar() is not None  # todavía no hay archivos
    ventas = _agregar(raiz, paso, elegir, "ventas.csv")
    assert ventas.key_pairs == [("ID cliente", "ID cliente")]
    assert ventas.mode is None  # nunca se elige por la persona
    fila = paso.filas["ventas.csv"]
    assert fila.combos_clave == {}
    _asentado(raiz, paso)
    texto = str(fila.indicador.cget("text"))
    assert texto.startswith("✅")
    assert "de 30" in texto
    assert paso.validar() is None
    assert ventas.mode is None  # el modo provisorio no se guarda
    fila.boton_info.invoke()
    assert explorados[-1].path.name == "ventas.csv"


def test_paso3_cobranzas_a_mano_y_ceros(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    paso = _al_paso3(raiz, asistente, elegir)
    cobranzas = _agregar(raiz, paso, elegir, "cobranzas.xlsx")
    assert cobranzas.key_pairs == []
    fila = paso.filas["cobranzas.xlsx"]
    assert list(fila.combos_clave) == ["ID cliente"]
    assert "columna en común" in str(fila.indicador.cget("text"))
    assert "Elegí la columna" in paso.validar()
    combo = fila.combos_clave["ID cliente"]
    assert "Nro Cliente" in combo.cget("values")
    combo.set("Nro Cliente")
    combo.event_generate("<<ComboboxSelected>>")
    assert cobranzas.key_pairs == [("ID cliente", "Nro Cliente")]
    assert paso.validar() is None
    _asentado(raiz, paso)
    sin_ceros = str(fila.indicador.cget("text"))
    assert sin_ceros.startswith("⚠️")
    assert "5 de 30" in sin_ceros
    assert "Sacar ceros a la izquierda" in sin_ceros
    assert "27" in sin_ceros

    paso._variables["strip_leading_zeros"].set(True)
    paso._cambio_limpieza()
    assert asistente.estado.key_options.strip_leading_zeros is True
    _asentado(raiz, paso)
    con_ceros = str(fila.indicador.cget("text"))
    assert con_ceros.startswith("✅")
    assert "27 de 30" in con_ceros


def test_paso3_opciones_de_limpieza_por_defecto(raiz):
    asistente = Asistente(raiz)
    raiz.update()
    paso = asistente.pasos[2]
    assert {k: v.get() for k, v in paso._variables.items()} == {
        "ignore_case": True,
        "trim": True,
        "strip_leading_zeros": False,
    }
    textos = _textos(paso)
    assert "Sacar ceros a la izquierda (00012 = 12)" in textos


def test_paso3_reordenar_y_quitar(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    paso = _al_paso3(raiz, asistente, elegir)
    _agregar(raiz, paso, elegir, "ventas.csv")
    _agregar(raiz, paso, elegir, "cobranzas.xlsx")
    nombres = lambda: [s.nombre for s in asistente.estado.secundarios]  # noqa: E731
    assert nombres() == ["ventas.csv", "cobranzas.xlsx"]
    assert paso.filas["ventas.csv"].boton_subir.instate(["disabled"])
    assert paso.filas["cobranzas.xlsx"].boton_bajar.instate(["disabled"])
    paso.filas["cobranzas.xlsx"].boton_subir.invoke()
    assert nombres() == ["cobranzas.xlsx", "ventas.csv"]
    paso.filas["cobranzas.xlsx"].boton_bajar.invoke()
    assert nombres() == ["ventas.csv", "cobranzas.xlsx"]
    paso.filas["ventas.csv"].boton_quitar.invoke()
    assert nombres() == ["cobranzas.xlsx"]
    assert list(paso.filas) == ["cobranzas.xlsx"]
    paso.filas["cobranzas.xlsx"].boton_quitar.invoke()
    assert paso.validar() is not None
    assert "al menos un archivo" in paso.validar()


def test_paso3_nombres_repetidos(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    paso = _al_paso3(raiz, asistente, elegir)
    _agregar(raiz, paso, elegir, "ventas.csv")
    _agregar(raiz, paso, elegir, "ventas.csv")
    _agregar(raiz, paso, elegir, "ventas.csv")
    nombres = [s.nombre for s in asistente.estado.secundarios]
    assert nombres == ["ventas.csv", "ventas.csv (2)", "ventas.csv (3)"]
    assert len(set(nombres)) == 3
    assert asistente.estado.secundarios[1].spec.name == "ventas.csv (2)"
    _asentado(raiz, paso)
    # el core acepta el plan con nombres únicos
    for sec in asistente.estado.secundarios:
        sec.mode = "first"
    run_join(asistente.estado.plan(), asistente.estado.tablas())


def test_paso3_el_principal_tambien_cuenta_para_los_nombres(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    paso = _al_paso3(raiz, asistente, elegir)
    sec = _agregar(raiz, paso, elegir, "clientes.xlsx")
    assert sec.nombre == "clientes.xlsx (2)"


def test_paso3_faltantes_de_una_receta(raiz, elegir, carpeta):
    estado = _estado(carpeta)
    cobranzas = estado.secundarios[1]
    cobranzas.df = None
    cobranzas.perfil = None
    asistente = Asistente(raiz, estado, paso_inicial=2)
    raiz.update()
    paso = asistente.pasos[2]
    assert estado.faltantes == ["cobranzas.xlsx"]
    fila = paso.filas["cobranzas.xlsx"]
    textos = _textos(fila.marco)
    assert any("Falta el archivo" in t for t in textos)
    assert str(fila.boton_elegir.cget("text")) == "Elegir…"
    assert "cobranzas.xlsx" in paso.validar()
    asistente.siguiente()
    assert asistente.indice == 2
    elegir("cobranzas.xlsx")
    fila.boton_elegir.invoke()
    _esperar(raiz, lambda: not estado.faltantes)
    # se conserva lo que traía la receta
    nueva = estado.secundarios[1]
    assert nueva.mode == "summary"
    assert nueva.key_pairs == [("ID cliente", "Nro Cliente")]
    assert paso.validar() is None
    assert not paso.filas["cobranzas.xlsx"].boton_elegir
    _asentado(raiz, paso)
    assert str(paso.filas["cobranzas.xlsx"].indicador.cget("text")).startswith("⚠️")


def test_paso3_principal_faltante_de_una_receta(raiz, elegir, carpeta):
    estado = _estado(carpeta)
    estado.principal = Fuente(estado.principal.spec)
    asistente = Asistente(raiz, estado, paso_inicial=2)
    raiz.update()
    paso = asistente.pasos[2]
    assert estado.faltantes == ["clientes.xlsx"]
    assert paso.fila_principal is not None
    assert "clientes.xlsx" in paso.validar()
    elegir("clientes.xlsx")
    paso.fila_principal.boton_elegir.invoke()
    _esperar(raiz, lambda: not estado.faltantes)
    assert estado.claves == ["ID cliente"]
    assert paso.fila_principal is None


def test_paso3_union_completar_o_solo_los_dos(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    paso = _al_paso3(raiz, asistente, elegir)
    ventas = _agregar(raiz, paso, elegir, "ventas.csv")
    assert ventas.how == "left"
    combo = paso.filas["ventas.csv"].combo_union
    assert str(combo.get()) == "Completar (todos los del principal)"
    combo.set("Solo los que están en los dos")
    combo.event_generate("<<ComboboxSelected>>")
    assert ventas.how == "inner"


def test_paso3_validaciones(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    paso = _al_paso3(raiz, asistente, elegir)
    asistente.siguiente()
    assert asistente.indice == 2
    assert "al menos un archivo" in str(asistente.estado_texto.cget("text"))
    _agregar(raiz, paso, elegir, "cobranzas.xlsx")
    asistente.siguiente()
    assert asistente.indice == 2
    assert "columna en común" in str(asistente.estado_texto.cget("text"))
    combo = paso.filas["cobranzas.xlsx"].combos_clave["ID cliente"]
    combo.set("Nro Cliente")
    combo.event_generate("<<ComboboxSelected>>")
    _asentado(raiz, paso)
    asistente.siguiente()
    assert asistente.indice == 3


def test_paso3_cambiar_la_clave_vuelve_a_emparejar(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    paso = _al_paso3(raiz, asistente, elegir)
    ventas = _agregar(raiz, paso, elegir, "ventas.csv")
    assert ventas.key_pairs == [("ID cliente", "ID cliente")]
    asistente.anterior()
    asistente.pasos[1].casillas["ID cliente"].invoke()
    asistente.pasos[1].casillas["Región"].invoke()
    assert asistente.estado.claves == ["Región"]
    asistente.siguiente()
    assert ventas.key_pairs == []
    fila = asistente.pasos[2].filas["ventas.csv"]
    assert list(fila.combos_clave) == ["Región"]
    assert "Elegí la columna" in asistente.pasos[2].validar()


def test_nombres_y_textos_sin_jerga(raiz, elegir):
    asistente = Asistente(raiz)
    raiz.update()
    paso = _al_paso3(raiz, asistente, elegir)
    _agregar(raiz, paso, elegir, "ventas.csv")
    _asentado(raiz, paso)
    todo = " ".join(_textos(paso)).lower()
    for jerga in ("join", "merge", "left", "inner", "dataframe", "key"):
        assert jerga not in todo
