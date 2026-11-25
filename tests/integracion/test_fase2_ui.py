"""Integración de la Fase 2 (UI con xvfb y el core REAL, sin perfiles falsos)."""

from __future__ import annotations

import os
import subprocess
import sys
import textwrap
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk

import pytest

from excel_tool import platform_utils
from excel_tool.core.io import CsvOptions, list_sheets, read_table
from excel_tool.ui import selector
from excel_tool.ui.datos import PanelDatos
from excel_tool.ui.explorador import abrir_explorador
from excel_tool.ui.selector import Seleccion

pytestmark = pytest.mark.pantalla

CSV_AR = CsvOptions(";", "utf-8-sig")


def _csv(gen: Path) -> Seleccion:
    return Seleccion(gen / "explorar" / "para_explorar.csv", None, CSV_AR, 0)


def _xlsx(gen: Path) -> Seleccion:
    return Seleccion(gen / "explorar" / "para_explorar.xlsx", None, None, 0)


def _textos(widget) -> list[str]:
    hallados = []
    for hijo in widget.winfo_children():
        try:
            hallados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        hallados.extend(_textos(hijo))
    return hallados


def _abrir(entorno, seleccion):
    ex = abrir_explorador(entorno.raiz, seleccion)
    entorno.esperar(lambda: ex.perfil is not None)
    entorno.esperar(lambda: not ex.analizando)
    return ex


def _datos_quietos(entorno, panel: PanelDatos):
    """Espera a que termine el filtro/orden en segundo plano (Aplicar vuelve a estar libre)."""
    entorno.esperar(lambda: not panel.boton_aplicar.instate(["disabled"]))


@pytest.fixture(params=["csv", "xlsx"])
def abierto(request, entorno_ui, fixtures_generados):
    seleccion = (_csv if request.param == "csv" else _xlsx)(fixtures_generados)
    ex = _abrir(entorno_ui, seleccion)
    return entorno_ui, ex


# ---------------------------------------------------------------- recorrer las 4 pestañas
def test_recorrer_las_cuatro_pestanas(abierto):
    entorno, ex = abierto
    nombres = [ex.pestanas.tab(i, "text") for i in range(4)]
    assert nombres == ["General", "Columnas", "Valores distintos", "Datos"]
    for pestana in (
        ex.pestana_general,
        ex.pestana_columnas,
        ex.pestana_distintos,
        ex.pestana_datos,
    ):
        ex.pestanas.select(pestana)
        entorno.bombear(0.05)
    # General
    textos = _textos(ex.pestana_general)
    assert "200" in textos and "13" in textos  # filas y columnas
    assert any("Región" in t and "espacios" in t for t in textos)
    # Columnas
    filas = [ex.tabla.item(i, "values") for i in ex.tabla.get_children()]
    assert len(filas) == 13
    assert [f[1] for f in filas][:3] == ["ID", "Importe", "Fecha"]
    assert filas[1][2] == "Número decimal"
    assert filas[12][3] == "196"
    ex.seleccionar_columna("Región")
    assert ex.columna_elegida().name == "Región"
    assert ex.tabla_frecuentes is not None and len(ex.tabla_frecuentes.get_children()) == 6
    assert any("Norte" in t for t in _textos(ex.contenido_detalle))
    # Datos: arranca con todas las filas, sin filtros
    assert ex.panel_datos.grilla.total_filas == 200
    assert ex.panel_datos.grilla.filas_mostradas() == 200
    assert ex.panel_datos.etiqueta_conteo.cget("text") == "Mostrando 200 de 200 filas"
    # nada de jerga ni de tracebacks en ninguna ventana abierta
    todo = " ".join(_textos(ex))
    for mala in ("DataFrame", "dtype", "encoding", "Traceback", "null", "distinct ", "group by"):
        assert mala not in todo
    assert entorno.carteles == []


# ---------------------------------------------------------------- distintos, doble clic y filtro
def test_distintos_doble_clic_filtra_datos_y_se_quita(abierto):
    entorno, ex = abierto
    pd_ = ex.panel_distintos
    assert pd_.resultado is None  # los distintos se calculan solo a pedido
    ex.seleccionar_columna("Región")
    ex._ver_distintos()  # el botón "Ver valores distintos de esta columna"
    entorno.esperar(lambda: pd_.resultado is not None)
    assert ex.pestanas.select() == str(ex.pestana_distintos)
    assert len(pd_.tabla.get_children()) == 7
    assert "7 valores distintos en 200 filas" in pd_.etiqueta_resumen.cget("text")
    assert "parecen el mismo escrito distinto" in pd_.etiqueta_variantes.cget("text")
    iid = next(i for i in pd_.tabla.get_children() if pd_.tabla.item(i, "values")[0] == "Norte")
    pd_.tabla.selection_set(iid)
    entorno.raiz.update()
    ex.update()
    x, y, ancho, alto = pd_.tabla.bbox(iid)
    for _ in range(2):  # doble clic real: dos clics seguidos sobre la fila
        for evento in ("<ButtonPress-1>", "<ButtonRelease-1>"):
            pd_.tabla.event_generate(evento, x=x + 5, y=y + alto // 2)
    entorno.esperar(lambda: ex.panel_datos.aviso_valor is not None)
    assert ex.pestanas.select() == str(ex.pestana_datos)
    assert len(ex.panel_datos.filas) == 31
    assert ex.panel_datos.grilla.filas_mostradas() == 31
    assert "Región es «Norte»" in ex.panel_datos.aviso_valor
    assert ex.panel_datos.etiqueta_conteo.cget("text") == "Mostrando 31 de 200 filas"
    celdas = [
        ex.panel_datos.grilla.arbol.item(i, "values")
        for i in ex.panel_datos.grilla.arbol.get_children()
    ]
    posicion = ex.panel_datos.grilla.columnas.index("Región")
    assert {c[posicion] for c in celdas} == {"Norte"}
    # quitar el filtro
    ex.panel_datos.boton_quitar_filtro.invoke()
    entorno.esperar(lambda: ex.panel_datos.aviso_valor is None)
    assert len(ex.panel_datos.filas) == 200
    assert ex.panel_datos.grilla.filas_mostradas() == 200
    assert entorno.carteles == []


def test_juntar_variantes_en_distintos(abierto):
    entorno, ex = abierto
    pd_ = ex.panel_distintos
    ex.seleccionar_columna("Región")
    pd_.elegir_columnas(["Región"])
    pd_.var_juntar.set(True)
    pd_.calcular()
    entorno.esperar(lambda: pd_.resultado is not None)
    assert len(pd_.tabla.get_children()) == 5  # Centro, Sur, Norte (junto), Cuyo, (vacío)
    assert pd_.etiqueta_variantes.cget("text") == ""


# ---------------------------------------------------------------- filtros y orden en Datos
def test_filtros_contiene_y_mayor_y_orden(abierto):
    entorno, ex = abierto
    datos = ex.panel_datos
    ex.pestanas.select(ex.pestana_datos)
    datos.agregar_condicion("Descripción", "contiene", "ñandú")
    datos.aplicar()
    _datos_quietos(entorno, datos)
    entorno.esperar(lambda: len(datos.filas) == 32)
    assert datos.grilla.filas_mostradas() == 32
    assert datos.etiqueta_conteo.cget("text") == "Mostrando 32 de 200 filas"

    datos.agregar_condicion("Importe", "es mayor que", "5000")
    datos.aplicar()
    entorno.esperar(lambda: len(datos.filas) < 32)
    esperado = len(
        [1 for i in datos.filas if float(str(ex.df["Importe"].iloc[i]).replace(",", ".")) > 5000]
    )
    assert len(datos.filas) == esperado > 0

    for fila in list(datos.condiciones):
        datos.quitar_condicion(fila)
    datos.agregar_condicion("Importe", "es mayor que", "5000")
    datos.aplicar()
    entorno.esperar(lambda: len(datos.filas) == 109)
    assert datos.grilla.filas_mostradas() == 109

    # orden: clic en el encabezado "Importe" dos veces (de mayor a menor)
    datos.grilla._clic_encabezado("Importe")
    entorno.esperar(lambda: datos._orden == ("Importe", True))
    datos.grilla._clic_encabezado("Importe")
    entorno.esperar(lambda: datos._orden == ("Importe", False))
    _datos_quietos(entorno, datos)
    arbol = datos.grilla.arbol
    columnas = datos.grilla.columnas
    primeras = [arbol.item(i, "values") for i in arbol.get_children()[:3]]
    assert [p[columnas.index("N.º fila")] for p in primeras] == ["122", "197", "160"]
    importes = [
        float(arbol.item(i, "values")[columnas.index("Importe")].replace(",", "."))
        for i in arbol.get_children()
    ]
    assert importes == sorted(importes, reverse=True)
    assert entorno.carteles == []


def test_filtro_con_valor_invalido_muestra_mensaje_amigable(abierto):
    entorno, ex = abierto
    datos = ex.panel_datos
    datos.agregar_condicion("Importe", "es mayor que", "abc")
    datos.aplicar()
    entorno.esperar(lambda: len(entorno.carteles) == 1)
    tipo, _, texto = entorno.carteles[0]
    assert tipo == "showerror"
    assert "Traceback" not in texto and "Error" not in texto.split()[:1]
    assert len(datos.filas) == 200  # no cambió nada


def test_la_grilla_nunca_pasa_de_una_pagina(abierto):
    entorno, ex = abierto
    chica = PanelDatos(ex.pestana_datos, ex, page_size=50)
    chica.cargar(ex.df, ex.perfil)
    chica.pack()
    entorno.raiz.update()
    assert chica.grilla.page_size == 50
    assert chica.grilla.filas_mostradas() == 50
    assert chica.grilla.cantidad_paginas == 4
    chica.grilla.siguiente()
    assert chica.grilla.filas_mostradas() == 50 and chica.grilla.pagina == 1
    chica.grilla.ultima()
    assert chica.grilla.filas_mostradas() == 50
    assert chica.grilla.arbol.item(chica.grilla.arbol.get_children()[0], "values")[0] == "151"
    # y la de verdad usa el tamaño de página del contrato
    assert ex.panel_datos.grilla.page_size == 1000


# ---------------------------------------------------------------- exportar
def test_exportar_informe_y_lista(abierto, tmp_path):
    entorno, ex = abierto
    ex.seleccionar_columna("Región")
    ex._ver_distintos()
    entorno.esperar(lambda: ex.panel_distintos.resultado is not None)

    entorno.destino = tmp_path / "informe.xlsx"
    ex.guardar_informe()
    entorno.esperar(lambda: any(c[0] == "showinfo" for c in entorno.carteles))
    assert entorno.carteles[-1][1] == "Listo" and "informe.xlsx" in entorno.carteles[-1][2]
    hojas = [h.name for h in list_sheets(entorno.destino)]
    assert hojas[:2] == ["General", "Columnas"] and "Región" in hojas
    assert read_table(entorno.destino).df.shape == (12, 2)

    entorno.destino = tmp_path / "valores.csv"
    ex.panel_distintos.exportar()
    entorno.esperar(lambda: sum(c[0] == "showinfo" for c in entorno.carteles) == 2)
    leido = read_table(entorno.destino).df
    assert list(leido.columns) == ["Región", "Cantidad", "Porcentaje"] and len(leido) == 7
    assert not [c for c in entorno.carteles if c[0] != "showinfo"]


def test_exportar_en_carpeta_inexistente_da_mensaje_amigable(abierto, tmp_path):
    entorno, ex = abierto
    entorno.destino = tmp_path / "no_existe" / "informe.xlsx"
    ex.guardar_informe()
    entorno.esperar(lambda: len(entorno.carteles) == 1)
    tipo, _, texto = entorno.carteles[0]
    assert tipo == "showerror" and "Traceback" not in texto and "Errno" not in texto


# ---------------------------------------------------------------- varias hojas
def test_xlsx_con_varias_hojas_cambiando_de_hoja(entorno_ui, fixtures_generados):
    ex = _abrir(
        entorno_ui, Seleccion(fixtures_generados / "formatos" / "varias_hojas.xlsx", None, None, 0)
    )
    assert ex.perfil.sheet == "Clientes" and (ex.perfil.rows, ex.perfil.columns) == (12, 4)
    assert len(ex.combo_hoja.cget("values")) == 3
    ex.cambiar_hoja("Ventas")
    entorno_ui.esperar(lambda: ex.perfil.sheet == "Ventas" and not ex.analizando)
    assert (ex.perfil.rows, ex.perfil.columns) == (25, 6)
    assert len(ex.tabla.get_children()) == 6
    assert ex.panel_datos.grilla.total_filas == 25
    assert ex.panel_distintos.resultado is None  # se limpian los distintos al cambiar de hoja
    ex.cambiar_hoja("Resumen año")
    entorno_ui.esperar(lambda: ex.perfil.sheet == "Resumen año" and not ex.analizando)
    assert (ex.perfil.rows, ex.perfil.columns) == (2, 2)
    assert entorno_ui.carteles == []


@pytest.mark.parametrize("nombre", ["chico.xls", "chico.xlsb"])
def test_xls_y_xlsb_se_exploran(entorno_ui, fixtures_generados, nombre):
    ex = _abrir(entorno_ui, Seleccion(fixtures_generados / "formatos" / nombre, None, None, 0))
    assert (ex.perfil.rows, ex.perfil.columns) == (12, 4)
    assert ex.panel_datos.grilla.filas_mostradas() == 12


def test_archivo_ilegible_da_mensaje_amigable_y_cierra(entorno_ui, fixtures_generados):
    ex = abrir_explorador(
        entorno_ui.raiz,
        Seleccion(fixtures_generados / "fase1" / "danado.xlsx", None, None, 0),
    )
    entorno_ui.esperar(lambda: ex.cerrada)
    assert len(entorno_ui.carteles) == 1
    tipo, _, texto = entorno_ui.carteles[0]
    assert tipo == "showerror" and "Traceback" not in texto


# ---------------------------------------------------------------- cartel de memoria
def test_cartel_de_memoria_no_carga(entorno_ui, fixtures_generados, monkeypatch):
    monkeypatch.setattr(platform_utils, "available_memory", lambda: 5_000_000)
    lecturas = []
    monkeypatch.setattr(selector, "read_table", lambda *a, **k: lecturas.append(a), raising=False)
    aceptadas = []
    sel = selector.abrir_selector(
        entorno_ui.raiz, aceptadas.append, fixtures_generados / "explorar" / "para_explorar.csv"
    )
    entorno_ui.esperar(lambda: sel.cerrada)
    assert len(entorno_ui.carteles) == 1
    tipo, titulo, texto = entorno_ui.carteles[0]
    assert tipo == "showwarning" and titulo == "Archivo muy grande"
    assert "libre" in texto and "Cerrá otros programas" in texto and "," in texto
    assert sel.path is None and aceptadas == [] and lecturas == []


def test_con_memoria_suficiente_el_selector_carga(entorno_ui, fixtures_generados, monkeypatch):
    monkeypatch.setattr(platform_utils, "available_memory", lambda: 8 * 1024**3)
    sel = selector.abrir_selector(
        entorno_ui.raiz, lambda s: None, fixtures_generados / "explorar" / "para_explorar.csv"
    )
    entorno_ui.esperar(lambda: sel.listo)
    assert entorno_ui.carteles == []
    sel.cancelar()


# ---------------------------------------------------------------- CSV de 2M: cancelar
def _boton_cancelar(raiz: tk.Tk) -> ttk.Button | None:
    def buscar(w):
        for hijo in w.winfo_children():
            if isinstance(hijo, ttk.Button) and str(hijo.cget("text")) == "Cancelar":
                return hijo
            hallado = buscar(hijo)
            if hallado is not None:
                return hallado
        return None

    return buscar(raiz)


def _hilos_worker() -> list[threading.Thread]:
    return [t for t in threading.enumerate() if t.name == "worker" and t.is_alive()]


@pytest.mark.parametrize("espera", [1.0, 8.0], ids=["leyendo", "perfilando"])
def test_cancelar_el_analisis_del_csv_de_2m_cierra_en_menos_de_2_segundos(
    entorno_ui, fixtures_grandes, espera
):
    ex = abrir_explorador(
        entorno_ui.raiz, Seleccion(fixtures_grandes / "rendimiento_2M.csv", None, CSV_AR, 0)
    )
    entorno_ui.bombear(espera)
    assert ex.analizando and not ex.cerrada
    boton = _boton_cancelar(entorno_ui.raiz)
    assert boton is not None, "no apareció la ventana de progreso con el botón Cancelar"
    t0 = time.monotonic()
    boton.invoke()
    entorno_ui.esperar(lambda: ex.cerrada, 5)
    cierre = time.monotonic() - t0
    entorno_ui.esperar(lambda: not _hilos_worker(), 30)
    hilo_muerto = time.monotonic() - t0
    print(
        f"\ncancelar a los {espera} s: ventana cerrada en {cierre:.2f} s, "
        f"hilo terminado en {hilo_muerto:.2f} s"
    )
    assert cierre <= 2.0
    assert hilo_muerto <= 5.0, f"el hilo siguió trabajando {hilo_muerto:.1f} s"
    assert entorno_ui.carteles == []


# ------------------------------------------------------ cerrar la app con análisis en curso
GUION = textwrap.dedent(
    """
    import sys, time
    from pathlib import Path
    from excel_tool.core.io import CsvOptions
    from excel_tool.ui.app import crear_ventana
    from excel_tool.ui.explorador import abrir_explorador
    from excel_tool.ui.selector import Seleccion

    modo, ruta, espera = sys.argv[1], Path(sys.argv[2]), float(sys.argv[3])
    raiz = crear_ventana()
    ex = abrir_explorador(raiz, Seleccion(ruta, None, CsvOptions(";", "utf-8-sig"), 0))
    def cerrar():
        if modo == "explorador":
            ex.cerrar()
            raiz.after(1500, raiz.destroy)
        else:
            raiz.destroy()
    raiz.after(int(espera * 1000), cerrar)
    raiz.mainloop()
    print("FIN-LIMPIO")
    """
)


@pytest.mark.parametrize("modo", ["explorador", "ventana_principal"])
@pytest.mark.parametrize("espera", [1.5, 10.0], ids=["leyendo", "perfilando"])
def test_cerrar_la_app_con_un_analisis_en_curso(
    tmp_path, fixtures_grandes, modo, espera, monkeypatch
):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla")
    guion = tmp_path / "cerrar.py"
    guion.write_text(GUION, encoding="utf-8")
    entorno = {
        **os.environ,
        "XDG_DATA_HOME": str(tmp_path / "datos"),
        "LOCALAPPDATA": str(tmp_path / "datos"),
    }
    t0 = time.monotonic()
    proceso = subprocess.run(
        [
            sys.executable,
            str(guion),
            modo,
            str(fixtures_grandes / "rendimiento_2M.csv"),
            str(espera),
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=120,
        env=entorno,
    )
    duracion = time.monotonic() - t0
    print(f"\ncerrar ({modo}, a los {espera} s): código {proceso.returncode}, {duracion:.1f} s")
    assert proceso.returncode == 0, proceso.stderr[-2000:]
    assert "FIN-LIMPIO" in proceso.stdout
    assert "Traceback" not in proceso.stderr and "Fatal" not in proceso.stderr
    assert proceso.stderr.strip() == ""
    assert duracion < espera + 30
