import os
import sys
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from types import SimpleNamespace

import pandas as pd
import pytest

from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io import CsvOptions
from excel_tool.ui import explorador, exportar, selector

from .perfiles_falsos import columna, hojas_falsas, perfil_falso

pytestmark = pytest.mark.pantalla

REGIONES = ["Norte", "NORTE", "Norte ", "Sur", "Sur", "", "Este"]
IMPORTES = ["10", "20", "abc", "5", "100", "7", ""]
CIUDADES = ["A", "A", "B", "B", "B", "A", "C"]


def _esperar(raiz, condicion, segundos=10):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        raiz.update()
        if condicion():
            return
        time.sleep(0.02)
    raise AssertionError("se agotó la espera")


def _tabla(regiones=REGIONES, importes=IMPORTES, ciudades=CIUDADES):
    return pd.DataFrame(
        {"Región": regiones, "Importe": importes, "Ciudad": ciudades}, dtype="string[pyarrow]"
    )


def _perfil(df):
    perfil = perfil_falso()
    perfil.rows = len(df)
    perfil.columns_profile = [
        columna("Región", 0, "texto"),
        columna("Importe", 1, "entero"),
        columna("Ciudad", 2, "texto"),
    ]
    perfil.columns = 3
    return perfil


@pytest.fixture
def carteles(monkeypatch):
    dichos = SimpleNamespace(info=[], error=[], guardar=None)
    monkeypatch.setattr(messagebox, "showinfo", lambda *a, **k: dichos.info.append(a))
    monkeypatch.setattr(messagebox, "showerror", lambda *a, **k: dichos.error.append(a))
    monkeypatch.setattr(messagebox, "showwarning", lambda *a, **k: dichos.error.append(a))
    monkeypatch.setattr(filedialog, "asksaveasfilename", lambda **k: dichos.guardar or "")
    return dichos


@pytest.fixture
def raiz(tmp_path, monkeypatch, carteles):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos"))
    monkeypatch.setattr(explorador, "list_sheets", lambda path: hojas_falsas())
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


def _abrir(raiz, monkeypatch, df=None):
    df = _tabla() if df is None else df
    perfil = _perfil(df)
    monkeypatch.setattr(
        explorador, "read_table", lambda path, opciones: SimpleNamespace(sheets=[], df=df)
    )
    monkeypatch.setattr(
        explorador,
        "perfilar",
        lambda leido, *, on_column=None, progress=None, cancel=None, sheets=None: perfil,
    )
    seleccion = selector.Seleccion(perfil.file.path, None, CsvOptions(";", "utf-8-sig"), 0)
    ex = explorador.abrir_explorador(raiz, seleccion)
    _esperar(raiz, lambda: ex.perfil is not None)
    return ex


def _calcular(raiz, ex, columnas):
    pd_ = ex.panel_distintos
    pd_.elegir_columnas(columnas)
    pd_.calcular()
    _esperar(raiz, lambda: pd_.resultado is not None)
    return pd_


def _filas_tabla(tabla):
    return [tabla.item(i, "values") for i in tabla.get_children()]


def _esperar_filas(raiz, ex, cantidad):
    _esperar(raiz, lambda: len(ex.panel_datos.filas) == cantidad)


def test_distintos_una_columna(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    assert ex.panel_distintos.resultado is None  # nunca calcula solo
    p = _calcular(raiz, ex, ["Ciudad"])
    filas = _filas_tabla(p.tabla)
    assert filas == [("B", "3", "42,9 %"), ("A", "3", "42,9 %"), ("C", "1", "14,3 %")] or (
        filas[0][1] == "3" and filas[-1] == ("C", "1", "14,3 %")
    )
    assert "3 valores distintos en 7 filas" in p.etiqueta_resumen.cget("text")


def test_distintos_dos_columnas(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    p = _calcular(raiz, ex, ["Ciudad", "Región"])
    assert list(p.tabla["columns"]) == ["c0", "c1", "c2", "c3"]
    assert [p.tabla.heading(c, "text") for c in p.tabla["columns"]] == [
        "Región",
        "Ciudad",
        "Cantidad",
        "Porcentaje",
    ]
    assert len(p.tabla.get_children()) == p.resultado.distinct_total


def test_variantes_y_juntarlas(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    p = _calcular(raiz, ex, ["Región"])
    assert "parecen el mismo escrito distinto" in p.etiqueta_variantes.cget("text")
    assert "«Norte»" in p.etiqueta_variantes.cget("text")
    p.var_juntar.set(True)
    p.resultado = None
    p.calcular()
    _esperar(raiz, lambda: p.resultado is not None)
    assert p.etiqueta_variantes.cget("text") == ""
    valores = {f[0]: f[1] for f in _filas_tabla(p.tabla)}
    assert valores["Norte"] == "3"


def test_vacios_orden_y_buscar(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    p = _calcular(raiz, ex, ["Región"])
    assert "(vacío)" in [f[0] for f in _filas_tabla(p.tabla)]
    p.var_vacios.set(False)
    p.var_orden.set("alfabetico")
    p.var_buscar.set("sur")
    p.resultado = None
    p.calcular()
    _esperar(raiz, lambda: p.resultado is not None)
    assert [f[0] for f in _filas_tabla(p.tabla)] == ["Sur"]


def test_truncado(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    p = ex.panel_distintos
    from excel_tool.core import explore

    original = explore.value_counts
    monkeypatch.setattr(
        "excel_tool.ui.distintos.value_counts",
        lambda df, cols, opc, **kw: original(df, cols, explore.DistinctOptions(limit=2), **kw),
    )
    p.elegir_columnas(["Ciudad"])
    p.calcular()
    _esperar(raiz, lambda: p.resultado is not None)
    assert p.resultado.truncated
    assert len(p.tabla.get_children()) == 2
    assert "muy larga" in p.etiqueta_aviso.cget("text")


def test_doble_clic_lleva_a_datos_filtrado_y_se_quita(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    d = ex.panel_datos
    assert len(d.filas) == 7
    p = _calcular(raiz, ex, ["Ciudad"])
    iid = next(i for i in p.tabla.get_children() if p.tabla.item(i, "values")[0] == "B")
    p.tabla.selection_set(iid)
    raiz.update()
    p.boton_ver.invoke()
    _esperar_filas(raiz, ex, 3)
    assert ex.pestanas.select() == str(ex.pestana_datos)
    assert d.grilla.filas_mostradas() == 3
    assert d.etiqueta_conteo.cget("text") == "Mostrando 3 de 7 filas"
    assert "«B»" in d.aviso_valor
    assert d.marco_aviso.winfo_manager() == "pack"
    d.boton_quitar_filtro.invoke()
    _esperar_filas(raiz, ex, 7)
    assert d.aviso_valor is None
    assert d.grilla.filas_mostradas() == 7


def test_doble_clic_real_en_una_fila(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    ex.pestanas.select(ex.pestana_distintos)
    ex.update()
    p = _calcular(raiz, ex, ["Ciudad"])
    ex.update()
    primera = p.tabla.get_children()[0]
    p.tabla.selection_set(primera)
    x, y, ancho, alto = p.tabla.bbox(primera, "c0")
    for _ in range(2):
        for evento in ("<ButtonPress-1>", "<ButtonRelease-1>"):
            p.tabla.event_generate(evento, x=x + 3, y=y + alto // 2)
    _esperar(raiz, lambda: ex.panel_datos.aviso_valor is not None)


def test_gancho_de_columnas_calcula_solo_esa_columna(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    ex.pestanas.select(ex.pestana_columnas)
    ex.seleccionar_columna("Ciudad")
    assert ex.boton_distintos.winfo_manager() == "pack"
    ex.boton_distintos.invoke()
    assert ex.pestanas.select() == str(ex.pestana_distintos)
    _esperar(raiz, lambda: ex.panel_distintos.resultado is not None)
    assert ex.panel_distintos.resultado.columns == ["Ciudad"]


def test_filtro_contiene(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    d = ex.panel_datos
    d.agregar_condicion("Región", "contiene", "nor")
    d.aplicar()
    _esperar_filas(raiz, ex, 3)
    assert d.etiqueta_conteo.cget("text") == "Mostrando 3 de 7 filas"
    assert d.grilla.filas_mostradas() == 3
    nombres = [o for o in d.condiciones[0].combo_operacion.cget("values")]
    assert "no está vacío" in nombres and "es mayor que" in nombres


def test_filtro_esta_vacio_ignora_valor(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    d = ex.panel_datos
    fila = d.agregar_condicion("Región", "está vacío", "lo que sea")
    assert str(fila.entrada.cget("state")) == "disabled"
    d.aplicar()
    _esperar_filas(raiz, ex, 1)
    d.quitar_condicion(fila)
    d.aplicar()
    _esperar_filas(raiz, ex, 7)


def test_filtro_mayor_con_valor_invalido(raiz, monkeypatch, carteles):
    ex = _abrir(raiz, monkeypatch)
    d = ex.panel_datos
    d.agregar_condicion("Importe", "es mayor que", "xyz")
    d.aplicar()
    _esperar(raiz, lambda: carteles.error)
    assert "no es un número válido" in carteles.error[0][1]
    assert len(d.filas) == 7  # no se rompió nada
    assert str(d.boton_aplicar.cget("state")) != "disabled" or d.boton_aplicar.instate(
        ["!disabled"]
    )
    d.condiciones[0].var_valor.set("8")
    d.aplicar()
    _esperar_filas(raiz, ex, 3)  # 10, 20 y 100


def test_orden_por_encabezado(raiz, monkeypatch):
    ex = _abrir(raiz, monkeypatch)
    d = ex.panel_datos
    d.grilla._clic_encabezado("Importe")
    _esperar(raiz, lambda: d._orden == ("Importe", True))
    primera = d.grilla.arbol.item(d.grilla.arbol.get_children()[0], "values")
    assert primera[2] == "5"
    d.grilla._clic_encabezado("Importe")
    _esperar(raiz, lambda: d._orden == ("Importe", False))
    primera = d.grilla.arbol.item(d.grilla.arbol.get_children()[0], "values")
    assert primera[2] == "100"


def test_la_grilla_nunca_pasa_de_una_pagina(raiz, monkeypatch):
    n = 2500
    df = _tabla(["x"] * n, [str(i) for i in range(n)], ["c"] * n)
    ex = _abrir(raiz, monkeypatch, df)
    g = ex.panel_datos.grilla
    assert g.total_filas == n
    assert g.filas_mostradas() == g.page_size == 1000
    g.siguiente()
    g.ultima()
    assert g.filas_mostradas() == 500
    assert g.cantidad_paginas == 3


def test_exportar_lista_llama_al_core(raiz, monkeypatch, carteles, tmp_path):
    llamadas = []
    monkeypatch.setattr(exportar, "guardar_distintos", lambda *a: llamadas.append(a))
    ex = _abrir(raiz, monkeypatch)
    p = ex.panel_distintos
    assert str(p.boton_exportar.cget("state")) == "disabled"
    _calcular(raiz, ex, ["Ciudad"])
    carteles.guardar = str(tmp_path / "lista.csv")
    p.boton_exportar.invoke()
    _esperar(raiz, lambda: carteles.info)
    assert len(llamadas) == 1
    resultado, ruta, formato = llamadas[0]
    assert resultado is p.resultado and ruta == tmp_path / "lista.csv"
    assert formato == "csv_excel"
    carteles.guardar = str(tmp_path / "lista.xlsx")
    p.exportar()
    _esperar(raiz, lambda: len(llamadas) == 2)
    assert llamadas[1][2] == "xlsx"


def test_exportar_cancelando_el_dialogo_no_hace_nada(raiz, monkeypatch, carteles):
    llamadas = []
    monkeypatch.setattr(exportar, "guardar_distintos", lambda *a: llamadas.append(a))
    ex = _abrir(raiz, monkeypatch)
    _calcular(raiz, ex, ["Ciudad"])
    ex.panel_distintos.exportar()
    raiz.update()
    assert llamadas == []


def test_guardar_informe_llama_al_core_con_los_distintos(raiz, monkeypatch, carteles, tmp_path):
    llamadas = []
    monkeypatch.setattr(exportar, "guardar_informe", lambda *a: llamadas.append(a))
    ex = _abrir(raiz, monkeypatch)
    _calcular(raiz, ex, ["Ciudad"])
    carteles.guardar = str(tmp_path / "informe.xlsx")
    ex.boton_informe.invoke()
    _esperar(raiz, lambda: carteles.info)
    perfil, distintos, ruta = llamadas[0]
    assert perfil is ex.perfil
    assert distintos == [ex.panel_distintos.resultado]
    assert ruta == tmp_path / "informe.xlsx"


def test_exportar_sin_funcion_en_el_core_avisa_amable(raiz, monkeypatch, carteles, tmp_path):
    monkeypatch.setattr(exportar, "MODULOS", ("excel_tool.core.no_existe",))
    ex = _abrir(raiz, monkeypatch)
    carteles.guardar = str(tmp_path / "informe.xlsx")
    ex.guardar_informe()
    _esperar(raiz, lambda: carteles.error)
    assert "todavía no está disponible" in carteles.error[0][1]


def test_error_del_core_al_exportar_se_muestra(raiz, monkeypatch, carteles, tmp_path):
    def falla(*a):
        raise ExcelToolError("Cerrá el archivo en Excel y volvé a intentar.")

    monkeypatch.setattr(exportar, "guardar_informe", falla)
    ex = _abrir(raiz, monkeypatch)
    carteles.guardar = str(tmp_path / "informe.xlsx")
    ex.guardar_informe()
    _esperar(raiz, lambda: carteles.error)
    assert carteles.error[0][1].startswith("Cerrá el archivo")
