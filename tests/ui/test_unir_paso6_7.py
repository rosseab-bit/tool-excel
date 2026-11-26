"""Tests con pantalla (xvfb) de los pasos 6 y 7 del asistente y del explorador sobre una tabla
en memoria (F5-14). Usan el escenario real de 3 archivos y el core de verdad."""

from __future__ import annotations

import json
import os
import shutil
import sys
import time
import tkinter as tk
from pathlib import Path

import pandas as pd
import pytest

from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy
from excel_tool.core.io import ReadOptions, list_sheets, read_table
from excel_tool.core.join import SourceSpec
from excel_tool.core.limits import SizeCheck
from excel_tool.core.recipe import load_recipe
from excel_tool.ui import explorador, guardar
from excel_tool.ui.unir import tamano
from excel_tool.ui.unir.asistente import ALTO_MAXIMO_VENTANA, Asistente
from excel_tool.ui.unir.estado import EstadoUnion, Fuente, FuenteSecundaria, leer_fuente
from excel_tool.ui.unir.pasos import paso7_guardar
from excel_tool.ui.unir.pasos.paso6_resumen import PasoResumenFinal
from excel_tool.ui.unir.widgets_archivos import ver_informacion

pytestmark = pytest.mark.pantalla

CLAVE = [("ID cliente", "ID cliente")]
VENTAS_RESUMEN = AggregationSpec(
    [GroupBy("ID cliente")], [Aggregation("Importe", "sum", "Total comprado")]
)
COBRANZAS_RESUMEN = AggregationSpec(
    [GroupBy("Nro Cliente")], [Aggregation("Monto", "sum", "Total cobrado")]
)


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


def _fuente(carpeta, nombre, columnas=()):
    spec = SourceSpec(nombre, carpeta / nombre, ReadOptions(), list(columnas))
    return Fuente(spec, read_table(spec.path).df)


def _sec(carpeta, nombre, pares, **kw):
    base = _fuente(carpeta, nombre)
    base.spec = SourceSpec(
        nombre, base.spec.path, base.spec.read_options, [str(c) for c in base.df.columns]
    )
    return FuenteSecundaria(base.spec, base.df, key_pairs=pares, **kw)


def _principal(carpeta):
    return _fuente(carpeta, "clientes.xlsx", ["ID cliente", "Nombre", "Región"])


def _estado(carpeta, ventas="detail", cobranzas=None, ventas_resumen=VENTAS_RESUMEN):
    secundarios = [
        _sec(
            carpeta,
            "ventas.csv",
            CLAVE,
            mode=ventas,
            summary=ventas_resumen if ventas == "summary" else None,
        )
    ]
    if cobranzas:
        secundarios.append(
            _sec(
                carpeta,
                "cobranzas.xlsx",
                [("ID cliente", "Nro Cliente")],
                mode=cobranzas,
                summary=COBRANZAS_RESUMEN,
            )
        )
    return EstadoUnion(
        principal=_principal(carpeta), claves=["ID cliente"], secundarios=secundarios
    )


def _esperar(raiz, condicion, segundos=25):
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


def _al_paso_7(raiz, estado, desde=5):
    asistente = Asistente(raiz, estado, desde)
    raiz.update()
    while asistente.indice < 6:
        asistente.siguiente()
        raiz.update()
    paso = asistente.paso_actual
    _esperar(raiz, lambda: not paso.pendiente)
    return asistente, paso


def _unir(raiz, paso):
    paso.boton_unir.invoke()
    _esperar(raiz, lambda: paso.resultado is not None)
    return paso.resultado


def _guardar_resultado(raiz, monkeypatch, paso, destino, modo=None, filas=None):
    paso.guardar_resultado()
    v = paso.ventana_guardar
    assert v is not None
    _esperar(raiz, lambda: v._check is not None and not v._ocupado)
    if modo:
        v.v_modo.set(modo)
        v.v_filas.set(filas)
        v._n_editado = True
        v._etapa_b()
        _esperar(raiz, lambda: v.puede_guardar())
    monkeypatch.setattr(guardar.filedialog, "asksaveasfilename", lambda **k: str(destino))
    v._guardar()
    _esperar(raiz, lambda: v.salida_usada is not None)
    return v


# ------------------------------------------------------------------ escenario real
def test_detalle_88_filas_guardar_xlsx_y_releer(raiz, carpeta, monkeypatch, tmp_path):
    asistente, paso = _al_paso_7(raiz, _estado(carpeta))
    assert "88 filas" in str(paso.etiqueta_total.cget("text"))
    textos = " ".join(_textos(paso.marco_archivos))
    assert "ventas.csv" in textos and "27 de 30" in textos and "3 sin pareja" in textos
    assert "Control de filas" in str(paso.etiqueta_control.cget("text"))
    assert not asistente.boton_siguiente.winfo_ismapped()
    res = _unir(raiz, paso)
    assert len(res.df) == 88
    assert "El resultado tiene 88 filas y" in str(paso.etiqueta_filas.cget("text"))
    assert paso.grilla.total_filas == 88
    salida = tmp_path / "unido.xlsx"
    _guardar_resultado(raiz, monkeypatch, paso, salida)
    assert len(read_table(salida).df) == 88
    assert str(asistente.boton_cancelar.cget("text")) == "Terminar"
    asistente.boton_cancelar.invoke()  # Terminar: cierra sin preguntar
    assert asistente.cerrada


def test_dos_resumenes_dan_30_filas(raiz, carpeta):
    _, paso = _al_paso_7(raiz, _estado(carpeta, "summary", "summary"))
    assert "30 filas" in str(paso.etiqueta_total.cget("text"))
    assert "mismas filas que el archivo principal" in str(paso.etiqueta_control.cget("text"))
    res = _unir(raiz, paso)
    assert len(res.df) == 30
    assert {"Total comprado", "Total cobrado"} <= set(res.df.columns)


def test_resumen_por_cliente_y_articulo(raiz, carpeta):
    spec = AggregationSpec(
        [GroupBy("ID cliente"), GroupBy("Artículo")], [Aggregation("Importe", "sum", "Total")]
    )
    ventas = read_table(carpeta / "ventas.csv").df
    ids = set(read_table(carpeta / "clientes.xlsx").df["ID cliente"])
    pares = ventas[ventas["ID cliente"].isin(ids)][["ID cliente", "Artículo"]].drop_duplicates()
    esperadas = len(pares) + len(ids - set(pares["ID cliente"]))
    _, paso = _al_paso_7(raiz, _estado(carpeta, "summary", ventas_resumen=spec))
    assert "agrupa por más que la columna en común" in str(paso.etiqueta_control.cget("text"))
    res = _unir(raiz, paso)
    assert len(res.df) == esperadas
    assert "Artículo" in res.df.columns


def test_resumen_final_por_region_desde_el_paso_6(raiz, carpeta, fixtures_esperados):
    estado = _estado(carpeta, "summary")
    asistente = Asistente(raiz, estado, 5)
    raiz.update()
    paso6 = asistente.paso_actual
    assert str(paso6.radio_no.cget("text")) == "No, dejarlo así"
    assert paso6.var.get() == "no"
    paso6.radio_si.invoke()
    _esperar(raiz, lambda: paso6.editor is not None)
    editor = paso6.editor
    assert {"Región", "Total comprado", "Nombre"} <= set(editor.nombres)  # de varios archivos
    editor.var_check["Región"].set(True)
    editor.agregar_calculo().poner("Total comprado", "sum", None)
    editor.agregar_calculo().poner("", "count", "Clientes")
    paso6.boton_listo.invoke()
    assert estado.resumen_final is not None
    assert [g.column for g in estado.resumen_final.group_by] == ["Región"]
    asistente.siguiente()
    paso = asistente.paso_actual
    _esperar(raiz, lambda: not paso.pendiente)
    res = _unir(raiz, paso)
    esperado = json.loads(
        (fixtures_esperados / "escenario_3" / "resumen_final_por_region.json").read_text(
            encoding="utf-8"
        )
    )
    por_region = dict(zip(res.df["Región"], res.df["Suma de Total comprado"], strict=True))
    comparado = {r: str(v).replace(",", ".") for r, v in por_region.items()}
    assert {r: float(v) for r, v in comparado.items()} == pytest.approx(
        {r: float(v) for r, v in esperado["total_importe_por_region"].items()}
    )
    clientes = dict(zip(res.df["Región"], res.df["Clientes"], strict=True))
    assert {r: int(v) for r, v in clientes.items()} == esperado["clientes_por_region"]


def test_resumen_final_por_vendedor_y_articulo_de_dos_archivos(raiz, carpeta):
    estado = _estado(carpeta)
    paso6 = PasoResumenFinal(raiz, estado, None)
    paso6.pack()
    paso6.al_entrar()
    paso6.radio_si.invoke()
    _esperar(raiz, lambda: paso6.editor is not None)
    paso6.editor.var_check["Región"].set(True)
    paso6.editor.var_check["Artículo"].set(True)
    paso6.boton_listo.invoke()
    assert [g.column for g in estado.resumen_final.group_by] == ["Región", "Artículo"]


def test_paso_6_validar_y_no_dejarlo_asi(raiz, carpeta):
    estado = _estado(carpeta)
    paso = PasoResumenFinal(raiz, estado, None)
    paso.pack()
    paso.al_entrar()
    assert paso.validar() is None
    paso.var.set("si")
    assert "Armá el agrupamiento" in paso.validar()
    estado.resumen_final = AggregationSpec([GroupBy("Región")], [])
    assert paso.validar() is None
    paso.var.set("no")
    paso.al_salir()
    assert estado.resumen_final is None


def test_paso_6_descarta_agrupamiento_si_ya_no_existen_las_columnas(raiz, carpeta):
    estado = _estado(carpeta)
    estado.resumen_final = AggregationSpec([GroupBy("Columna que no existe")], [])
    paso = PasoResumenFinal(raiz, estado, None)
    paso.pack()
    paso.al_entrar()
    assert estado.resumen_final is None
    assert paso.var.get() == "no"
    assert "armarlo de nuevo" in str(paso.etiqueta_estado.cget("text"))


def test_paso_6_cancelar_el_editor_vuelve_a_no(raiz, carpeta):
    estado = _estado(carpeta)
    paso = PasoResumenFinal(raiz, estado, None)
    paso.pack()
    paso.al_entrar()
    paso.radio_si.invoke()
    _esperar(raiz, lambda: paso.editor is not None)
    paso.dialogo.event_generate("<Destroy>")
    paso._cerrar_dialogo()
    paso._sin_resumen_si_vacio()
    assert paso.var.get() == "no"


# ------------------------------------------------------------------ varias hojas
def test_guardar_en_varias_hojas_detalle_resumen_y_sin_coincidencia(
    raiz, carpeta, monkeypatch, tmp_path
):
    estado = _estado(carpeta)
    estado.resumen_final = AggregationSpec([GroupBy("Región")], [Aggregation("", "count", "N")])
    _, paso = _al_paso_7(raiz, estado)
    _unir(raiz, paso)
    assert paso.boton_libro.winfo_ismapped()
    salida = tmp_path / "libro.xlsx"
    monkeypatch.setattr(paso7_guardar.filedialog, "asksaveasfilename", lambda **k: str(salida))
    paso.guardar_libro()
    _esperar(
        raiz,
        lambda: salida.exists() and "se guardó" in str(paso.etiqueta_guardado.cget("text")).lower(),
    )
    hojas = [h.name for h in list_sheets(salida)]
    assert hojas == ["Detalle", "Resumen", "Sin coincidencia en ventas"]
    assert len(read_table(salida, ReadOptions(sheet="Detalle")).df) == 88
    assert len(read_table(salida, ReadOptions(sheet="Resumen")).df) == 5


def test_libro_no_aparece_si_no_hay_nada_mas_que_una_tabla(raiz, carpeta):
    estado = _estado(carpeta, "summary", "summary")
    _, paso = _al_paso_7(raiz, estado)
    res = _unir(raiz, paso)
    if all(len(df) == 0 for df in res.unmatched.values()):
        assert not paso.boton_libro.winfo_ismapped()
    else:  # ventas trae IDs que no están en clientes: sí hay hoja extra
        assert paso.boton_libro.winfo_ismapped()


def test_libro_en_rojo_se_explica_y_no_pide_ruta(raiz, carpeta, monkeypatch):
    estado = _estado(carpeta)
    _, paso = _al_paso_7(raiz, estado)
    _unir(raiz, paso)
    rojo = SizeCheck("obligatorio", ["La hoja «Resultado» no entra."], 2, 1000, [])
    monkeypatch.setattr(paso7_guardar, "plan_workbook", lambda hojas, destino: rojo)
    avisos = []
    monkeypatch.setattr(paso7_guardar.messagebox, "showwarning", lambda *a, **k: avisos.append(a))
    monkeypatch.setattr(
        paso7_guardar.filedialog,
        "asksaveasfilename",
        lambda **k: pytest.fail("no tendría que pedir ruta"),
    )
    paso.guardar_libro()
    _esperar(raiz, lambda: bool(avisos))
    assert "por separado" in avisos[0][1] and "no entra" in avisos[0][1]


# ------------------------------------------------------------------ achicar
def test_achicar_el_resultado_con_semaforo_en_rojo(raiz, carpeta, monkeypatch):
    real = tamano.check

    def check_falso(est, destino, formato, **kw):
        if destino == "excel" and est.rows > 50:
            return SizeCheck("obligatorio", ["No entra en una hoja de Excel."], 2, 50, [])
        return real(est, destino, formato, **kw)

    monkeypatch.setattr(tamano, "check", check_falso)
    estado = _estado(carpeta)
    _, paso = _al_paso_7(raiz, estado)
    assert "🔴" in str(paso.etiqueta_semaforo.cget("text"))
    assert paso.marco_achicar.winfo_ismapped()
    textos = [str(b.cget("text")) for b in paso.botones_reduccion]
    boton = next(b for b in paso.botones_reduccion if "ventas.csv a resumen" in b.cget("text"))
    assert "🟢" in str(boton.cget("text")) and "30 filas" in str(boton.cget("text")), textos
    boton.invoke()
    assert estado.secundarios[0].mode == "summary"
    assert estado.secundarios[0].summary is not None
    _esperar(
        raiz, lambda: not paso.pendiente and "30 filas" in str(paso.etiqueta_total.cget("text"))
    )
    assert not paso.marco_achicar.winfo_ismapped()
    assert len(_unir(raiz, paso).df) == 30


# ------------------------------------------------------------------ receta
def test_receta_guardada_se_abre_y_da_el_mismo_resultado(raiz, carpeta, monkeypatch, tmp_path):
    estado = _estado(carpeta)
    _, paso = _al_paso_7(raiz, estado)
    primero = _unir(raiz, paso)
    _guardar_resultado(raiz, monkeypatch, paso, tmp_path / "salida.xlsx", "rows", "50")
    assert estado.salida is not None and estado.salida.split.rows_per_part == 50
    assert estado.salida.column_types is None
    receta = tmp_path / "mi receta.json"
    monkeypatch.setattr(paso7_guardar.filedialog, "asksaveasfilename", lambda **k: str(receta))
    paso.guardar_receta()
    assert receta.exists()
    assert "receta quedó guardada" in str(paso.etiqueta_guardado.cget("text"))

    carga = load_recipe(receta)
    nuevo = EstadoUnion.desde_receta(carga, leer=leer_fuente)
    assert nuevo.salida is not None and nuevo.salida.split.rows_per_part == 50
    _, paso2 = _al_paso_7(raiz, nuevo, desde=2)
    segundo = _unir(raiz, paso2)
    pd.testing.assert_frame_equal(primero.df, segundo.df)
    # Guardar arranca con la división y el formato de la receta
    paso2.guardar_resultado()
    v = paso2.ventana_guardar
    _esperar(raiz, lambda: v._check is not None)
    assert v.v_modo.get() == "rows" and v.v_filas.get() == "50"
    assert v.v_formato.get() == "xlsx" and v.v_destino.get() == "excel"


def test_receta_incompleta_muestra_el_motivo(raiz, carpeta, monkeypatch):
    estado = _estado(carpeta)
    _, paso = _al_paso_7(raiz, estado)
    estado.secundarios[0].mode = None
    errores = []
    monkeypatch.setattr(paso7_guardar.messagebox, "showerror", lambda *a, **k: errores.append(a))
    paso.guardar_receta()
    assert errores and "ventas.csv" in errores[0][1]


# ------------------------------------------------------------------ explorador sobre tabla
def test_ver_informacion_de_un_archivo_no_lo_relee(raiz, carpeta, monkeypatch):
    fuente = leer_fuente(SourceSpec("ventas.csv", carpeta / "ventas.csv", ReadOptions(), []))

    def prohibido(*a, **k):
        raise AssertionError("no tendría que volver a leer el archivo")

    monkeypatch.setattr(explorador, "read_table", prohibido)
    abiertos = []
    original = explorador.abrir_explorador

    def espiar(*a, **k):
        ex = original(*a, **k)
        abiertos.append(ex)
        return ex

    from excel_tool.ui.unir import widgets_archivos

    monkeypatch.setattr(widgets_archivos, "abrir_explorador", espiar)
    ver_informacion(raiz, fuente)
    ex = abiertos[0]
    _esperar(raiz, lambda: ex.perfil is not None)
    assert ex.df is fuente.df
    assert ex.perfil is fuente.perfil
    assert not ex.combo_hoja.winfo_ismapped()


def test_ver_informacion_del_resultado(raiz, carpeta, monkeypatch):
    _, paso = _al_paso_7(raiz, _estado(carpeta))
    res = _unir(raiz, paso)
    monkeypatch.setattr(
        explorador, "read_table", lambda *a, **k: pytest.fail("no tendría que leer nada")
    )
    paso.ver_informacion()
    ex = paso.explorador
    _esperar(raiz, lambda: ex.perfil is not None)
    assert ex.df is res.df
    assert ex.perfil.rows == 88 and ex.perfil.columns == len(res.df.columns)
    assert str(ex.title()) == "Información de Resultado de la unión"
    assert "El archivo" not in _textos(ex.contenido_general)


# ------------------------------------------------------------------ varios
def test_nombres_finales_usan_el_core(carpeta):
    estado = _estado(carpeta, "summary", "summary")
    nombres = tamano.nombres_finales(estado.plan(), estado.tablas())
    assert nombres[("clientes.xlsx", "ID cliente")] == "ID cliente"
    assert nombres[("ventas.csv", "Total comprado")] == "Total comprado"


def test_el_paso_7_entra_en_768(raiz, carpeta):
    asistente, paso = _al_paso_7(raiz, _estado(carpeta))
    _unir(raiz, paso)
    raiz.update()
    assert asistente.winfo_reqheight() <= ALTO_MAXIMO_VENTANA
    assert asistente.winfo_height() <= ALTO_MAXIMO_VENTANA
    assert asistente.boton_cancelar.winfo_ismapped()


def test_error_de_memoria_del_core_se_muestra(raiz, carpeta, monkeypatch):
    from excel_tool.core.errors import ExcelToolError

    _, paso = _al_paso_7(raiz, _estado(carpeta))

    def falla(*a, **k):
        raise ExcelToolError("No alcanza la memoria para unir estos archivos.")

    monkeypatch.setattr(paso7_guardar, "run_join", falla)
    paso.boton_unir.invoke()
    _esperar(raiz, lambda: "memoria" in str(paso.etiqueta_estado.cget("text")))
    assert paso.resultado is None


def test_se_suelta_el_resultado_al_cerrar_guardar_y_al_volver(raiz, carpeta, monkeypatch, tmp_path):
    import gc
    import weakref

    asistente, paso = _al_paso_7(raiz, _estado(carpeta))
    res = _unir(raiz, paso)
    ref = weakref.ref(res.df)
    del res
    v = _guardar_resultado(raiz, monkeypatch, paso, tmp_path / "unido.xlsx")
    v.cerrar()
    raiz.update()
    assert paso.ventana_guardar is None
    assert paso.resultado is not None  # se puede guardar de nuevo (otro destino) sin unir otra vez
    # volver atrás suelta el resultado y no queda ninguna otra referencia (grilla, ventana, cierres)
    asistente.anterior()
    raiz.update()
    assert paso.resultado is None and not paso.marco_resultado.winfo_ismapped()
    gc.collect()
    assert ref() is None
    # unir de nuevo también suelta el anterior
    asistente.siguiente()
    raiz.update()
    _esperar(raiz, lambda: not paso.pendiente)
    _unir(raiz, paso)
    ref2 = weakref.ref(paso.resultado.df)
    _unir_otra_vez = paso.resultado
    del _unir_otra_vez
    paso.boton_unir.invoke()
    assert paso.resultado is None  # soltado antes de empezar
    _esperar(raiz, lambda: paso.resultado is not None)
    gc.collect()
    assert ref2() is None
