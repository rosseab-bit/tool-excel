"""F4-10: editor de resumen, vista previa y punto de entrada "Resumir…" (con el core falso)."""

from __future__ import annotations

import importlib.util
import os
import sys
import tkinter as tk
from dataclasses import dataclass
from types import SimpleNamespace

import pandas as pd
import pytest

from excel_tool.core.errors import ExcelToolError
from excel_tool.ui import editor_resumen, guardar

from .test_exportar_semaforo import _abrir, _esperar, _tabla, _textos

pytestmark = pytest.mark.pantalla


@dataclass(frozen=True)
class GroupBy:
    column: str
    period: str | None = None


@dataclass(frozen=True)
class Aggregation:
    column: str
    func: str
    name: str | None = None


@dataclass(frozen=True)
class AggregationSpec:
    group_by: list
    aggregations: list


NOMBRES_FALSOS = {"sum": "Suma de", "count": "Cantidad de filas", "count_distinct": "Distintos de"}


def default_name(agg):
    if agg.func == "count":
        return "Cantidad de filas"
    return f"{NOMBRES_FALSOS.get(agg.func, agg.func)} {agg.column}".strip()


class CoreFalso:
    GroupBy = GroupBy
    Aggregation = Aggregation
    AggregationSpec = AggregationSpec
    default_name = staticmethod(default_name)

    def __init__(self, not_numeric=None):
        self.llamadas = []
        self.not_numeric = not_numeric or {}

    def summarize(self, df, spec, *, column_types=None, progress=None, cancel=None):
        self.llamadas.append((df, spec, column_types))
        salida = pd.DataFrame({"Fila": [str(i) for i in range(len(df.head(3)))]})
        return SimpleNamespace(
            df=salida,
            column_types={"Fila": "entero"},
            not_numeric=self.not_numeric,
            warnings=[],
        )


@pytest.fixture
def raiz(tmp_path, monkeypatch):
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        pytest.skip("no hay pantalla (en Linux usar xvfb-run -a)")
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "datos"))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "datos"))
    ventana = tk.Tk()
    yield ventana
    ventana.destroy()


@pytest.fixture
def core(monkeypatch):
    falso = CoreFalso()
    monkeypatch.setattr(editor_resumen, "_core", lambda: falso)
    return falso


TIPOS = {
    "Cliente": "texto",
    "Fecha": "fecha",
    "Importe": "mixto",
    "Cantidad": "entero",
    "Región": "texto",
}


def _df(filas=4):
    return pd.DataFrame(
        {
            "Cliente": ["a"] * filas,
            "Fecha": ["2024-03-01"] * filas,
            "Importe": ["10"] * filas,
            "Cantidad": ["1"] * filas,
            "Región": ["N"] * filas,
        },
        dtype="string[pyarrow]",
    )


def _editor(raiz, **kwargs):
    editor = editor_resumen.EditorResumen(raiz, _df(), TIPOS, **kwargs)
    editor.pack()
    raiz.update()
    return editor


def test_clave_fija_marcada_y_sin_poder_quitarse(raiz, core):
    editor = _editor(raiz, fixed_group=["Cliente"])
    assert editor.var_check["Cliente"].get() is True
    assert editor.checks["Cliente"].instate(["disabled"])
    assert "(clave)" in editor.checks["Cliente"].cget("text")
    editor.set_spec(AggregationSpec([], []))
    assert [g.column for g in editor.spec().group_by] == ["Cliente"]


def test_por_mes_arma_groupby_con_periodo(raiz, core):
    editor = _editor(raiz)
    editor.var_check["Fecha"].set(True)
    editor.var_periodo["Fecha"].set("por mes")
    assert editor.spec().group_by == [GroupBy("Fecha", "month")]
    editor.var_periodo["Fecha"].set("por día")
    assert editor.spec().group_by == [GroupBy("Fecha", None)]
    editor.var_periodo["Fecha"].set("por año")
    assert editor.spec().group_by == [GroupBy("Fecha", "year")]
    assert "Cliente" not in editor.var_periodo


def test_frase_de_agrupacion(raiz, core):
    editor = _editor(raiz)
    editor.var_check["Cliente"].set(True)
    editor.var_check["Región"].set(True)
    editor._cambio()
    assert editor.etiqueta_frase.cget("text") == "Una fila por cada Cliente y Región"


def test_sumar_no_aparece_para_texto_y_cambia_a_contar_distintos(raiz, core):
    editor = _editor(raiz)
    fila = editor.agregar_calculo()
    fila.var_operacion.set("Sumar")
    fila.var_columna.set("Cantidad")
    fila.ajustar(cambio_columna=True)
    assert "Sumar" in fila.combo_operacion.cget("values")
    fila.var_columna.set("Cliente")
    fila.ajustar(cambio_columna=True)
    assert "Sumar" not in fila.combo_operacion.cget("values")
    assert "Promedio" not in fila.combo_operacion.cget("values")
    assert fila.var_operacion.get() == "Contar distintos"
    assert fila.etiqueta_tipo.cget("text") == "Texto"


def test_columna_mixta_avisa(raiz, core):
    editor = _editor(raiz)
    fila = editor.agregar_calculo()
    fila.var_operacion.set("Sumar")
    fila.var_columna.set("Importe")
    fila.ajustar()
    assert "no son números" in fila.etiqueta_aviso.cget("text")


def test_nombre_automatico_y_editado_no_se_pisa(raiz, core):
    editor = _editor(raiz)
    fila = editor.agregar_calculo()
    fila.var_operacion.set("Sumar")
    fila.var_columna.set("Cantidad")
    fila.ajustar()
    assert fila.var_nombre.get() == "Suma de Cantidad"
    assert editor.spec().aggregations == [Aggregation("Cantidad", "sum", None)]
    fila.var_nombre.set("Total vendido")
    fila.var_columna.set("Importe")
    fila.ajustar()
    assert fila.var_nombre.get() == "Total vendido"
    assert editor.spec().aggregations == [Aggregation("Importe", "sum", "Total vendido")]


def test_contar_filas_deshabilita_la_columna(raiz, core):
    editor = _editor(raiz)
    fila = editor.agregar_calculo()
    fila.var_operacion.set("Sumar")
    fila.var_columna.set("Cantidad")
    fila.ajustar()
    fila.var_operacion.set("Contar filas")
    fila.ajustar()
    assert fila.combo_columna.instate(["disabled"])
    assert fila.var_columna.get() == ""
    assert editor.spec().aggregations == [Aggregation("", "count", None)]


def test_quitar_calculo_y_filas_incompletas(raiz, core):
    editor = _editor(raiz)
    fila = editor.agregar_calculo()
    fila.var_operacion.set("Sumar")
    fila.ajustar()
    assert editor.spec().aggregations == []
    fila.boton_quitar.invoke()
    assert editor.filas == []


def test_spec_y_set_spec_ida_y_vuelta(raiz, core):
    editor = _editor(raiz, fixed_group=["Cliente"])
    spec = AggregationSpec(
        [GroupBy("Cliente"), GroupBy("Fecha", "month")],
        [
            Aggregation("Cantidad", "sum", "Total"),
            Aggregation("", "count"),
            Aggregation("Región", "count_distinct"),
        ],
    )
    editor.set_spec(spec)
    assert editor.spec() == spec
    editor.set_spec(AggregationSpec([GroupBy("Región")], []))
    assert editor.spec() == AggregationSpec([GroupBy("Cliente"), GroupBy("Región")], [])


def test_on_change_se_llama_con_el_spec(raiz, core):
    recibidos = []
    editor = _editor(raiz, on_change=recibidos.append)
    recibidos.clear()
    editor.var_check["Región"].set(True)
    editor._cambio()
    assert recibidos[-1] == AggregationSpec([GroupBy("Región")], [])
    fila = editor.agregar_calculo()
    assert recibidos[-1].aggregations == [Aggregation("", "count", None)]
    n = len(recibidos)
    fila.var_nombre.set("Cuántas")
    assert len(recibidos) > n and recibidos[-1].aggregations[0].name == "Cuántas"
    n = len(recibidos)
    editor.set_spec(AggregationSpec([], []))
    assert len(recibidos) == n + 1


def test_buscador_solo_con_muchas_columnas(raiz, core):
    editor = _editor(raiz)
    assert not hasattr(editor, "entrada_buscar")
    columnas = {f"Col {i}": ["x"] for i in range(20)}
    grande = editor_resumen.EditorResumen(raiz, pd.DataFrame(columnas), {})
    grande.pack()
    assert hasattr(grande, "entrada_buscar")
    grande.var_buscar.set("col 1")
    raiz.update()
    visibles = [n for n, m in grande._marcos.items() if m.winfo_manager()]
    assert visibles and all("col 1" in n.lower() for n in visibles)


def test_validacion_sin_nada_que_resumir(raiz, core):
    editor = _editor(raiz)
    editor.ver_como_queda()
    assert editor.etiqueta_estado.cget("text") == editor_resumen.TEXTO_VACIO
    assert core.llamadas == []


def test_vista_previa_muestra_resultado_y_avisos(raiz, monkeypatch):
    falso = CoreFalso(not_numeric={"Importe": 3})
    monkeypatch.setattr(editor_resumen, "_core", lambda: falso)
    editor = _editor(raiz)
    editor.var_check["Cliente"].set(True)
    editor.ver_como_queda()
    _esperar(raiz, lambda: editor.resultado is not None)
    assert editor.grilla.total_filas == 3
    estado = editor.etiqueta_estado.cget("text")
    assert "3 celdas de Importe no son números y no se sumaron." in estado
    assert "Vista previa con las primeras" not in estado
    assert falso.llamadas[0][1].group_by == [GroupBy("Cliente")]


def test_vista_previa_con_muestra_de_200000(raiz, monkeypatch):
    falso = CoreFalso()
    monkeypatch.setattr(editor_resumen, "_core", lambda: falso)
    monkeypatch.setattr(editor_resumen, "MAX_FILAS_VISTA", 3)
    editor = _editor(raiz)
    editor.var_check["Cliente"].set(True)
    editor.ver_como_queda()
    _esperar(raiz, lambda: editor.resultado is not None)
    assert len(falso.llamadas[0][0]) == 3
    assert "Vista previa con las primeras" in editor.etiqueta_estado.cget("text")
    assert editor_resumen.TEXTO_MUESTRA == "Vista previa con las primeras 200.000 filas."


def test_core_ausente_da_mensaje_en_castellano(monkeypatch):
    def falla(nombre):
        raise ImportError(nombre)

    monkeypatch.setattr(editor_resumen.importlib, "import_module", falla)
    with pytest.raises(ExcelToolError) as info:
        editor_resumen._core()
    assert "todavía no está disponible" in info.value.mensaje


def test_avisos_en_singular():
    frases = editor_resumen.avisos_de(SimpleNamespace(not_numeric={"X": 1}, warnings=["Ojo"]))
    assert frases == ["1 celda de X no es un número y no se sumó.", "Ojo"]


# ---------------------------------------------------------------- punto de entrada
def test_guardar_el_resumen_pasa_column_types(raiz, monkeypatch):
    falso = CoreFalso()
    monkeypatch.setattr(editor_resumen, "_core", lambda: falso)
    registro = []
    monkeypatch.setattr(guardar, "abrir_guardar", lambda *a, **k: registro.append((a, k)))
    ventana = editor_resumen.abrir_resumen(raiz, _df(), TIPOS, "ventas")
    ventana.editor.var_check["Cliente"].set(True)
    ventana.boton_guardar.invoke()
    _esperar(raiz, lambda: registro)
    (args, kwargs) = registro[0]
    assert list(args[1].columns) == ["Fila"]
    assert kwargs["column_types"] == {"Fila": "entero"}
    assert kwargs["nombre_sugerido"] == "ventas_resumen"
    assert len(falso.llamadas[0][0]) == 4  # tabla completa


def test_guardar_el_resumen_valida(raiz, core):
    ventana = editor_resumen.abrir_resumen(raiz, _df(), TIPOS, "ventas")
    ventana.boton_guardar.invoke()
    assert ventana.editor.etiqueta_estado.cget("text") == editor_resumen.TEXTO_VACIO


def test_abrir_resumen_sin_core_avisa(raiz, monkeypatch):
    avisos = []

    def falla():
        raise ExcelToolError("no está")

    monkeypatch.setattr(editor_resumen, "_core", falla)
    monkeypatch.setattr(editor_resumen.messagebox, "showinfo", lambda *a, **k: avisos.append(a))
    assert editor_resumen.abrir_resumen(raiz, _df(), TIPOS, "x") is None
    assert avisos and avisos[0][1] == "no está"


def test_resumir_desde_datos_con_filtro(raiz, monkeypatch):
    llamadas = []
    monkeypatch.setattr(editor_resumen, "abrir_resumen", lambda *a, **k: llamadas.append((a, k)))
    ex = _abrir(raiz, monkeypatch, _tabla())
    panel = ex.panel_datos
    assert "Resumir…" in _textos(panel)
    panel.agregar_condicion("Ciudad", "es igual a", "B")
    panel.aplicar()
    _esperar(raiz, lambda: len(panel.filas) == 3)
    panel.boton_resumir.invoke()
    _esperar(raiz, lambda: llamadas)
    (args, _kw) = llamadas[0]
    tabla = args[1]
    assert list(tabla["Ciudad"]) == ["B", "B", "B"]
    assert list(tabla.index) == [0, 1, 2]
    assert args[2]["Importe"] == "entero"
    assert args[3] == "ventas"


def test_resumir_desde_datos_sin_filtro_pasa_la_tabla(raiz, monkeypatch):
    llamadas = []
    monkeypatch.setattr(editor_resumen, "abrir_resumen", lambda *a, **k: llamadas.append((a, k)))
    df = _tabla()
    ex = _abrir(raiz, monkeypatch, df)
    ex.panel_datos.boton_resumir.invoke()
    assert llamadas[0][0][1] is df


@pytest.mark.skipif(
    importlib.util.find_spec("excel_tool.core.aggregate") is None,
    reason="core/aggregate.py todavía no existe",
)
def test_con_el_core_real(raiz):
    from excel_tool.core import aggregate as real

    editor = _editor(raiz, fixed_group=["Cliente"])
    editor.set_spec(
        real.AggregationSpec(
            [], [real.Aggregation("Cantidad", "sum"), real.Aggregation("", "count")]
        )
    )
    editor.ver_como_queda()
    _esperar(raiz, lambda: editor.resultado is not None)
    assert len(editor.resultado.df) == 1


def test_ventana_resumir_entra_en_768_con_varios_calculos_y_avisos(raiz, core):
    ventana = editor_resumen.abrir_resumen(raiz, _df(), TIPOS, "ventas")
    editor = ventana.editor
    for _ in range(5):
        fila = editor.agregar_calculo()
        fila.poner("Importe", "sum", None)  # columna "mixta": muestra el aviso de celdas
    raiz.update()
    assert ventana.ventana.winfo_reqheight() <= 768
    assert ventana.cuerpo.barra_v.winfo_ismapped()
    ventana.ventana.geometry("800x600+0+0")
    raiz.update()
    for boton in (ventana.boton_guardar, ventana.boton_cerrar):
        assert boton.winfo_viewable()
        abajo = boton.winfo_rooty() - ventana.ventana.winfo_rooty() + boton.winfo_height()
        assert abajo <= ventana.ventana.winfo_height()
    ventana.cuerpo.lienzo.yview_moveto(1.0)
    raiz.update()
    assert ventana.cuerpo.lienzo.yview()[1] == pytest.approx(1.0)
    ventana.cerrar()
