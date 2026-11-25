"""Integración de la Fase 3: explorador + paso Guardar en pantalla (xvfb) con el core REAL.

Nada de perfiles ni de `save_result` falsos. Solo se reemplazan los diálogos del sistema (en
`tests/integracion/conftest.py`) y `platform_utils.open_folder`. Los esperados salen de recontar los
CSV con el módulo `csv` de la biblioteca estándar.
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import tkinter as tk
from collections import Counter
from pathlib import Path

import openpyxl
import pytest

from excel_tool import platform_utils
from excel_tool.core.io import CsvOptions, read_table
from excel_tool.core.output import NOMBRE_INDICE
from excel_tool.core.preferences import load_preferences
from excel_tool.ui import guardar
from excel_tool.ui.explorador import abrir_explorador
from excel_tool.ui.selector import Seleccion

pytestmark = pytest.mark.pantalla

CSV_AR = CsvOptions(";", "utf-8-sig")
ALTO_PANTALLA = 768
JERGA = (
    "join",
    "merge",
    "group by",
    "distinct",
    "split",
    "null",
    "dataframe",
    "encoding",
    "dtype",
    "chunk",
    "spec",
    "output",
    "traceback",
)


def _seleccion(path: Path) -> Seleccion:
    return Seleccion(path, None, CSV_AR, 0)


def _textos(widget) -> list[str]:
    hallados = []
    for hijo in widget.winfo_children():
        try:
            hallados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        hallados.extend(_textos(hijo))
    return hallados


def _abrir(entorno, path: Path, espera: float = 120.0):
    ex = abrir_explorador(entorno.raiz, _seleccion(path))
    entorno.esperar(lambda: ex.perfil is not None, espera)
    entorno.esperar(lambda: not ex.analizando, espera)
    return ex


@pytest.fixture
def capturas(monkeypatch):
    """Guarda las ventanas de `abrir_guardar` y las llamadas a `open_folder`."""
    ventanas: list = []
    carpetas: list[Path] = []
    original = guardar.abrir_guardar

    def espia(*args, **kwargs):
        v = original(*args, **kwargs)
        ventanas.append(v)
        return v

    monkeypatch.setattr(guardar, "abrir_guardar", espia)
    monkeypatch.setattr(platform_utils, "open_folder", lambda p: carpetas.append(Path(p)))
    return ventanas, carpetas


def _borde_inferior(v, widget) -> int:
    """Borde de abajo del widget respecto del borde de arriba de la ventana (no de la pantalla:
    el administrador de ventanas la puede poner donde quiera)."""
    return widget.winfo_rooty() - v.win.winfo_rooty() + widget.winfo_height()


def _filtrar_importe_mayor(entorno, ex, minimo: str = "5000"):
    datos = ex.panel_datos
    ex.pestanas.select(ex.pestana_datos)
    datos.agregar_condicion("Importe", "es mayor que", minimo)
    datos.aplicar()
    entorno.esperar(lambda: len(datos.filas) < 200)
    entorno.esperar(lambda: not datos.boton_aplicar.instate(["disabled"]))
    return datos


def _esperados_importe_mayor(path: Path, minimo: float) -> list[list[str]]:
    """Filas del CSV con Importe > minimo, con el módulo csv (independiente del producto)."""
    with path.open(encoding="utf-8-sig", newline="") as f:
        lector = csv.reader(f, delimiter=";")
        next(lector)
        return [fila for fila in lector if fila[1] and float(fila[1].replace(",", ".")) > minimo]


def _listo_para_guardar(entorno, v, segundos: float = 60.0):
    entorno.esperar(lambda: v.puede_guardar(), segundos)


def _guardar_y_esperar(entorno, v, destino: Path, segundos: float = 240.0):
    entorno.destino = destino
    _listo_para_guardar(entorno, v)
    v._guardar()
    entorno.esperar(lambda: v._resultado is not None or entorno.carteles, segundos)
    assert entorno.carteles == [], entorno.carteles
    return v._resultado


# ------------------------------------------------------------ exportar lo filtrado


def test_exportar_lo_filtrado_a_excel_con_tipos_reales(
    entorno_ui, fixtures_generados, tmp_path, capturas
):
    ventanas, _ = capturas
    csv_path = fixtures_generados / "explorar" / "para_explorar.csv"
    esperadas = _esperados_importe_mayor(csv_path, 5000)
    assert len(esperadas) == 109  # el mismo número que la Fase 2

    ex = _abrir(entorno_ui, csv_path)
    datos = _filtrar_importe_mayor(entorno_ui, ex)
    assert len(datos.filas) == 109
    assert datos.boton_exportar.cget("text") == "Exportar lo filtrado…"
    datos.boton_exportar.invoke()
    entorno_ui.esperar(lambda: len(ventanas) == 1)
    v = ventanas[0]
    assert v is not None and v.v_destino.get() == "excel" and v.v_formato.get() == "xlsx"
    assert v.nombre == "para_explorar_filtrado"
    assert len(v.df) == 109

    destino = tmp_path / "filtrado.xlsx"
    res = _guardar_y_esperar(entorno_ui, v, destino)
    assert res.files == [destino] and res.rows_total == 109
    assert v.lbl_resultado.cget("text") == (
        "Se guardó 1 archivo con 109 filas en total (coincide con el resultado ✅)"
    )

    releido = read_table(destino).df
    assert releido["ID"].tolist() == [f[0] for f in esperadas]
    assert releido["Región"].tolist() == [f[12] for f in esperadas]
    # tipos reales (con openpyxl): número, fecha y texto con ceros
    ws = openpyxl.load_workbook(destino).worksheets[0]
    cols = [c.value for c in ws[1]]
    assert ws.max_row == 110
    assert ws.freeze_panes == "A2" and ws.auto_filter.ref
    for i, fila in enumerate(esperadas, start=2):
        celda = {c: ws.cell(row=i, column=n + 1) for n, c in enumerate(cols)}
        if fila[0]:
            assert isinstance(celda["ID"].value, (int, float)) and celda["ID"].value == int(fila[0])
        else:
            assert celda["ID"].value is None
        assert celda["Importe"].value == pytest.approx(float(fila[1].replace(",", ".")))
        if fila[2]:
            d, m, a = (int(x) for x in fila[2].split("/"))
            assert celda["Fecha"].value == dt.datetime(a, m, d)
        if fila[6]:
            assert celda["Código"].value == fila[6] and celda["Código"].data_type == "s"
    assert (
        load_preferences().last_destination == "excel"
        or load_preferences().last_destination is None
    )


def test_exportar_sin_filtrar_dice_guardar_estas_filas(
    entorno_ui, fixtures_generados, capturas, tmp_path
):
    ventanas, _ = capturas
    ex = _abrir(entorno_ui, fixtures_generados / "explorar" / "para_explorar.csv")
    ex.pestanas.select(ex.pestana_datos)
    datos = ex.panel_datos
    assert datos.boton_exportar.cget("text") == "Guardar estas filas…"
    datos.boton_exportar.invoke()
    entorno_ui.esperar(lambda: len(ventanas) == 1)
    v = ventanas[0]
    res = _guardar_y_esperar(entorno_ui, v, tmp_path / "todo.xlsx")
    assert res.rows_total == 200
    assert len(read_table(tmp_path / "todo.xlsx").df) == 200


def test_cancelar_el_dialogo_de_guardar_no_escribe(
    entorno_ui, fixtures_generados, capturas, tmp_path, monkeypatch
):
    from tkinter import filedialog

    ventanas, _ = capturas
    ex = _abrir(entorno_ui, fixtures_generados / "explorar" / "para_explorar.csv")
    ex.panel_datos.boton_exportar.invoke()
    entorno_ui.esperar(lambda: len(ventanas) == 1)
    v = ventanas[0]
    monkeypatch.setattr(filedialog, "asksaveasfilename", lambda **k: "")
    _listo_para_guardar(entorno_ui, v)
    v._guardar()
    entorno_ui.bombear(0.3)
    assert v._resultado is None
    assert [p.name for p in tmp_path.iterdir()] == ["datos"]
    assert entorno_ui.carteles == []


# ------------------------------------------------------------ recordar el destino


def test_recuerda_el_destino_de_la_ultima_vez(entorno_ui, fixtures_generados, capturas, tmp_path):
    ventanas, _ = capturas
    ex = _abrir(entorno_ui, fixtures_generados / "explorar" / "para_explorar.csv")
    ex.panel_datos.boton_exportar.invoke()
    entorno_ui.esperar(lambda: len(ventanas) == 1)
    v1 = ventanas[0]
    assert v1.v_destino.get() == "excel"  # la primera vez, Excel (ADR-032)
    v1.rb_destino["google_sheets"].invoke()
    assert v1.v_destino.get() == "google_sheets" and v1.v_formato.get() == "csv_google"
    assert load_preferences().last_destination == "google_sheets"
    assert "hojas" not in v1.rb_hojas.state() or v1.rb_hojas.state()  # no se ofrece (ver abajo)
    assert not v1.rb_hojas.winfo_ismapped()
    res = _guardar_y_esperar(entorno_ui, v1, tmp_path / "para_gs.csv")
    assert res.files == [tmp_path / "para_gs.csv"]
    primera_linea = (tmp_path / "para_gs.csv").read_text(encoding="utf-8").splitlines()[0]
    assert primera_linea.startswith("ID,Importe,Fecha")  # `,` y utf-8 para Google Sheets
    assert not (tmp_path / "para_gs.csv").read_bytes().startswith(b"\xef\xbb\xbf")
    v1.cerrar()

    ex.panel_datos.boton_exportar.invoke()
    entorno_ui.esperar(lambda: len(ventanas) == 2)
    v2 = ventanas[1]
    assert v2.v_destino.get() == "google_sheets"
    assert v2.v_formato.get() == "csv_google"
    # y lo guardado se conserva entre sesiones: lo lee de las preferencias
    assert load_preferences().last_destination == "google_sheets"
    assert load_preferences().last_folder == tmp_path


# ------------------------------------------------------------ límites y semáforo en la UI


def _poner_limites(entorno, contenido: dict) -> None:
    (platform_utils.user_data_dir() / "limites.json").write_text(
        json.dumps(contenido), encoding="utf-8"
    )


def test_limites_json_cambia_el_semaforo_y_la_ventana_entra_en_768(
    entorno_ui, fixtures_generados, capturas, tmp_path
):
    """Con un límite de Excel de 101 filas, 200 filas es rojo: no deja guardar en un archivo."""
    _poner_limites(entorno_ui, {"excel": {"max_rows": 101, "comfortable_rows": 50}})
    ventanas, _ = capturas
    ex = _abrir(entorno_ui, fixtures_generados / "explorar" / "para_explorar.csv")
    # semáforo de General con el límite chico
    assert ex.semaforo[0][0] == "Excel" and ex.semaforo[0][1].level == "obligatorio"
    assert ex.boton_dividir is not None
    ex.boton_dividir.invoke()
    entorno_ui.esperar(lambda: len(ventanas) == 1)
    v = ventanas[0]
    entorno_ui.esperar(lambda: v._check is not None)
    assert v._check.level == "obligatorio"
    assert v.lbl_nivel.cget("text").endswith("NO ENTRA EN UN SOLO ARCHIVO")
    assert v.lbl_nivel.cget("bg") == "#c62828"
    assert v.rb_uno.instate(["disabled"])
    assert v.v_modo.get() in ("rows", "keep")
    # rojo: el modo "un solo archivo" no se puede guardar, pero partes sí
    _listo_para_guardar(entorno_ui, v)
    res = _guardar_y_esperar(entorno_ui, v, tmp_path / "chico.xlsx")
    assert res.rows_total == 200 and len(res.parts) >= 2
    assert max(p.rows for p in res.parts) <= 100
    entorno_ui.bombear(0.2)
    # la ventana (con el resultado a la vista) entra en una pantalla de 768 px
    v.win.update_idletasks()
    assert v.win.winfo_reqheight() <= ALTO_PANTALLA
    assert _borde_inferior(v, v.btn_guardar) <= ALTO_PANTALLA
    assert _borde_inferior(v, v.btn_carpeta) <= ALTO_PANTALLA


def test_la_ventana_guardar_entra_en_768_en_todos_los_estados(
    entorno_ui, fixtures_generados, capturas
):
    ventanas, _ = capturas
    ex = _abrir(entorno_ui, fixtures_generados / "explorar" / "para_explorar.csv")
    ex.panel_datos.boton_exportar.invoke()
    entorno_ui.esperar(lambda: len(ventanas) == 1)
    v = ventanas[0]
    entorno_ui.esperar(lambda: v._check is not None)
    for destino in ("excel", "google_sheets", "solo_guardar", "excel"):
        v.rb_destino[destino].invoke()
        entorno_ui.esperar(lambda: v._check is not None)
        for modo in ("none", "rows", "keep", "value", "sheets"):
            if modo == "sheets" and not v.hojas_disponibles():
                continue
            v.v_modo.set(modo)
            v._al_elegir_modo()
            entorno_ui.bombear(0.15)
            v.win.update_idletasks()
            assert v.win.winfo_reqheight() <= ALTO_PANTALLA, (
                destino,
                modo,
                v.win.winfo_reqheight(),
            )
            assert _borde_inferior(v, v.btn_guardar) <= ALTO_PANTALLA
    # con la ventana mucho más baja, el cuerpo se desplaza y el pie sigue a la vista
    assert v.lienzo.winfo_height() <= guardar.ALTO_MAXIMO_CUERPO
    assert v.barra_v.winfo_ismapped() or v.marco.winfo_reqheight() <= guardar.ALTO_MAXIMO_CUERPO
    # textos visibles sin jerga
    for t in _textos(v.win):
        for palabra in JERGA:
            assert palabra not in t.lower(), (palabra, t)


# ------------------------------------------------------------ archivos grandes: semáforo


def test_semaforo_de_400k_es_amarillo_excel_y_verde_google(entorno_ui, fixtures_grandes, capturas):
    ex = _abrir(entorno_ui, fixtures_grandes / "fase3_400k.csv")
    niveles = {nombre: chk.level for nombre, chk in ex.semaforo}
    assert niveles == {"Excel": "recomendado", "Google Sheets": "ok"}
    textos = [m.cget("text") for m, _ in ex.lineas_semaforo]
    assert textos == ["●  Excel: CONVIENE DIVIDIR", "●  Google Sheets: TODO BIEN"]
    colores = [m.cget("bg") for m, _ in ex.lineas_semaforo]
    assert colores == ["#f9a825", "#2e7d32"]
    assert ex.boton_dividir is None  # amarillo no ofrece el botón de "no entra"
    assert entorno_ui.carteles == []


def test_semaforo_de_1600k_es_rojo_excel_y_amarillo_google(entorno_ui, fixtures_grandes, capturas):
    ex = _abrir(entorno_ui, fixtures_grandes / "fase3_dividir_1600k.csv")
    niveles = {nombre: chk.level for nombre, chk in ex.semaforo}
    assert niveles == {"Excel": "obligatorio", "Google Sheets": "recomendado"}
    assert ex.lineas_semaforo[0][0].cget("text") == "●  Excel: NO ENTRA EN UN SOLO ARCHIVO"
    assert ex.lineas_semaforo[0][0].cget("bg") == "#c62828"
    assert ex.boton_dividir is not None and ex.boton_dividir.cget("text") == "Dividir y guardar…"
    assert "1.048.576" in ex.lineas_semaforo[0][1].cget(
        "text"
    ) or "1.600.000" in ex.lineas_semaforo[0][1].cget("text")


def test_dividir_y_guardar_por_valor_de_region_con_el_archivo_de_1600k(
    entorno_ui, fixtures_grandes, capturas, tmp_path
):
    ventanas, carpetas = capturas
    origen = fixtures_grandes / "fase3_dividir_1600k.csv"
    ex = _abrir(entorno_ui, origen)
    ex.boton_dividir.invoke()
    entorno_ui.esperar(lambda: len(ventanas) == 1)
    v = ventanas[0]
    entorno_ui.esperar(lambda: v._check is not None and v._sug, 120)
    assert v._check.level == "obligatorio"
    assert v.rb_uno.instate(["disabled"])  # no se puede guardar todo junto
    assert not v.btn_guardar.instate(["!disabled"]) or v.v_modo.get() != "none"

    # elegir "un archivo por cada valor" de Región, sin subdividir: Norte (1,2M) entra en Excel
    texto_region = next(t for t, col in v._sug_por_texto.items() if col == "Región")
    v.v_modo.set("value")
    v.v_col_valor.set(texto_region)
    v._al_elegir_modo()
    # CSV para Excel (escribir 1,6M filas en .xlsx tarda minutos y lo cubre la medición)
    v.rb_formato["csv_excel"].invoke()
    entorno_ui.esperar(lambda: v._check is not None)
    v.v_modo.set("value")
    v.v_col_valor.set(texto_region)
    v._al_elegir_modo()
    entorno_ui.esperar(lambda: v._partes is not None)
    cuentas = Counter()
    with origen.open(encoding="utf-8-sig", newline="") as f:
        lector = csv.reader(f, delimiter=";")
        next(lector)
        for fila in lector:
            cuentas[fila[1]] += 1
    # vista previa: una fila por parte, con los nombres y las filas del core
    assert sum(p.rows for p in v._partes) == 1_600_000
    assert len(v.tabla.get_children()) == len(v._partes)
    assert v.lbl_vista.cget("text").startswith(f"Se van a generar {len(v._partes)} archivos")

    res = _guardar_y_esperar(entorno_ui, v, tmp_path / "ventas.csv", 300)
    assert res.folder == tmp_path / "ventas"
    assert res.rows_total == 1_600_000 == sum(p.rows for p in res.parts)
    por_region = Counter()
    for p in res.parts:
        por_region[p.value[0]] += p.rows
        assert p.rows <= 1_048_575
    assert por_region == cuentas
    archivos = sorted(x.name for x in res.folder.iterdir())
    assert NOMBRE_INDICE in archivos and len(archivos) == len(res.parts) + 1
    idx = read_table(res.folder / NOMBRE_INDICE).df
    assert sum(int(x) for x in idx["Filas"]) == 1_600_000
    n = len({p.path for p in res.parts})
    assert v.lbl_resultado.cget("text") == (
        f"Se guardaron {n} archivos con 1.600.000 filas en total (coincide con el resultado ✅)"
    )
    # "Abrir carpeta" llama a open_folder con la carpeta de las partes
    v.btn_carpeta.invoke()
    assert carpetas == [tmp_path / "ventas"]
    assert entorno_ui.carteles == []
    # el cartel de resultado y los botones siguen dentro de la pantalla de 768 px
    v.win.update_idletasks()
    assert v.win.winfo_reqheight() <= ALTO_PANTALLA
