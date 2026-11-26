"""Integración de la Fase 5: el asistente "Unir archivos" de punta a punta, desde la UI (xvfb).

Se parte de la pantalla de inicio, se hace clic en los botones reales y se espera al worker. Todo
con el core real. Solo se reemplazan el selector de archivos (devuelve archivos del escenario), los
diálogos del sistema (`conftest.py`) y `open_folder`. Los esperados salen de los JSON del generador
o se recuentan acá con `csv`, `Decimal` y openpyxl (nunca con el código que se prueba).
"""

from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from excel_tool.core.io import read_table

from .asistente_util import Conductor, textos

pytestmark = pytest.mark.pantalla

JERGA = re.compile(
    r"\b(join|merge|left|inner|key|plan|group\s?by|aggregate|aggregation|distinct|split|null|"
    r"dataframe|encoding|dtype|chunk|spec|output|traceback)\b",
    re.IGNORECASE,
)


def _dec(texto) -> Decimal:
    t = str(texto).strip()
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    return Decimal(t)


def _json(fixtures_esperados: Path, nombre: str) -> dict:
    return json.loads(
        (fixtures_esperados / "escenario_3" / f"{nombre}.json").read_text(encoding="utf-8")
    )


def _hasta_paso4(c: Conductor, esc: dict, *, cobranzas: bool = False, ceros: bool = False):
    """Pasos 1 a 3 como un usuario: clientes -> ID cliente -> ventas (y cobranzas)."""
    c.abrir()
    c.elegir_principal(esc["clientes"])
    c.siguiente()
    c.marcar_claves("ID cliente")
    c.siguiente()
    c.agregar(esc["ventas"])
    if cobranzas:
        c.agregar(esc["cobranzas"])
        if ceros:
            c.limpieza("strip_leading_zeros")
    c.siguiente()
    c.esperar_paso4()


def _hasta_guardar(c: Conductor):
    """Siguiente por 5 y 6 sin tocar nada, hasta el paso 7."""
    c.esperar_tamano()
    c.siguiente()  # 5 -> 6
    c.siguiente()  # 6 -> 7


def _sin_jerga(c: Conductor) -> None:
    for t in textos(c.asistente):
        assert not JERGA.search(t), t


def _leer_xlsx(ruta: Path) -> list[tuple]:
    ws = openpyxl.load_workbook(ruta, read_only=True).worksheets[0]
    return [tuple(f) for f in ws.iter_rows(values_only=True)]


# ----------------------------------------------------------------- variante 1: detalle


def test_variante_detalle_88_filas_guardada_y_releida(c, esc, tmp_path):
    _hasta_paso4(c, esc)
    # Sin opción elegida de entrada: no deja avanzar hasta elegir.
    assert c.tarjeta("ventas.csv").var.get() == ""
    assert "ventas.csv" in c.sin_avance()
    c.modo_detalle("ventas.csv")
    assert "88 filas" in c.paso(4).etiqueta_total.cget("text")
    c.siguiente()
    c.esperar_tamano()
    c.siguiente()
    c.siguiente()
    assert c.asistente.indice == 6
    c.esperar_revision()
    assert "88 filas" in c.paso(7).etiqueta_total.cget("text")
    assert not any("detalle" in a.lower() for a in [c.paso(7).etiqueta_avisos.cget("text")])
    res = c.unir()
    assert len(res.df) == 88
    assert "El resultado tiene 88 filas" in c.paso(7).etiqueta_filas.cget("text")

    destino = tmp_path / "detalle.xlsx"
    v = c.guardar_una_tabla(destino)
    assert v._resultado.rows_total == 88 and v._resultado.files == [destino]
    assert c.asistente.boton_cancelar.cget("text") == "Terminar"
    assert "quedó guardado" in c.paso(7).etiqueta_guardado.cget("text")

    filas = _leer_xlsx(destino)
    assert filas[0][:4] == ("ID cliente", "Nombre", "Región", "Vendedor")
    cuerpo = filas[1:]
    assert len(cuerpo) == 88
    por_cliente = Counter(str(int(f[0])) for f in cuerpo)
    clientes = read_table(esc["clientes"]).df["ID cliente"].tolist()
    for cid in clientes:
        esperadas = esc["ventas_esp"]["por_cliente"].get(cid, {"filas": 1})["filas"]
        assert por_cliente[cid] == esperadas, cid
    # Los importes originales salen todos (recuento con csv, no con el producto).
    idx = filas[0].index("Importe")
    importes = Counter(str(f[idx]) for f in cuerpo if f[idx] not in (None, ""))
    propias = [v for v in esc["ventas_csv"] if v["id"] in por_cliente]
    # "Importe" tiene celdas ilegibles ('n/a'...): la columna queda como texto y sale tal cual.
    assert importes == Counter(v["importe"] for v in propias if v["importe"] != "")
    _sin_jerga(c)
    assert c.e.carteles == []


# ----------------------------------------------------------------- variante 2: resumen por cliente

CALCULOS_VENTAS = [
    ("Sumar", "Importe", "Total comprado"),
    ("Contar filas", "", "Compras"),
    ("Contar distintos", "Artículo", "Artículos distintos"),
    ("Máximo", "Fecha", "Última compra"),
]


def _tabla(ruta: Path) -> tuple[list[str], list[dict]]:
    filas = _leer_xlsx(ruta)
    cab = [str(x) for x in filas[0]]
    return cab, [dict(zip(cab, f, strict=True)) for f in filas[1:]]


def test_variante_resumen_por_cliente_30_filas_con_ceros_ADR037(c, esc, tmp_path):
    _hasta_paso4(c, esc)
    c.modo_resumen("ventas.csv", calculos=CALCULOS_VENTAS)
    assert "30 filas" in c.paso(4).etiqueta_total.cget("text")
    assert "las mismas que el archivo principal" in c.paso(4).etiqueta_total.cget("text")
    c.siguiente()
    _hasta_guardar(c)
    c.esperar_revision()
    assert "30 filas" in c.paso(7).etiqueta_total.cget("text")
    assert "mismas filas que el archivo principal (30)" in c.paso(7).etiqueta_control.cget("text")
    res = c.unir()
    assert len(res.df) == 30

    destino = tmp_path / "resumen_cliente.xlsx"
    v = c.guardar_una_tabla(destino)
    assert v._resultado.rows_total == 30
    cab, filas = _tabla(destino)
    assert cab == [
        "ID cliente", "Nombre", "Región", "Vendedor",
        "Total comprado", "Compras", "Artículos distintos", "Última compra",
    ]  # fmt: skip
    assert len(filas) == 30
    clientes = read_table(esc["clientes"]).df["ID cliente"].tolist()
    assert [str(int(f["ID cliente"])) for f in filas] == clientes  # mismo orden que el principal
    por = esc["ventas_esp"]["por_cliente"]
    con_ventas = 0
    for f in filas:
        cid = str(int(f["ID cliente"]))
        if cid in por:
            e = por[cid]
            con_ventas += 1
            assert Decimal(str(f["Total comprado"])) == Decimal(e["suma_importe"]), cid
            assert (
                f["Compras"] == e["filas"] and f["Artículos distintos"] == e["articulos_distintos"]
            )
            assert f["Última compra"].strftime("%d/%m/%Y") == e["ultima_fecha"], cid
        else:
            # ADR-037: sin compras, Sumar y Contar dan 0 (número, no celda vacía); la fecha, vacía.
            assert cid in ("5", "17", "28")
            assert f["Total comprado"] == 0 and f["Compras"] == 0 and f["Artículos distintos"] == 0
            assert f["Última compra"] in (None, "")
    assert con_ventas == 27
    _sin_jerga(c)
    assert c.e.carteles == []


# ----------------------------------------------------------------- variante 3: cliente + artículo


def test_variante_cliente_y_articulo(c, esc, tmp_path):
    _hasta_paso4(c, esc)
    c.modo_resumen(
        "ventas.csv",
        grupos=[("Artículo", None)],
        calculos=[("Sumar", "Importe", "Total"), ("Contar filas", "", "Compras")],
    )
    clientes = read_table(esc["clientes"]).df["ID cliente"].tolist()
    # Recuento independiente con csv + Decimal.
    pares: dict[tuple[str, str], list] = defaultdict(lambda: [Decimal(0), 0])
    for v in esc["ventas_csv"]:
        if v["id"] in clientes:
            p = pares[(v["id"], v["articulo"])]
            p[1] += 1
            if re.fullmatch(r"\d{1,3}(\.\d{3})*,\d{2}|\d+,\d{2}", v["importe"]):
                p[0] += _dec(v["importe"])
    sin_compras = [x for x in clientes if x not in {k[0] for k in pares}]
    esperadas = len(pares) + len(sin_compras)
    assert f"{esperadas} filas" in c.paso(4).etiqueta_total.cget("text")
    c.siguiente()
    _hasta_guardar(c)
    res = c.unir()
    assert len(res.df) == esperadas
    destino = tmp_path / "cliente_articulo.xlsx"
    c.guardar_una_tabla(destino)
    cab, filas = _tabla(destino)
    assert len(filas) == esperadas
    vistos = {}
    for f in filas:
        cid = str(int(f["ID cliente"]))
        if cid in sin_compras:
            assert f["Artículo"] in (None, "") and f["Total"] == 0 and f["Compras"] == 0
            continue
        vistos[(cid, f["Artículo"])] = (Decimal(str(f["Total"])), f["Compras"])
    assert set(vistos) == set(pares)
    for k, (total, n) in vistos.items():
        assert total == pares[k][0] and n == pares[k][1], k
    assert c.e.carteles == []


# ------------------------------------------------ variante 4: resumen final por región


def test_variante_resumen_final_por_region(c, esc, tmp_path):
    _hasta_paso4(c, esc)
    c.modo_resumen("ventas.csv", calculos=CALCULOS_VENTAS)
    c.siguiente()
    c.esperar_tamano()
    c.siguiente()  # a Resumen final
    p6 = c.paso(6)
    p6.radio_si.invoke()
    c.e.esperar(lambda: p6.editor is not None, 120)
    ed = p6.editor
    ed.checks["Región"].invoke()
    for op, col, nombre in [
        ("Sumar", "Total comprado", "Total vendido"),
        ("Contar distintos", "ID cliente", "Clientes"),
    ]:
        fila = ed.agregar_calculo()
        fila.var_operacion.set(op)
        fila.combo_operacion.event_generate("<<ComboboxSelected>>")
        fila.var_columna.set(col)
        fila.combo_columna.event_generate("<<ComboboxSelected>>")
        fila.var_nombre.set(nombre)
    c.e.raiz.update()
    p6.boton_listo.invoke()
    c.e.raiz.update()
    assert "Agrupamiento armado" in p6.etiqueta_resumen.cget("text")
    c.siguiente()
    c.esperar_revision()
    assert "5 filas" in c.paso(7).etiqueta_total.cget("text")
    assert "agrupaste el resultado" in c.paso(7).etiqueta_control.cget("text")
    res = c.unir()
    assert len(res.df) == 5
    destino = tmp_path / "por_region.xlsx"
    c.guardar_una_tabla(destino)
    cab, filas = _tabla(destino)
    assert cab == ["Región", "Total vendido", "Clientes"]
    esp = esc["region_esp"]
    assert {f["Región"]: Decimal(str(f["Total vendido"])) for f in filas} == {
        r: Decimal(v) for r, v in esp["total_importe_por_region"].items()
    }
    assert {f["Región"]: f["Clientes"] for f in filas} == esp["clientes_por_region"]
    todas = sum(
        (
            _dec(v["importe"])
            for v in esc["ventas_csv"]
            if re.fullmatch(r"\d{1,3}(\.\d{3})*,\d{2}|\d+,\d{2}", v["importe"])
        ),
        Decimal(0),
    )
    total = sum((Decimal(str(f["Total vendido"])) for f in filas), Decimal(0))
    assert total + Decimal(esp["importe_excluido_ids_fuera_de_clientes"]) == todas
    assert c.e.carteles == []
