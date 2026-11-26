"""Checklist de PLAN §6 "Unir archivos" (detalle, resumen, general) que se puede automatizar,
manejando el asistente real desde la pantalla de inicio (xvfb) con el core real."""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

import pytest

from excel_tool.core.io import read_table

from .asistente_util import Conductor, textos
from .test_fase5_asistente import (  # noqa: F401  (fixtures y ayudas compartidas)
    CALCULOS_VENTAS,
    _dec,
    _hasta_guardar,
    _hasta_paso4,
    _leer_xlsx,
    _sin_jerga,
    _tabla,
)

pytestmark = pytest.mark.pantalla

ES_IMPORTE = re.compile(r"\d{1,3}(\.\d{3})*,\d{2}|\d+,\d{2}")


def _ventas_por_cliente(esc: dict) -> dict[str, list[dict]]:
    por: dict[str, list[dict]] = defaultdict(list)
    for v in esc["ventas_csv"]:
        por[v["id"]].append(v)
    return por


# ------------------------------------------------ detalle


def test_dos_detalles_avisan_y_cambiar_a_resumen_lo_resuelve_en_un_clic(c, esc):
    c.abrir()
    c.elegir_principal(esc["clientes"])
    c.siguiente()
    c.marcar_claves("ID cliente")
    c.siguiente()
    c.agregar(esc["ventas"])
    c.agregar(esc["cobranzas"])
    # "Nro Cliente" no se empareja solo: hay que elegirlo a mano (checklist: clave con otro nombre)
    assert "cobranzas.xlsx" in c.sin_avance()
    assert c.paso(3).filas["ventas.csv"].combos_clave == {}  # la de ventas se emparejó sola
    assert "ID cliente" in c.paso(3).filas["cobranzas.xlsx"].combos_clave
    c.elegir_columna_comun("cobranzas.xlsx", "ID cliente", "Nro Cliente")
    # Sin la opción de ceros solo coinciden 5 claves y sugiere la opción
    ind = c.indicador("cobranzas.xlsx")
    assert ind.startswith("⚠️ Solo 5 de 30") and "Sacar ceros a la izquierda" in ind
    assert "27" in ind
    c.limpieza("strip_leading_zeros")
    ind = c.indicador("cobranzas.xlsx")
    assert ind.startswith("✅ 27 de 30"), ind
    c.siguiente()
    c.esperar_paso4()
    c.modo_detalle("ventas.csv")
    p4 = c.paso(4)
    assert not p4.marco_varios.winfo_manager()
    c.modo_detalle("cobranzas.xlsx")
    assert p4.marco_varios.winfo_manager()  # aparece el aviso
    assert "totales quedan inflados" in " ".join(textos(p4.marco_varios))
    ced = esc["diag_esp"]["cobranzas"]["sin_ceros_a_la_izquierda"]
    esperadas_dos = ced["filas_resultado_ventas_detalle_y_cobranzas_detalle"]
    assert f"{esperadas_dos:,}".replace(",", ".") + " filas" in p4.etiqueta_varios_filas.cget(
        "text"
    )
    assert len(p4.botones_cambiar) == 1
    assert p4.botones_cambiar[0].cget("text") == "Cambiar cobranzas.xlsx a resumen"
    p4.botones_cambiar[0].invoke()  # un clic: resumen por defecto + editor
    c.e.esperar(lambda: p4.editor is not None, 30)
    p4.boton_listo.invoke()
    c.e.raiz.update()
    c.esperar_paso4()
    assert not p4.marco_varios.winfo_manager()  # el aviso se fue
    assert c.estado.secundarios[1].mode == "summary"
    esperadas_uno = ced["filas_resultado_ventas_detalle_y_cobranzas_resumen"]
    assert f"{esperadas_uno} filas" in p4.etiqueta_total.cget("text")
    assert esperadas_uno == 88
    c.siguiente()
    _hasta_guardar(c)
    res = c.unir()
    assert len(res.df) == 88
    _sin_jerga(c)
    assert c.e.carteles == []


def test_ceros_a_la_izquierda_5_contra_27_y_totales(c, esc, tmp_path):
    c.abrir()
    c.elegir_principal(esc["clientes"])
    c.siguiente()
    c.marcar_claves("ID cliente")
    c.siguiente()
    c.agregar(esc["cobranzas"])
    c.elegir_columna_comun("cobranzas.xlsx", "ID cliente", "Nro Cliente")
    assert c.indicador("cobranzas.xlsx").startswith("⚠️ Solo 5 de 30")
    c.limpieza("strip_leading_zeros")
    assert c.indicador("cobranzas.xlsx").startswith("✅ 27 de 30")
    c.limpieza("strip_leading_zeros", False)  # volver atrás recalcula
    assert c.indicador("cobranzas.xlsx").startswith("⚠️ Solo 5 de 30")
    c.limpieza("strip_leading_zeros", True)
    c.siguiente()
    c.esperar_paso4()
    c.modo_resumen(
        "cobranzas.xlsx",
        calculos=[("Sumar", "Monto", "Total cobrado"), ("Contar filas", "", "Cobros")],
    )
    c.siguiente()
    _hasta_guardar(c)
    c.unir()
    destino = tmp_path / "cobros.xlsx"
    c.guardar_una_tabla(destino)
    cab, filas = _tabla(destino)
    esp = esc["cobranzas_esp"]["por_cliente"]
    con = 0
    for f in filas:
        cid = str(int(f["ID cliente"]))
        if cid in esp:
            con += 1
            assert Decimal(str(f["Total cobrado"])) == Decimal(esp[cid]["suma_monto"]), cid
            assert f["Cobros"] == esp[cid]["filas"]
        else:
            assert f["Total cobrado"] == 0 and f["Cobros"] == 0
    assert con == 27
    assert c.e.carteles == []


# ------------------------------------------------ resumen


def test_varios_calculos_con_nombres_editados_y_aviso_de_invalidos(c, esc, tmp_path):
    _hasta_paso4(c, esc)
    c.modo_resumen("ventas.csv", calculos=CALCULOS_VENTAS)
    nombres = [a.name for a in c.estado.secundarios[0].summary.aggregations]
    assert nombres == ["Total comprado", "Compras", "Artículos distintos", "Última compra"]
    c.siguiente()
    _hasta_guardar(c)
    res = c.unir()
    assert list(res.df.columns)[4:] == nombres
    # el aviso de la suma con inválidos, en la revisión y en el resultado
    assert "3 celdas con contenido que no se pudieron leer como números" in (
        c.paso(7).etiqueta_avisos_resultado.cget("text")
    )
    assert "1 celda vacía" in c.paso(7).etiqueta_avisos_resultado.cget("text")


def test_el_editor_avisa_de_los_importes_invalidos_al_ver_como_queda(c, esc):
    _hasta_paso4(c, esc)
    p4 = c.paso(4)
    c.tarjeta("ventas.csv").radio_resumen.invoke()
    c.e.esperar(lambda: p4.editor is not None, 30)
    ed = p4.editor
    fila = ed.agregar_calculo()
    fila.var_operacion.set("Sumar")
    fila.combo_operacion.event_generate("<<ComboboxSelected>>")
    fila.var_columna.set("Importe")
    fila.combo_columna.event_generate("<<ComboboxSelected>>")
    ed.resultado = None
    ed.boton_previa.invoke()
    c.e.esperar(lambda: ed.resultado is not None, 60)
    assert "3 celdas de Importe no son números y no se sumaron" in ed.etiqueta_estado.cget("text")
    p4.boton_listo.invoke()


def test_listar_valores_sin_repetidos(c, esc):
    _hasta_paso4(c, esc)
    c.modo_resumen(
        "ventas.csv",
        calculos=[("Listar valores", "Artículo", "Qué compró"), ("Contar filas", "", "Compras")],
    )
    c.siguiente()
    _hasta_guardar(c)
    res = c.unir()
    por = _ventas_por_cliente(esc)
    for _, f in res.df.iterrows():
        cid = f["ID cliente"]
        if cid not in por:
            assert f["Qué compró"] in ("", None)
            continue
        valores = f["Qué compró"].split(", ")
        assert len(valores) == len(set(valores)), (cid, valores)  # sin repetidos
        assert set(valores) == {v["articulo"] for v in por[cid]}, cid
    assert any(", " in x for x in res.df["Qué compró"]), "ningún cliente con 2 artículos"


def test_agrupar_por_mes_de_fecha(c, esc, tmp_path):
    _hasta_paso4(c, esc)
    c.modo_resumen(
        "ventas.csv",
        grupos=[("Fecha", "por mes")],
        calculos=[("Sumar", "Importe", "Total"), ("Contar filas", "", "Compras")],
    )
    c.siguiente()
    _hasta_guardar(c)
    res = c.unir()
    col = "Fecha (mes)"
    assert col in res.df.columns
    clientes = set(read_table(esc["clientes"]).df["ID cliente"])
    pares: dict[tuple[str, str], list] = defaultdict(lambda: [Decimal(0), 0])
    for v in esc["ventas_csv"]:
        if v["id"] in clientes and v["fecha"]:
            d, m, a = v["fecha"].split("/")
            p = pares[(v["id"], f"{a}-{m}")]
            p[1] += 1
            if ES_IMPORTE.fullmatch(v["importe"]):
                p[0] += _dec(v["importe"])
    vistos = {
        (f["ID cliente"], f[col]): (_dec(f["Total"]), int(f["Compras"]))
        for _, f in res.df.iterrows()
        if f[col]
    }
    assert vistos == {k: (v[0], v[1]) for k, v in pares.items()}
    meses = [m for m in res.df[col] if m]
    assert all(re.fullmatch(r"\d{4}-\d{2}", m) for m in meses)
    assert c.e.carteles == []


# ------------------------------------------------ hojas


def test_detalle_y_resumen_en_dos_hojas_y_hoja_sin_coincidencia(c, esc, tmp_path):
    _hasta_paso4(c, esc)
    c.modo_detalle("ventas.csv")
    c.siguiente()
    c.esperar_tamano()
    c.siguiente()
    p6 = c.paso(6)
    p6.radio_si.invoke()
    c.e.esperar(lambda: p6.editor is not None, 120)
    ed = p6.editor
    ed.checks["Región"].invoke()
    fila = ed.agregar_calculo()
    fila.var_operacion.set("Sumar")
    fila.combo_operacion.event_generate("<<ComboboxSelected>>")
    fila.var_columna.set("Importe")
    fila.combo_columna.event_generate("<<ComboboxSelected>>")
    fila.var_nombre.set("Total vendido")
    c.e.raiz.update()
    p6.boton_listo.invoke()
    c.siguiente()
    res = c.unir()
    assert len(res.df) == 5
    p7 = c.paso(7)
    assert p7.boton_libro.winfo_manager()  # "Guardar en un Excel con varias hojas"
    assert p7.boton_libro.cget("text") == "Guardar en un Excel con varias hojas"
    destino = tmp_path / "libro.xlsx"
    c.e.destino = destino
    p7.boton_libro.invoke()
    c.e.esperar(lambda: destino.exists() or bool(c.e.carteles), 120)
    c.e.esperar(lambda: "✅" in p7.etiqueta_guardado.cget("text"), 60)
    assert c.e.carteles == []
    import openpyxl

    wb = openpyxl.load_workbook(destino, read_only=True)
    assert wb.sheetnames == ["Detalle", "Resumen", "Sin coincidencia en ventas"]
    filas = {n: [tuple(r) for r in wb[n].iter_rows(values_only=True)] for n in wb.sheetnames}
    assert len(filas["Detalle"]) - 1 == 88
    assert len(filas["Resumen"]) - 1 == 5
    sin = filas["Sin coincidencia en ventas"]
    assert {str(r[0]) for r in sin[1:]} == {"98", "99"}  # ventas de IDs que no son clientes
    esperadas = [v for v in esc["ventas_csv"] if v["id"] in ("98", "99")]
    assert len(sin) - 1 == len(esperadas) == 2
    assert "Total vendido" in filas["Resumen"][0]
    assert "quedó" in p7.etiqueta_guardado.cget("text") or "Se guardó" in (
        p7.etiqueta_guardado.cget("text")
    )
    assert c.asistente.boton_cancelar.cget("text") == "Terminar"


# ------------------------------------------------ general


def test_clave_con_otro_nombre_y_emparejada_sola_con_mayusculas_y_espacios(c, esc, tmp_path):
    # ventas con el encabezado de la clave distinto en mayúsculas y con espacios alrededor
    texto = esc["ventas"].read_text(encoding="cp1252")
    nuevo = tmp_path / "ventas_mes.csv"
    nuevo.write_text(texto.replace("ID cliente", "  id CLIENTE ", 1), encoding="cp1252", newline="")
    c.abrir()
    c.elegir_principal(esc["clientes"])
    c.siguiente()
    c.marcar_claves("ID cliente")
    c.siguiente()
    c.agregar(nuevo)
    assert c.paso(3).filas["ventas_mes.csv"].combos_clave == {}  # emparejada sola
    assert c.indicador("ventas_mes.csv").startswith("✅ 27 de 30")
    assert c.estado.secundarios[0].key_pairs == [("ID cliente", "  id CLIENTE ")] or (
        c.estado.secundarios[0].key_pairs[0][0] == "ID cliente"
    )
    c.siguiente()
    c.esperar_paso4()
    c.modo_detalle("ventas_mes.csv")
    assert "88 filas" in c.paso(4).etiqueta_total.cget("text")


def test_elegir_y_reordenar_columnas_la_clave_una_sola_vez(c, esc, tmp_path):
    _hasta_paso4(c, esc, cobranzas=False)
    c.modo_resumen("ventas.csv", calculos=CALCULOS_VENTAS)
    c.siguiente()
    p5 = c.paso(5)
    c.esperar_tamano()
    # la clave del principal está marcada y no se puede quitar
    seccion = c.seccion("clientes.xlsx")
    assert "ID cliente" in seccion.fijas
    assert seccion.lista.casillas["ID cliente"].instate(["disabled"])
    # Quitar "Vendedor" y mover "Región" al final
    seccion.lista.casillas["Vendedor"].invoke()
    c.e.raiz.update()
    c.esperar_tamano()
    assert "Vendedor" not in p5.lista_final.textos()
    assert p5.lista_final.textos().count("ID cliente") == 1  # la clave aparece una sola vez
    lista = p5.lista_final
    lista.seleccionar(lista.textos().index("Región"))
    lista.botones["abajo_todo"].invoke()
    c.e.raiz.update()
    esperado = lista.textos()
    assert esperado[-1] == "Región" and "Vendedor" not in esperado
    c.siguiente()
    c.siguiente()
    c.esperar_revision()
    res = c.unir()
    assert list(res.df.columns) == esperado
    assert list(res.df.columns).count("ID cliente") == 1
    destino = tmp_path / "ordenado.xlsx"
    c.guardar_una_tabla(destino)
    assert _leer_xlsx(destino)[0] == tuple(esperado)
    # el indicador de tamaño baja al desmarcar columnas: lo prueba también el escenario grande
    assert c.e.carteles == []


def test_columnas_con_el_mismo_nombre_llevan_el_nombre_del_archivo(c, esc):
    c.abrir()
    c.elegir_principal(esc["clientes"])
    c.siguiente()
    c.marcar_claves("ID cliente")
    c.siguiente()
    c.agregar(esc["ventas"])
    c.agregar(esc["cobranzas"])
    c.elegir_columna_comun("cobranzas.xlsx", "ID cliente", "Nro Cliente")
    c.limpieza("strip_leading_zeros")
    c.siguiente()
    c.esperar_paso4()
    c.modo_primera("ventas.csv")
    c.modo_primera("cobranzas.xlsx")
    c.siguiente()
    c.esperar_tamano()
    textos = c.paso(5).lista_final.textos()
    fechas = [t for t in textos if "Fecha" in t]
    assert len(fechas) == 2 and len(set(fechas)) == 2, textos  # "Fecha" en dos archivos
    # el primero conserva el nombre; el que choca lleva el nombre del archivo (sin la extensión)
    assert sorted(fechas) == ["Fecha", "Fecha (cobranzas)"], fechas
    c.siguiente()
    c.siguiente()
    res = c.unir()
    cols = list(res.df.columns)
    assert len(cols) == len(set(cols))
    assert sorted(t for t in textos) == sorted(cols) or set(textos) == set(cols)
    assert len(res.df) == 30


def test_dos_columnas_clave_y_columna_cantidad_de_ventas(c, esc, tmp_path):
    """Principal = pares (cliente, artículo); las ventas se unen por las DOS columnas."""
    pares = sorted(
        {(v["id"], v["articulo"]) for v in esc["ventas_csv"]}, key=lambda p: (int(p[0]), p[1])
    )
    principal = tmp_path / "pares.csv"
    principal.write_text(
        "ID cliente;Artículo\n" + "".join(f"{a};{b}\n" for a, b in pares),
        encoding="cp1252",
        newline="",
    )
    c.abrir()
    c.elegir_principal(principal)
    c.siguiente()
    c.marcar_claves("ID cliente", "Artículo")
    c.siguiente()
    c.agregar(esc["ventas"])
    fila = c.paso(3).filas["ventas.csv"]
    assert fila.combos_clave == {}  # las dos se emparejaron solas
    assert [p[0] for p in c.estado.secundarios[0].key_pairs] == ["ID cliente", "Artículo"]
    assert "combinaciones de las columnas clave" in c.indicador("ventas.csv")
    assert c.indicador("ventas.csv").startswith("✅ ")
    c.siguiente()
    c.esperar_paso4()
    if c.estado.secundarios[0].mode is None:
        c.modo_detalle("ventas.csv")
    c.siguiente()
    c.esperar_tamano()
    # "Cantidad" de ventas está entre las columnas (bug F5-04) y la clave doble aparece una vez
    assert "Cantidad" in c.paso(5).lista_final.textos()
    c.siguiente()
    c.siguiente()
    c.esperar_revision()
    res = c.unir()
    # cada par aparece tantas veces como compras tiene (87 ventas en total: todas tienen pareja)
    assert len(res.df) == len(esc["ventas_csv"]) == 87
    cuenta = Counter((v["id"], v["articulo"]) for v in esc["ventas_csv"])
    assert Counter(zip(res.df["ID cliente"], res.df["Artículo"], strict=True)) == cuenta
    assert (
        list(res.df.columns).count("ID cliente") == 1
        and list(res.df.columns).count("Artículo") == 1
    )
    destino = tmp_path / "pares_unidos.xlsx"
    c.guardar_una_tabla(destino)
    cab, filas = _tabla(destino)
    assert "Cantidad" in cab and len(filas) == 87
    # recuento de cantidades con csv
    assert Counter(str(f["Cantidad"]) for f in filas if f["Cantidad"] not in (None, "")) == Counter(
        v["cantidad"] for v in esc["ventas_csv"] if v["cantidad"] != ""
    )
    assert c.e.carteles == []


def test_ocho_archivos_secundarios_con_solo_en_ambos(c, esc, tmp_path):
    clientes = read_table(esc["clientes"]).df["ID cliente"].tolist()
    extras = []
    for n in range(1, 7):
        ruta = tmp_path / f"extra_{n}.csv"
        ids = [cid for cid in clientes if int(cid) % (n + 1) != 0] if n < 6 else clientes[:20]
        ruta.write_text(
            f"ID cliente;Dato {n}\n" + "".join(f"{cid};D{n}-{cid}\n" for cid in ids),
            encoding="cp1252",
            newline="",
        )
        extras.append((n, ruta, set(ids)))
    c.abrir()
    c.elegir_principal(esc["clientes"])
    c.siguiente()
    c.marcar_claves("ID cliente")
    c.siguiente()
    for _, ruta, _ in extras:
        c.agregar(ruta)
    c.agregar(esc["ventas"])
    c.agregar(esc["cobranzas"])
    c.elegir_columna_comun("cobranzas.xlsx", "ID cliente", "Nro Cliente")
    c.limpieza("strip_leading_zeros")
    assert len(c.estado.secundarios) == 8
    assert len(c.paso(3).filas) == 8
    c.solo_en_ambos("extra_6.csv")  # solo quedan los clientes que están en este archivo
    c.siguiente()
    c.esperar_paso4()
    c.modo_resumen("ventas.csv", calculos=[("Sumar", "Importe", "Total comprado")])
    c.modo_resumen("cobranzas.xlsx", calculos=[("Sumar", "Monto", "Total cobrado")])
    assert "20 filas" in c.paso(4).etiqueta_total.cget("text")
    c.siguiente()
    c.esperar_tamano()
    c.siguiente()
    c.siguiente()
    c.esperar_revision()
    assert "20 filas" in c.paso(7).etiqueta_total.cget("text")
    assert len(c.paso(7).etiquetas_archivo) == 8
    res = c.unir()
    assert len(res.df) == 20
    assert list(res.df["ID cliente"]) == clientes[:20]  # orden del principal
    for n, _, ids in extras:
        valores = dict(zip(res.df["ID cliente"], res.df[f"Dato {n}"], strict=True))
        for cid in clientes[:20]:
            assert valores[cid] == (f"D{n}-{cid}" if cid in ids else ""), (n, cid)
    assert {"Total comprado", "Total cobrado"} <= set(res.df.columns)
    assert c.e.carteles == []
    # la ventana del asistente (con 8 archivos y 9 casillas de columnas) entra en 768 de alto
    assert c.asistente.winfo_height() <= 768


# ------------------------------------------------ receta


def _formatear_es_ar(valor: Decimal) -> str:
    entero, dec = f"{valor:.2f}".split(".")
    partes = []
    while len(entero) > 3:
        partes.insert(0, entero[-3:])
        entero = entero[:-3]
    partes.insert(0, entero)
    return ".".join(partes) + "," + dec


def _mes_siguiente(esc: dict, carpeta: Path) -> None:
    """Archivos del 'mes siguiente': mismos nombres; en las ventas todos los importes valen el
    doble (los inválidos quedan igual). Clientes y cobranzas no cambian."""
    import shutil

    carpeta.mkdir(parents=True, exist_ok=True)
    shutil.copy(esc["clientes"], carpeta / "clientes.xlsx")
    shutil.copy(esc["cobranzas"], carpeta / "cobranzas.xlsx")
    lineas = ["ID cliente;Fecha;Artículo;Cantidad;Importe"]
    for v in esc["ventas_csv"]:
        imp = v["importe"]
        if ES_IMPORTE.fullmatch(imp):
            imp = _formatear_es_ar(_dec(imp) * 2)
        lineas.append(";".join([v["id"], v["fecha"], v["articulo"], v["cantidad"], imp]))
    (carpeta / "ventas.csv").write_text("\n".join(lineas) + "\n", encoding="cp1252", newline="")


def _armar_union_de_la_receta(c: Conductor, carpeta: Path) -> None:
    c.abrir()
    c.elegir_principal(carpeta / "clientes.xlsx")
    c.siguiente()
    c.marcar_claves("ID cliente")
    c.siguiente()
    c.agregar(carpeta / "ventas.csv")
    c.agregar(carpeta / "cobranzas.xlsx")
    c.elegir_columna_comun("cobranzas.xlsx", "ID cliente", "Nro Cliente")
    c.limpieza("strip_leading_zeros")
    c.siguiente()
    c.esperar_paso4()
    c.modo_resumen("ventas.csv", calculos=CALCULOS_VENTAS)
    c.modo_resumen(
        "cobranzas.xlsx",
        calculos=[("Sumar", "Monto", "Total cobrado"), ("Contar filas", "", "Cobros")],
    )
    c.siguiente()
    c.esperar_tamano()
    c.seccion("clientes.xlsx").lista.casillas["Vendedor"].invoke()  # sin Vendedor
    c.e.raiz.update()
    c.esperar_tamano()
    lista = c.paso(5).lista_final
    lista.seleccionar(lista.textos().index("Región"))
    lista.botones["abajo_todo"].invoke()  # Región al final
    c.e.raiz.update()
    c.siguiente()
    c.siguiente()
    c.esperar_revision()


def _abrir_receta(c: Conductor, monkeypatch, receta: Path) -> None:
    from tkinter import filedialog

    monkeypatch.setattr(filedialog, "askopenfilename", lambda *a, **k: str(receta))
    boton = next(b for b in c.inicio.botones if b.cget("text") == "Abrir una receta…")
    c.inicio.asistente = None
    boton.invoke()
    c.e.esperar(lambda: c.inicio.asistente is not None or bool(c.e.carteles), 120)
    c.asistente = c.inicio.asistente


def test_receta_guardada_abierta_con_archivos_del_mes_siguiente(c, esc, tmp_path, monkeypatch):
    import json
    import shutil

    mes1 = tmp_path / "mes1"
    mes1.mkdir()
    for k in ("clientes", "ventas", "cobranzas"):
        shutil.copy(esc[k], mes1 / esc[k].name)
    _armar_union_de_la_receta(c, mes1)
    esperado_cols = c.paso(5).lista_final.textos()
    res1 = c.unir()
    assert len(res1.df) == 30

    def partes_de_12(v):
        v.rb_formato["csv_excel"].invoke()
        c.e.esperar(lambda: v._check is not None, 60)
        v.v_modo.set("rows")
        v.v_filas.set("12")
        v._al_elegir_modo()
        v._al_editar_filas()

    v1 = c.guardar_una_tabla(mes1 / "salida.csv", configurar=partes_de_12)
    assert v1._resultado.rows_total == 30 and len({p.path for p in v1._resultado.parts}) == 3
    receta = mes1 / "receta.json"
    c.e.destino = receta
    c.paso(7).boton_receta.invoke()
    c.e.raiz.update()
    assert receta.is_file() and "quedó guardada" in c.paso(7).etiqueta_guardado.cget("text")
    datos = json.loads(receta.read_text(encoding="utf-8"))
    texto = json.dumps(datos, ensure_ascii=False)
    assert str(tmp_path) not in texto and "\\" not in texto  # sin rutas absolutas del equipo
    assert datos.get("output") or "output" in texto  # recuerda cómo se guardó

    # --- mes siguiente: otra carpeta, la misma receta, archivos con los mismos nombres
    c.asistente.cerrar()
    mes2 = tmp_path / "mes2"
    _mes_siguiente(esc, mes2)
    shutil.copy(receta, mes2 / "receta.json")
    _abrir_receta(c, monkeypatch, mes2 / "receta.json")
    assert c.asistente is not None and c.asistente.indice == 2  # arranca en "Archivos a sumar"
    assert c.estado.faltantes == [] and not any("Faltan" in t for _, t, _ in c.e.carteles)
    assert [s.nombre for s in c.estado.secundarios] == ["ventas.csv", "cobranzas.xlsx"]
    c.esperar_comparacion()
    c.siguiente()
    c.esperar_paso4()
    assert [s.mode for s in c.estado.secundarios] == ["summary", "summary"]
    c.siguiente()
    c.esperar_tamano()
    assert c.paso(5).lista_final.textos() == esperado_cols  # mismas columnas y mismo orden
    c.siguiente()
    c.siguiente()
    c.esperar_revision()
    res2 = c.unir()
    assert len(res2.df) == 30 and list(res2.df.columns) == esperado_cols
    por = esc["ventas_esp"]["por_cliente"]
    for _, f in res2.df.iterrows():
        cid = f["ID cliente"]
        if cid in por:
            assert _dec(f["Total comprado"]) == Decimal(por[cid]["suma_importe"]) * 2, cid
            assert int(f["Compras"]) == por[cid]["filas"]
        else:
            assert f["Total comprado"] == "0"
    # el guardado ya viene preparado como la vez anterior (CSV en partes de 12)
    c.paso(7).boton_guardar.invoke()
    c.e.esperar(lambda: len(c.guardados) == 2, 30)
    v = c.guardados[-1]
    assert v.v_formato.get() == "csv_excel" and v.v_modo.get() == "rows"
    assert v.v_filas.get() == "12"
    c.e.destino = mes2 / "salida.csv"
    c.e.esperar(lambda: v.puede_guardar(), 60)
    v._guardar()
    c.e.esperar(lambda: v._resultado is not None or bool(c.e.carteles), 120)
    assert c.e.carteles == []
    assert v._resultado.rows_total == 30 and len({p.path for p in v._resultado.parts}) == 3
    # los archivos guardados tienen el contenido nuevo (x2)
    total = Decimal(0)
    for archivo in sorted({p.path for p in v._resultado.parts}):
        t = read_table(archivo).df
        total += sum((_dec(x) for x in t["Total comprado"] if x not in ("", "0")), Decimal(0))
    de_clientes = sum(
        (Decimal(v) for v in esc["region_esp"]["total_importe_por_region"].values()), Decimal(0)
    )
    assert total == de_clientes * 2  # los IDs 98 y 99 no son clientes: no entran


def test_receta_con_archivos_que_faltan_se_completan_en_el_asistente(c, esc, tmp_path, monkeypatch):
    import shutil

    mes1 = tmp_path / "mes1"
    mes1.mkdir()
    for k in ("clientes", "ventas", "cobranzas"):
        shutil.copy(esc[k], mes1 / esc[k].name)
    _armar_union_de_la_receta(c, mes1)
    c.e.destino = mes1 / "receta.json"
    c.paso(7).boton_receta.invoke()
    c.asistente.cerrar()
    solo = tmp_path / "solo_receta"
    solo.mkdir()
    shutil.copy(mes1 / "receta.json", solo / "receta.json")
    mes2 = tmp_path / "mes2"
    _mes_siguiente(esc, mes2)
    _abrir_receta(c, monkeypatch, solo / "receta.json")
    # avisa cuáles faltan y abre el asistente en "Archivos a sumar"
    assert [t for _, t, _ in c.e.carteles] == ["Faltan archivos"]
    cuerpo = c.e.carteles[0][2]
    assert all(n in cuerpo for n in ("clientes.xlsx", "ventas.csv", "cobranzas.xlsx"))
    c.e.carteles.clear()
    c.e.esperar(lambda: c.inicio.asistente is not None, 30)
    c.asistente = c.inicio.asistente
    assert set(c.estado.faltantes) == {"clientes.xlsx", "ventas.csv", "cobranzas.xlsx"}
    assert "Falta" in c.sin_avance()
    p3 = c.paso(3)
    for nombre, ruta in (
        ("clientes.xlsx", mes2 / "clientes.xlsx"),
        ("ventas.csv", mes2 / "ventas.csv"),
        ("cobranzas.xlsx", mes2 / "cobranzas.xlsx"),
    ):
        c.cola.append(ruta)
        fila = p3.fila_principal if nombre == "clientes.xlsx" else p3.filas[nombre]
        fila.boton_elegir.invoke()
        c.e.esperar(lambda n=nombre: n not in c.estado.faltantes, 60)
    c.esperar_comparacion()
    assert c.estado.faltantes == []
    c.siguiente()
    c.esperar_paso4()
    c.siguiente()
    c.esperar_tamano()
    c.siguiente()
    c.siguiente()
    c.esperar_revision()
    res = c.unir()
    assert len(res.df) == 30
    por = esc["ventas_esp"]["por_cliente"]
    for _, f in res.df.iterrows():
        if f["ID cliente"] in por:
            assert _dec(f["Total comprado"]) == Decimal(por[f["ID cliente"]]["suma_importe"]) * 2
    assert c.e.carteles == []


# ------------------------------------------------ columna "Cantidad" (bug F5-04)


def test_columna_cantidad_de_ventas_en_guardar_y_en_valores_distintos(c, esc, tmp_path):
    _hasta_paso4(c, esc)
    c.modo_detalle("ventas.csv")
    c.siguiente()
    _hasta_guardar(c)
    res = c.unir()
    assert "Cantidad" in res.df.columns and len(res.df) == 88
    clientes = set(read_table(esc["clientes"]).df["ID cliente"])
    esperadas = Counter(v["cantidad"] for v in esc["ventas_csv"] if v["id"] in clientes)
    esperadas[""] = 3  # los clientes sin compras (5, 17 y 28) no tienen cantidad
    assert sum(esperadas.values()) == 88

    # --- Valores distintos del resultado
    p7 = c.paso(7)
    p7.ver_informacion()
    ex = p7.explorador
    c.e.esperar(lambda: ex.perfil is not None and not ex.analizando, 60)
    ex.seleccionar_columna("Cantidad")
    ex._ver_distintos()
    pd_ = ex.panel_distintos
    c.e.esperar(lambda: pd_.resultado is not None, 30)
    assert pd_.resultado.count_column == "Cantidad de filas"
    assert pd_.resultado.percent_column == "Porcentaje"
    titulos = [pd_.tabla.heading(i, "text") for i in pd_.tabla["columns"]]
    assert titulos == ["Cantidad", "Cantidad de filas", "Porcentaje"]
    filas = [pd_.tabla.item(i, "values") for i in pd_.tabla.get_children()]
    vistas = {f[0]: int(f[1]) for f in filas if f[0] not in ("", "(vacío)")}
    assert vistas == {k: n for k, n in esperadas.items() if k != ""}
    assert "21 valores distintos en 88 filas" in pd_.etiqueta_resumen.cget("text")
    # ... y combinado con otra columna
    ex.seleccionar_columna("Región")
    pd_.elegir_columnas(["Cantidad", "Región"])
    pd_.calcular()
    c.e.esperar(lambda: pd_.resultado is not None and len(pd_.resultado.columns) == 2, 30)
    assert pd_.resultado.count_column == "Cantidad de filas"
    assert sum(int(pd_.tabla.item(i, "values")[2]) for i in pd_.tabla.get_children()) == 88

    # --- Guardar: sugerencias de columnas para dividir con "Cantidad", y dividir por ella
    def por_cantidad(v):
        v.rb_formato["csv_excel"].invoke()
        c.e.esperar(lambda: v._check is not None and bool(v._sug), 60)
        texto = next(t for t, col in v._sug_por_texto.items() if col == "Cantidad")
        assert "21 archivos" in texto or "22 archivos" in texto
        v.v_modo.set("value")
        v.v_col_valor.set(texto)
        v.v_subdividir.set(False)
        v._al_elegir_modo()
        c.e.esperar(lambda: v._partes is not None, 60)

    v = c.guardar_una_tabla(tmp_path / "por_cantidad.csv", configurar=por_cantidad)
    r = v._resultado
    assert r.rows_total == 88 == sum(p.rows for p in r.parts)
    por_valor = Counter()
    for p in r.parts:
        por_valor[p.value[0]] += p.rows
    assert {k: n for k, n in por_valor.items() if k not in ("", "(vacío)")} == {
        k: n for k, n in esperadas.items() if k != ""
    }
    assert c.e.carteles == []


# ------------------------------------------------ ventanas en 768 px


def _entra(ventana, nombre: str) -> None:
    ventana.update_idletasks()
    assert ventana.winfo_reqheight() <= 768 or ventana.winfo_height() <= 768, (
        nombre,
        ventana.winfo_reqheight(),
        ventana.winfo_height(),
    )
    assert ventana.winfo_height() <= 768, (nombre, ventana.winfo_height())


def test_todas_las_ventanas_nuevas_entran_en_768(c, esc, tmp_path):
    c.abrir()
    c.elegir_principal(esc["clientes"])
    c.siguiente()
    c.marcar_claves("ID cliente")
    c.siguiente()
    c.agregar(esc["ventas"])
    c.agregar(esc["cobranzas"])
    c.elegir_columna_comun("cobranzas.xlsx", "ID cliente", "Nro Cliente")
    _entra(c.asistente, "asistente paso 3")
    c.siguiente()
    c.esperar_paso4()
    _entra(c.asistente, "asistente paso 4")
    p4 = c.paso(4)
    c.modo_detalle("ventas.csv")
    c.modo_detalle("cobranzas.xlsx")  # con el aviso de dos detalles visible
    c.e.raiz.update()
    _entra(c.asistente, "asistente paso 4 con aviso")
    c.verificar_entra_en_768()
    # editor de resumen del paso 4 (con todas las operaciones y varios cálculos)
    c.tarjeta("cobranzas.xlsx").radio_resumen.invoke()
    c.e.esperar(lambda: p4.editor is not None, 30)
    for op, col in (
        ("Sumar", "Monto"),
        ("Contar filas", ""),
        ("Máximo", "Fecha"),
        ("Listar valores", "Medio de pago"),
    ):
        fila = p4.editor.agregar_calculo()
        fila.var_operacion.set(op)
        fila.combo_operacion.event_generate("<<ComboboxSelected>>")
        if col:
            fila.var_columna.set(col)
            fila.combo_columna.event_generate("<<ComboboxSelected>>")
    c.e.raiz.update()
    _entra(p4.dialogo, "editor de resumen (paso 4)")
    bot = p4.boton_listo
    assert (
        bot.winfo_rooty() - p4.dialogo.winfo_rooty() + bot.winfo_height()
        <= p4.dialogo.winfo_height()
    )
    p4.boton_listo.invoke()
    c.modo_detalle("ventas.csv")
    c.siguiente()
    c.esperar_tamano()
    _entra(c.asistente, "asistente paso 5")
    c.siguiente()
    p6 = c.paso(6)
    p6.radio_si.invoke()
    c.e.esperar(lambda: p6.editor is not None, 120)
    c.e.raiz.update()
    _entra(p6.dialogo, "editor del resumen final (paso 6)")
    bot = p6.boton_listo
    assert (
        bot.winfo_rooty() - p6.dialogo.winfo_rooty() + bot.winfo_height()
        <= p6.dialogo.winfo_height()
    )
    p6.editor.checks["Región"].invoke()
    p6.boton_listo.invoke()
    c.siguiente()
    c.esperar_revision()
    _entra(c.asistente, "asistente paso 7")
    c.unir()
    _entra(c.asistente, "asistente paso 7 con resultado")
    p7 = c.paso(7)
    # el botón de guardar y el de la receta se alcanzan (el cuerpo tiene scroll)
    assert p7.boton_guardar.winfo_ismapped() and p7.boton_receta.winfo_ismapped()
    p7.ver_informacion()
    c.e.esperar(lambda: p7.explorador.perfil is not None, 60)
    _entra(p7.explorador, "explorador del resultado")
    c.guardar_una_tabla(tmp_path / "x.xlsx")
    _entra(c.guardados[-1].win, "ventana Guardar desde el asistente")
    assert c.e.carteles == []
