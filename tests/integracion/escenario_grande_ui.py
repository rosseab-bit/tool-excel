"""Escenario grande desde la UI (no es un test): 200.000 clientes x 2.500.000 ventas en detalle.

`recorrer(c, grandes, salida)` maneja el asistente real (con el `Conductor`) y verifica por el
camino. Lo usan el test `test_fase5_grande.py` y `medicion_fase5.py` (pico de memoria con
`/usr/bin/time -v`). Los esperados salen de `tests/fixtures/esperados/grandes.json` y de aritmética
independiente (no del código que se prueba).
"""

from __future__ import annotations

import csv
import json
import math
import time
from pathlib import Path

from excel_tool.core.io import CsvOptions

from .asistente_util import Conductor

CSV_GRANDE = CsvOptions(";", "utf-8-sig")
ESTIMADAS = 2_519_998  # grandes.json: detalle_left.total
COLUMNAS = 8  # clientes (4) + ventas (4, sin la clave repetida)
MAX_FILAS_EXCEL = 1_048_576
MAX_CELDAS_SHEETS = 20_000_000


def _contar_filas(ruta: Path) -> int:
    with ruta.open(encoding="utf-8-sig", newline="") as f:
        lector = csv.reader(f, delimiter=";")
        next(lector)
        return sum(1 for _ in lector)


def _contar_filas_coma(ruta: Path) -> int:
    with ruta.open(encoding="utf-8", newline="") as f:
        lector = csv.reader(f, delimiter=",")
        next(lector)
        return sum(1 for _ in lector)


def _guardar_partes(c: Conductor, destino: Path, destino_tipo: str, formato: str) -> dict:
    """Abre "Guardar el resultado", elige destino y formato, verifica semáforo y partes y guarda."""
    p7 = c.paso(7)
    previos = len(c.guardados)
    p7.boton_guardar.invoke()
    c.e.esperar(lambda: len(c.guardados) == previos + 1, 60)
    v = c.guardados[-1]
    c.e.esperar(lambda: v._check is not None, 600)
    if v.v_destino.get() != destino_tipo:
        v.rb_destino[destino_tipo].invoke()
        c.e.esperar(lambda: v._check is not None, 600)
    if v.v_formato.get() != formato:
        v.rb_formato[formato].invoke()
        c.e.esperar(lambda: v._check is not None, 600)
    hechos: dict = {"nivel": v._check.level, "razones": list(v._check.reasons)}
    # 🔴: obliga a dividir (no deja "todo en un solo archivo") y deja elegido cómo
    assert v._check.level == "obligatorio", (destino_tipo, v._check.level, v._check.reasons)
    assert v.rb_uno.instate(["disabled"])
    assert v.v_modo.get() != "none"
    c.e.esperar(lambda: v._partes is not None, 600)
    n_partes = len(v._partes)
    # La cantidad de partes correcta, por aritmética independiente de los límites de
    # `limites_default.json`: Excel recomienda <= 300.000 filas por parte; Google Sheets <= 5 M de
    # celdas por parte (y 50 MB), y siempre por debajo de los máximos duros.
    if destino_tipo == "excel":
        esperadas = math.ceil(ESTIMADAS / 300_000)
    else:
        esperadas = math.ceil(ESTIMADAS * COLUMNAS / 5_000_000)
    hechos["partes_esperadas"] = esperadas
    hechos["partes_sugeridas"] = v._check.suggested_parts
    hechos["partes_en_la_vista_previa"] = n_partes
    assert v._check.suggested_parts == esperadas, (
        destino_tipo,
        v._check.suggested_parts,
        esperadas,
    )
    assert n_partes in (esperadas, esperadas + 1), (destino_tipo, n_partes, esperadas)
    assert sum(p.rows for p in v._partes) == ESTIMADAS
    assert v.lbl_vista.cget("text").startswith(f"Se van a generar {n_partes} archivos")
    hechos["modo"] = v.v_modo.get()
    hechos["partes_filas"] = [p.rows for p in v._partes]
    hechos["filas_por_parte_sugeridas"] = v.v_filas.get()
    c.e.destino = destino
    c.e.esperar(lambda: v.puede_guardar(), 600)
    t0 = time.monotonic()
    v._guardar()
    c.e.esperar(lambda: v._resultado is not None or bool(c.e.carteles), 1800)
    assert c.e.carteles == [], c.e.carteles
    r = v._resultado
    hechos["guardar_s"] = round(time.monotonic() - t0, 1)
    assert r.rows_total == ESTIMADAS == sum(p.rows for p in r.parts)
    archivos = sorted({p.path for p in r.parts})
    assert len(archivos) == n_partes
    hechos["mb_por_archivo"] = [round(a.stat().st_size / 1e6, 1) for a in archivos]
    leer = _contar_filas if formato == "csv_excel" else _contar_filas_coma
    por_archivo = [leer(a) for a in archivos]
    assert sum(por_archivo) == ESTIMADAS, por_archivo
    if destino_tipo == "excel":
        assert max(por_archivo) <= MAX_FILAS_EXCEL - 1
    else:
        assert max(por_archivo) * COLUMNAS <= MAX_CELDAS_SHEETS
    assert "coincide con el resultado" in v.lbl_resultado.cget("text")
    hechos["filas_por_archivo"] = por_archivo
    v.btn_carpeta.invoke()
    assert c.carpetas and c.carpetas[-1] == r.folder
    v.cerrar()
    c.e.raiz.update()
    return hechos


def recorrer(c: Conductor, grandes: Path, salida: Path, esperado: dict, log=print) -> dict:
    hechos: dict = {}
    t0 = time.monotonic()

    def marca(texto: str) -> None:
        import resource

        pico = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss // 1024
        log(f"[{time.monotonic() - t0:7.1f} s] pico {pico} MB: {texto}")

    c.csv = CSV_GRANDE
    c.abrir()
    c.elegir_principal(grandes / "clientes_grande.csv")
    marca("principal leído")
    c.siguiente()
    c.marcar_claves("ID cliente")
    c.siguiente()
    c.agregar(grandes / "ventas_grande.csv")
    marca("ventas leídas y comparadas")
    assert c.indicador("ventas_grande.csv").startswith("✅"), c.indicador("ventas_grande.csv")
    c.siguiente()
    c.esperar_paso4(1800)
    marca("paso 4 listo")
    p4 = c.paso(4)
    # --- paso 4: elegir detalle -> el aviso de tamaño aparece ya acá
    assert p4.etiqueta_tamano.cget("text") == ""
    c.modo_detalle("ventas_grande.csv")
    aviso4 = p4.etiqueta_tamano.cget("text")
    assert aviso4.startswith("🔴") and "2.519.998 filas" in aviso4, aviso4
    assert "no entra en una sola hoja de Excel" in aviso4 and "dividirlo" in aviso4
    assert "2.519.998 filas" in p4.etiqueta_total.cget("text")
    hechos["aviso_paso4"] = aviso4
    c.siguiente()
    # --- paso 5: el indicador baja al desmarcar columnas
    c.esperar_tamano(1800)
    p5 = c.paso(5)
    inicial = p5.etiqueta_tamano.cget("text")
    assert inicial.startswith("2.519.998 filas × 8 columnas"), inicial
    celdas0 = p5.calculo.diagnostico.size.cells
    assert celdas0 in (ESTIMADAS * COLUMNAS, (ESTIMADAS + 1) * COLUMNAS), celdas0
    assert "No entra en Google Sheets" in p5.etiqueta_sheets.cget("text")
    assert "No entra en una hoja de Excel" in p5.etiqueta_excel.cget("text")
    hechos["razones_sheets_8_columnas"] = list(p5.calculo.sheets.reasons)
    assert any("20.000.000 celdas" in r for r in p5.calculo.sheets.reasons), (
        p5.calculo.sheets.reasons
    )
    sec_c = c.seccion("clientes_grande.csv")
    sec_c.lista.casillas["Vendedor"].invoke()
    sec_c.lista.casillas["Nombre"].invoke()
    c.e.raiz.update()
    c.esperar_tamano(1800)
    menor = p5.etiqueta_tamano.cget("text")
    assert menor.startswith("2.519.998 filas × 6 columnas"), menor
    assert p5.calculo.diagnostico.size.cells < celdas0
    assert p5.calculo.diagnostico.size.cells in (ESTIMADAS * 6, (ESTIMADAS + 1) * 6), (
        p5.calculo.diagnostico.size.cells
    )
    # en Google Sheets pasa de "no entra" a "entra" (15 M < 20 M), y Excel sigue sin entrar
    # Google Sheets sigue en 🔴, pero ya no por las celdas (15 M < 20 M): ahora por el peso
    assert p5.calculo.sheets.level == "obligatorio"
    assert not any("20.000.000 celdas" in r for r in p5.calculo.sheets.reasons), (
        p5.calculo.sheets.reasons
    )
    assert any("100,0 MB" in r for r in p5.calculo.sheets.reasons), p5.calculo.sheets.reasons
    assert "No entra en una hoja de Excel" in p5.etiqueta_excel.cget("text")
    hechos["indicador_paso5"] = {"todas": inicial, "sin_dos": menor}
    # se vuelven a marcar (para la prueba de Google Sheets con 8 columnas)
    sec_c.lista.casillas["Vendedor"].invoke()
    sec_c.lista.casillas["Nombre"].invoke()
    c.e.raiz.update()
    c.esperar_tamano(1800)
    assert p5.etiqueta_tamano.cget("text").startswith("2.519.998 filas × 8 columnas")
    marca("paso 5 verificado")
    c.siguiente()
    c.siguiente()  # paso 6 -> 7 (sin resumen final)
    c.esperar_revision(1800)
    p7 = c.paso(7)
    assert "2.519.998 filas" in p7.etiqueta_total.cget("text")
    assert p7.etiqueta_semaforo.cget("text").startswith("🔴 Excel")
    assert p7.botones_reduccion, "no ofrece 'Achicar el resultado'"
    textos_red = [b.cget("text") for b in p7.botones_reduccion]
    assert any("resumen por ID cliente" in t for t in textos_red), textos_red
    hechos["achicar"] = textos_red
    marca("paso 7: revisión con 🔴 y sugerencias")
    res = c.unir(3600)
    assert len(res.df) == ESTIMADAS
    marca("unido en memoria")
    # --- guardar: Excel (🔴, obliga a dividir) y Google Sheets (🔴, obliga a dividir)
    (salida / "excel").mkdir(parents=True, exist_ok=True)
    (salida / "sheets").mkdir(parents=True, exist_ok=True)
    hechos["excel"] = _guardar_partes(c, salida / "excel" / "unido.csv", "excel", "csv_excel")
    marca("guardado para Excel")
    # Al abrir "Guardar" otra vez, viene preelegida la división del guardado anterior (la receta
    # la recuerda): se anota y se vuelve a abrir como si fuera la primera vez.
    p7.boton_guardar.invoke()
    c.e.esperar(lambda: len(c.guardados) == 2, 60)
    heredada = c.guardados[-1]
    c.e.esperar(lambda: heredada._check is not None, 600)
    hechos["sheets_heredada"] = {
        "modo": heredada.v_modo.get(),
        "filas": heredada.v_filas.get(),
        "destino": heredada.v_destino.get(),
        "formato": heredada.v_formato.get(),
    }
    # Al pasar de Excel a Google Sheets se vuelve a la división sugerida (F5-A)
    heredada.rb_destino["google_sheets"].invoke()
    c.e.esperar(lambda: heredada._check is not None and heredada._partes is not None, 600)
    c.e.esperar(lambda: heredada.v_destino.get() == "google_sheets", 60)
    hechos["sheets_heredada"]["al_cambiar"] = {
        "partes_sugeridas": heredada._check.suggested_parts,
        "filas_sugeridas": heredada._check.suggested_rows_per_part,
        "filas": heredada.v_filas.get(),
        "partes": len(heredada._partes),
    }
    heredada.cerrar()
    c.estado.salida = None
    hechos["sheets"] = _guardar_partes(
        c, salida / "sheets" / "unido.csv", "google_sheets", "csv_google"
    )
    marca("guardado para Google Sheets")

    # --- achicar el resultado: pasar ventas a resumen por la clave, con un clic
    boton = next(b for b in p7.botones_reduccion if "resumen por ID cliente" in b.cget("text"))
    assert boton.cget("text").rstrip().endswith(("🟢", "🟡"))  # la alternativa sí entra en Excel
    boton.invoke()
    c.e.raiz.update()
    assert "Listo: " in p7.etiqueta_guardado.cget("text")
    c.esperar_revision(1800)
    assert "200.000 filas" in p7.etiqueta_total.cget("text")
    assert p7.etiqueta_semaforo.cget("text").startswith("🟢"), p7.etiqueta_semaforo.cget("text")
    res2 = c.unir(1800)
    assert len(res2.df) == 200_000
    hechos["achicado"] = p7.etiqueta_semaforo.cget("text")
    marca("achicado a 200.000 filas (🟢)")

    import os

    if os.environ.get("LIBERAR"):  # experimento de memoria: suelta lo que retiene la UI
        import gc

        c.guardados.clear()
        p7.ventana_guardar = None
        p7.resultado = None
        p7.revision = None
        gc.collect()
        marca("liberado lo retenido (experimento)")
    # --- resumen final con > 200.000 filas: usa una muestra y lo avisa
    for _ in range(3):  # 7 -> 4
        c.anterior()
    c.esperar_paso4(1800)
    c.modo_detalle("ventas_grande.csv")
    marca("C0: vuelve a detalle (paso 4)")
    c.siguiente()
    c.esperar_tamano(1800)
    marca("C0b: paso 5 recalculado")
    c.siguiente()
    p6 = c.paso(6)
    p6.radio_si.invoke()
    c.e.esperar(lambda: p6.editor is not None or bool(c.e.carteles), 1800)
    assert p6.tabla is not None and p6.tabla.es_muestra
    marca("C1: tabla de muestra para el resumen final lista")
    # la muestra es el resultado de unir solo los primeros 5.000 clientes: bastante menos que todo
    assert p6.tabla.filas_total == ESTIMADAS
    assert 5_000 <= len(p6.tabla.df) < 200_000, len(p6.tabla.df)
    avisos = [t for t in _textos_dialogo(p6.dialogo) if "muestra" in t]
    n_muestra = f"{len(p6.tabla.df):,}".replace(",", ".")
    assert avisos and "2.519.998 filas" in avisos[0] and f"{n_muestra} filas" in avisos[0], avisos
    assert "Al unir se calcula con todo" in avisos[0]
    hechos["aviso_muestra"] = avisos[0]
    ed = p6.editor
    ed.checks["Región"].invoke()
    fila = ed.agregar_calculo()
    fila.var_operacion.set("Contar filas")
    fila.combo_operacion.event_generate("<<ComboboxSelected>>")
    fila.var_nombre.set("Filas")
    c.e.raiz.update()
    p6.boton_listo.invoke()
    c.e.raiz.update()
    c.siguiente()
    c.esperar_revision(1800)
    marca("C2: revisión (diagnóstico) del plan con resumen final")
    res3 = c.unir(3600)
    assert len(res3.df) == 5
    por_region = {r["Región"]: int(r["Filas"]) for _, r in res3.df.iterrows()}
    assert por_region == esperado["detalle_left"]["por_region"], por_region
    hechos["resumen_final"] = por_region
    marca("resumen final por Región sobre el resultado completo")
    assert c.e.carteles == []
    return hechos


def _textos_dialogo(dialogo) -> list[str]:
    from .asistente_util import textos

    return textos(dialogo)


def cargar_esperado(fixtures_esperados: Path) -> dict:
    return json.loads((fixtures_esperados / "grandes.json").read_text(encoding="utf-8"))
