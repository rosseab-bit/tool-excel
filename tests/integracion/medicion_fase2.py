"""Medición de la Fase 2 (no es un test). Se corre con /usr/bin/time -v (pico de memoria):

    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase2.py perfil
    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase2.py distintos
    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase2.py filtro_orden
    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase2.py xlsx

Escenarios: perfil = leer + perfilar el CSV de 2M; distintos = perfil + valores distintos de ID (2M
únicos); filtro_orden = perfil + filtro y orden numéricos; xlsx = leer + perfilar el xlsx de 500k.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path

from excel_tool.core.explore import FilterCondition, filter_rows, get_page, sort_rows, value_counts
from excel_tool.core.io import read_table
from excel_tool.core.profile import profile_table

GRANDES = Path(__file__).resolve().parent.parent / "fixtures" / "generados" / "grandes"


def _paso(nombre: str, t0: float) -> float:
    ahora = time.monotonic()
    print(f"  {nombre}: {ahora - t0:.1f} s", flush=True)
    return ahora


def main(escenario: str) -> None:
    ruta = GRANDES / ("rendimiento_500k.xlsx" if escenario == "xlsx" else "rendimiento_2M.csv")
    t0 = time.monotonic()
    leido = read_table(ruta)
    t = _paso("leer", t0)
    perfil = profile_table(leido)
    t = _paso("perfilar", t)
    print(f"  total leer+perfil: {time.monotonic() - t0:.1f} s ({perfil.rows} filas)", flush=True)
    tipos = {c.name: c.detected_type for c in perfil.columns_profile}
    if escenario == "distintos":
        r = value_counts(leido.df, ["ID"])
        t = _paso(f"distintos de ID ({r.distinct_total} valores, truncado={r.truncated})", t)
    elif escenario == "filtro_orden":
        filas = filter_rows(leido.df, [FilterCondition("Importe", "mayor", "500000")], tipos)
        t = _paso(f"filtro Importe > 500000 ({len(filas)} filas)", t)
        orden = sort_rows(leido.df, filas, "Importe", False, tipos["Importe"])
        t = _paso("orden por Importe", t)
        todas = sort_rows(leido.df, leido.df.index, "Importe", True, tipos["Importe"])
        t = _paso(f"orden de todas las filas ({len(todas)})", t)
        pagina = get_page(leido.df, orden, 0)
        _paso(f"pagina ({len(pagina)} filas)", t)
    print(f"  TOTAL: {time.monotonic() - t0:.1f} s", flush=True)


if __name__ == "__main__" and sys.argv[1] != "ui":
    main(sys.argv[1])


def medir_ui() -> None:
    """Pico de memoria por el camino real de la UI (correr bajo xvfb):
    xvfb-run -a /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase2.py ui"""
    import tempfile
    import time as _t

    from excel_tool.core.io import CsvOptions
    from excel_tool.ui.app import crear_ventana
    from excel_tool.ui.explorador import abrir_explorador
    from excel_tool.ui.selector import Seleccion

    raiz = crear_ventana()
    t0 = _t.monotonic()
    ex = abrir_explorador(
        raiz, Seleccion(GRANDES / "rendimiento_2M.csv", None, CsvOptions(";", "utf-8-sig"), 0)
    )

    def esperar(cond, tope=300):
        limite = _t.monotonic() + tope
        while _t.monotonic() < limite and not cond():
            raiz.update()
            _t.sleep(0.02)

    esperar(lambda: ex.perfil is not None and not ex.analizando)
    print(f"  UI: perfil listo a los {_t.monotonic() - t0:.1f} s", flush=True)
    t = _t.monotonic()
    ex.seleccionar_columna("ID")
    ex._ver_distintos()
    esperar(lambda: ex.panel_distintos.resultado is not None)
    print(f"  UI: distintos de ID en {_t.monotonic() - t:.1f} s", flush=True)
    t = _t.monotonic()
    datos = ex.panel_datos
    datos.agregar_condicion("Importe", "es mayor que", "500000")
    datos.aplicar()
    esperar(lambda: len(datos.filas) < 2_000_000)
    print(f"  UI: filtro en {_t.monotonic() - t:.1f} s ({len(datos.filas)} filas)", flush=True)
    t = _t.monotonic()
    datos.grilla._clic_encabezado("Importe")
    esperar(lambda: datos._orden is not None)
    esperar(lambda: not datos.boton_aplicar.instate(["disabled"]))
    print(
        f"  UI: orden en {_t.monotonic() - t:.1f} s; "
        f"grilla muestra {datos.grilla.filas_mostradas()} filas",
        flush=True,
    )
    ex.panel_distintos.var_juntar.set(False)
    destino = Path(tempfile.mkdtemp()) / "informe.xlsx"
    from excel_tool.ui import exportar

    t = _t.monotonic()
    exportar.guardar_informe(ex.perfil, list(ex.panel_distintos.calculados.values()), destino)
    print(f"  UI: informe en {_t.monotonic() - t:.1f} s", flush=True)
    raiz.destroy()


if __name__ == "__main__" and sys.argv[1] == "ui":
    medir_ui()
