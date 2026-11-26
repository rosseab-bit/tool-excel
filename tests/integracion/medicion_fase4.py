"""Medición de la Fase 4 (no es un test). Se corre con /usr/bin/time -v (pico de memoria):

    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase4.py detalle
    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase4.py resumen

Escenario grande: `clientes_grande.csv` (200.000 filas) como principal y `ventas_grande.csv`
(2.500.000 filas) como secundario, en detalle o en resumen por cliente (suma y cantidad). Imprime
una línea `RESULTADO {...}` con las filas que estimó `diagnose_join`, las reales de `run_join` y los
tiempos. El presupuesto de memoria es 1,5 GB (ADR-023).
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy
from excel_tool.core.io import ReadOptions, read_table
from excel_tool.core.join import JoinPlan, SecondarySpec, SourceSpec, diagnose_join, run_join
from excel_tool.core.normalize import KeyOptions

GRANDES = Path(__file__).resolve().parent.parent / "fixtures" / "generados" / "grandes"


def main(modo: str) -> None:
    ro = ReadOptions()
    t0 = time.monotonic()
    clientes = read_table(GRANDES / "clientes_grande.csv", ro).df
    ventas = read_table(GRANDES / "ventas_grande.csv", ro).df
    t_leer = time.monotonic() - t0
    principal = SourceSpec(
        "clientes", GRANDES / "clientes_grande.csv", ro, ["ID cliente", "Nombre", "Región"]
    )
    columnas = ["Fecha", "Artículo", "Cantidad", "Importe"]
    if modo == "detalle":
        sec = SecondarySpec(
            "ventas", GRANDES / "ventas_grande.csv", ro, columnas, [("ID cliente", "ID cliente")]
        )
    else:
        resumen = AggregationSpec(
            [GroupBy("ID cliente")],
            [Aggregation("Importe", "sum", "Total comprado"), Aggregation("", "count")],
        )
        sec = SecondarySpec(
            "ventas",
            GRANDES / "ventas_grande.csv",
            ro,
            columnas,
            [("ID cliente", "ID cliente")],
            mode="summary",
            summary=resumen,
        )
    plan = JoinPlan(principal, KeyOptions(), [sec])
    tablas = {"clientes": clientes, "ventas": ventas}
    t1 = time.monotonic()
    diag = diagnose_join(plan, tablas)
    t_diag = time.monotonic() - t1
    t2 = time.monotonic()
    res = run_join(plan, tablas)
    t_join = time.monotonic() - t2
    print(
        "RESULTADO "
        + json.dumps(
            {
                "modo": modo,
                "filas_estimadas": diag.rows_result,
                "filas_reales": len(res.df),
                "columnas": len(res.df.columns),
                "leer_s": round(t_leer, 1),
                "diagnostico_s": round(t_diag, 1),
                "unir_s": round(t_join, 1),
                "memoria_ok": diag.memory.ok,
                "avisos": len(res.warnings),
            }
        ),
        flush=True,
    )


if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1] not in {"detalle", "resumen"}:
        raise SystemExit("uso: medicion_fase4.py detalle|resumen")
    main(sys.argv[1])
