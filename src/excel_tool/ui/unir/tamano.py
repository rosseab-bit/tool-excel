"""Cálculo del paso Columnas: nombres finales de las columnas y tamaño del resultado.

Todo lo calculado sale del core (`final_column_names`, `diagnose_join`, `limits.check`,
`suggest_reductions` y `run_join`). Las funciones pensadas para el worker no tocan widgets.
"""

from __future__ import annotations

import dataclasses
import logging
from collections.abc import Callable
from dataclasses import dataclass

import pandas as pd

from excel_tool.core.aggregate import AggregationSpec, group_column_name, result_columns
from excel_tool.core.join import (
    JoinDiagnosis,
    JoinPlan,
    Reduction,
    SecondarySpec,
    diagnose_join,
    final_column_names,
    run_join,
    suggest_reductions,
)
from excel_tool.core.limits import SizeCheck, check, load_limits
from excel_tool.core.progress import CancelToken, ProgressCallback

log = logging.getLogger(__name__)

# Igual que `diagnose_join`; el asistente pasa `estado.diagnosticar` (con memoria).
Diagnosticar = Callable[..., JoinDiagnosis]

FILAS_PARA_ARMAR_RESUMEN = 200_000  # hasta acá, el editor del resumen final trabaja con todo
FILAS_MUESTRA_RESUMEN = 5_000  # filas del principal cuando el resultado es más grande


@dataclass(frozen=True)
class Calculo:
    nombres: dict[tuple[str, str], str]  # (archivo, columna de origen) -> nombre en el resultado
    diagnostico: JoinDiagnosis
    excel: SizeCheck
    sheets: SizeCheck


def columnas_de_resumen(summary: AggregationSpec, claves: set[str]) -> list[str]:
    """Columnas que aporta un archivo en resumen (sin la clave, que sale una sola vez)."""
    en_grupo = {
        group_column_name(g) for g in summary.group_by if g.period is None and g.column in claves
    }
    return [c for c in result_columns(summary) if c not in en_grupo]


def columnas_de_detalle(sec: SecondarySpec) -> list[str]:
    claves = {q for _, q in sec.key_pairs}
    vistas: list[str] = []
    for c in sec.columns:
        if c not in claves and c not in vistas:
            vistas.append(c)
    return vistas


def nombres_finales(plan: JoinPlan, tablas: dict[str, pd.DataFrame]) -> dict[tuple[str, str], str]:
    """(archivo, columna de origen) -> nombre en el resultado, sin unir (`join.final_column_names`).

    La clave del principal figura con el nombre del archivo principal. Si el core no puede
    armar los nombres, devuelve {} y el paso muestra los de origen.
    """
    try:
        lista = final_column_names(dataclasses.replace(plan, final_summary=None), tablas)
    except Exception:  # noqa: BLE001 - los nombres son solo para mostrar; se usan los de origen
        log.exception("No se pudieron calcular los nombres finales de las columnas")
        return {}
    return {(a if a is not None else plan.principal.name, o): f for a, o, f in lista}


def calcular(
    plan: JoinPlan,
    tablas: dict[str, pd.DataFrame],
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
    diagnosticar: Diagnosticar = diagnose_join,
) -> Calculo:
    """Nombres finales y tamaño del resultado (sin el resumen final, que es otro paso)."""
    plan = dataclasses.replace(plan, final_summary=None)
    diag = diagnosticar(plan, tablas, progress=progress, cancel=cancel)
    nombres = nombres_finales(plan, tablas)
    limites = load_limits()
    return Calculo(
        nombres,
        diag,
        check(diag.size, "excel", "xlsx", limits=limites),
        check(diag.size, "google_sheets", "csv", limits=limites),
    )


@dataclass(frozen=True)
class Revision:
    """Lo que muestra el paso 7 antes de unir."""

    diagnostico: JoinDiagnosis
    excel: SizeCheck
    sheets: SizeCheck
    reducciones: list[tuple[Reduction, SizeCheck]]  # cada una con su semáforo de Excel


def revisar(
    plan: JoinPlan,
    tablas: dict[str, pd.DataFrame],
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
    diagnosticar: Diagnosticar = diagnose_join,
) -> Revision:
    """Diagnóstico del plan completo (con el resumen final) y, si Excel da 🟡 o 🔴, las
    alternativas para achicar el resultado."""
    diag = diagnosticar(plan, tablas, progress=progress, cancel=cancel)
    limites = load_limits()
    excel = check(diag.size, "excel", "xlsx", limits=limites)
    sheets = check(diag.size, "google_sheets", "csv", limits=limites)
    reducciones: list[tuple[Reduction, SizeCheck]] = []
    if excel.level != "ok":
        for r in suggest_reductions(plan, tablas, progress=progress, cancel=cancel):
            reducciones.append((r, check(r.size, "excel", "xlsx", limits=limites)))
    return Revision(diag, excel, sheets, reducciones)


@dataclass(frozen=True)
class TablaParaResumen:
    df: pd.DataFrame
    column_types: dict[str, str]
    es_muestra: bool
    filas_total: int


def tabla_para_resumen(
    plan: JoinPlan,
    tablas: dict[str, pd.DataFrame],
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
    diagnosticar: Diagnosticar = diagnose_join,
) -> TablaParaResumen:
    """La tabla unida (sin resumen final) sobre la que se arma el resumen final.

    Si el resultado tiene más de `FILAS_PARA_ARMAR_RESUMEN` filas se une solo una muestra
    del principal: el resumen se arma igual y se calcula completo al unir.
    """
    plan = dataclasses.replace(plan, final_summary=None, output_column_order=None)
    diag = diagnosticar(plan, tablas, progress=progress, cancel=cancel)
    muestra = diag.rows_result > FILAS_PARA_ARMAR_RESUMEN
    if muestra:
        nombre = plan.principal.name
        tablas = {**tablas, nombre: tablas[nombre].head(FILAS_MUESTRA_RESUMEN)}
    res = run_join(plan, tablas, progress=progress, cancel=cancel)
    return TablaParaResumen(res.df, dict(res.column_types), muestra, diag.rows_result)
