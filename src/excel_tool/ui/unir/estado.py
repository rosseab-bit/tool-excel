"""Estado compartido del asistente para unir archivos.

Lógica de UI pura (sin tkinter): guarda lo que el usuario fue eligiendo y arma el `JoinPlan` del
core. No calcula nada: coincidencias, tamaños y resúmenes los hace el core.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Literal

import pandas as pd

from excel_tool.core.aggregate import AggregationSpec
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io import read_table
from excel_tool.core.join import JoinDiagnosis, JoinPlan, SecondarySpec, SourceSpec, diagnose_join
from excel_tool.core.normalize import KeyOptions
from excel_tool.core.output import OutputSpec
from excel_tool.core.profile import TableProfile, profile_table
from excel_tool.core.progress import CancelToken, ProgressCallback
from excel_tool.core.recipe import RecipeLoad

Modo = Literal["detail", "summary", "first"]
Union = Literal["left", "inner"]


@dataclass
class Fuente:
    """Un archivo del asistente: qué es (`spec`), lo leído (`df`, sin copia) y su perfil.

    Si el archivo de una receta no existe, `df` y `perfil` quedan en None (`falta`).
    """

    spec: SourceSpec
    df: pd.DataFrame | None = None
    perfil: TableProfile | None = None

    @property
    def nombre(self) -> str:
        return self.spec.name

    @property
    def falta(self) -> bool:
        return self.df is None


@dataclass
class FuenteSecundaria(Fuente):
    key_pairs: list[tuple[str, str]] = field(default_factory=list)
    how: Union = "left"
    mode: Modo | None = None  # None hasta que el usuario elige (nunca hay opción por defecto)
    summary: AggregationSpec | None = None
    prefix: str | None = None


def leer_fuente(
    spec: SourceSpec,
    *,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> Fuente:
    """Lee el archivo y arma su perfil. Pensada para correr en el worker."""
    leido = read_table(spec.path, spec.read_options, progress=progress, cancel=cancel)
    perfil = profile_table(leido, progress=progress, cancel=cancel)
    return Fuente(spec, leido.df, perfil)


@dataclass
class EstadoUnion:
    principal: Fuente | None = None
    claves: list[str] = field(default_factory=list)
    key_options: KeyOptions = field(default_factory=KeyOptions)
    secundarios: list[FuenteSecundaria] = field(default_factory=list)
    columnas: dict[str, list[str]] = field(default_factory=dict)  # por nombre de fuente
    orden: list[str] = field(default_factory=list)
    resumen_final: AggregationSpec | None = None
    salida: OutputSpec | None = None  # solo para conservar lo de una receta
    _cache_diagnosticos: dict = field(default_factory=dict, repr=False, compare=False)

    # ---- consultas ----
    @property
    def hay_datos(self) -> bool:
        """¿Hay algo cargado que se perdería al cancelar?"""
        return self.principal is not None or bool(self.secundarios)

    @property
    def faltantes(self) -> list[str]:
        """Nombres de los archivos que la receta pedía y no se encontraron (hook de F5-11)."""
        fuentes = ([self.principal] if self.principal else []) + self.secundarios
        return [f.nombre for f in fuentes if f.falta]

    def fuentes(self) -> list[Fuente]:
        return ([self.principal] if self.principal else []) + list(self.secundarios)

    def tablas(self) -> dict[str, pd.DataFrame]:
        """Las tablas ya leídas por nombre de fuente, como las pide el core."""
        tablas: dict[str, pd.DataFrame] = {}
        for fuente in self.fuentes():
            if fuente.df is None:
                raise ExcelToolError(f"Falta elegir el archivo «{fuente.nombre}».")
            tablas[fuente.nombre] = fuente.df
        return tablas

    # ---- diagnóstico con memoria ----
    def _firma(self, plan: JoinPlan, tablas: dict[str, pd.DataFrame]) -> str:
        """Firma del plan y de las tablas. El orden de las columnas no cambia ni las filas ni
        el tamaño, así que no entra en la firma."""
        sin_orden = dataclasses.replace(plan, output_column_order=None)
        claves = [(n, id(df), df.shape) for n, df in tablas.items()]
        return f"{sin_orden!r}|{claves!r}"

    def diagnosticar(
        self,
        plan: JoinPlan,
        tablas: dict[str, pd.DataFrame],
        progress: ProgressCallback | None = None,
        cancel: CancelToken | None = None,
    ) -> JoinDiagnosis:
        """`diagnose_join` del core, reusando el resultado si el plan y las tablas no cambiaron
        (cada llamada completa cuesta cientos de MB pasajeros en archivos enormes)."""
        firma = self._firma(plan, tablas)
        cache = self._cache_diagnosticos
        if firma in cache:
            return cache[firma]
        diag = diagnose_join(plan, tablas, progress=progress, cancel=cancel)
        if len(cache) >= 6:
            cache.pop(next(iter(cache)))
        cache[firma] = diag
        return diag

    def olvidar_diagnosticos(self) -> None:
        self._cache_diagnosticos.clear()

    # ---- armado del plan ----
    def _columnas_de(self, fuente: Fuente) -> list[str]:
        if fuente.nombre in self.columnas:
            return list(self.columnas[fuente.nombre])
        return list(fuente.spec.columns)

    def plan(self) -> JoinPlan:
        """Arma el `JoinPlan`. Si falta algo, `ExcelToolError` con el motivo en castellano."""
        if self.principal is None:
            raise ExcelToolError("Falta elegir el archivo principal.")
        if not self.secundarios:
            raise ExcelToolError("Falta sumar al menos un archivo.")
        for fuente in self.fuentes():
            if fuente.falta:
                raise ExcelToolError(f"Falta elegir el archivo «{fuente.nombre}».")
        principal = dataclasses.replace(
            self.principal.spec, columns=self._columnas_de(self.principal)
        )
        secundarios: list[SecondarySpec] = []
        for sec in self.secundarios:
            if sec.mode is None:
                raise ExcelToolError(
                    f"Falta elegir si de «{sec.nombre}» querés ver el detalle o un resumen."
                )
            if not sec.key_pairs:
                raise ExcelToolError(f"Falta elegir la columna en común de «{sec.nombre}».")
            if sec.mode == "summary" and sec.summary is None:
                raise ExcelToolError(f"Falta armar el resumen de «{sec.nombre}».")
            secundarios.append(
                SecondarySpec(
                    name=sec.spec.name,
                    path=sec.spec.path,
                    read_options=sec.spec.read_options,
                    columns=self._columnas_de(sec),
                    key_pairs=[tuple(p) for p in sec.key_pairs],  # type: ignore[misc]
                    how=sec.how,
                    mode=sec.mode,
                    summary=sec.summary if sec.mode == "summary" else None,
                    prefix=sec.prefix,
                )
            )
        return JoinPlan(
            principal=principal,
            key_options=self.key_options,
            secondaries=secundarios,
            output_column_order=list(self.orden) or None,
            final_summary=self.resumen_final,
            output=self.salida,
        )

    def aplicar_plan(self, plan: JoinPlan) -> None:
        """Vuelca un plan alternativo (por ejemplo una sugerencia de "achicar el resultado")
        en el estado: modos, resúmenes, columnas, orden y resumen final."""
        self.columnas[plan.principal.name] = list(plan.principal.columns)
        por_nombre = {s.name: s for s in plan.secondaries}
        for sec in self.secundarios:
            nuevo = por_nombre.get(sec.nombre)
            if nuevo is None:
                continue
            sec.mode = nuevo.mode
            sec.summary = nuevo.summary
            self.columnas[sec.nombre] = list(nuevo.columns)
        self.orden = list(plan.output_column_order or [])
        self.resumen_final = plan.final_summary

    # ---- archivos que faltan ----
    def completar(self, nombre: str, nuevo: Fuente) -> None:
        """Reemplaza el archivo `nombre` por `nuevo` conservando el resto de lo elegido."""

        def spec_nuevo(viejo: Fuente) -> SourceSpec:
            return dataclasses.replace(
                viejo.spec, path=nuevo.spec.path, read_options=nuevo.spec.read_options
            )

        if self.principal is not None and self.principal.nombre == nombre:
            self.principal = Fuente(spec_nuevo(self.principal), nuevo.df, nuevo.perfil)
            return
        for i, sec in enumerate(self.secundarios):
            if sec.nombre == nombre:
                self.secundarios[i] = dataclasses.replace(
                    sec, spec=spec_nuevo(sec), df=nuevo.df, perfil=nuevo.perfil
                )
                return
        raise ExcelToolError(f"No hay ningún archivo llamado «{nombre}».")

    # ---- recetas ----
    @classmethod
    def desde_receta(
        cls,
        carga: RecipeLoad,
        leer: Callable[[SourceSpec], Fuente] | None = None,
    ) -> EstadoUnion:
        """Arma el estado desde una receta. Lee los archivos que existen con `leer`
        (por defecto `leer_fuente`, que conviene correr en el worker); los que faltan quedan
        marcados (`faltantes`)."""
        leer = leer or leer_fuente
        plan = carga.plan

        def fuente(spec: SourceSpec) -> Fuente:
            return Fuente(spec) if spec.name in carga.missing else leer(spec)

        principal = fuente(plan.principal)
        secundarios: list[FuenteSecundaria] = []
        for sec in plan.secondaries:
            base = fuente(sec)
            secundarios.append(
                FuenteSecundaria(
                    spec=SourceSpec(sec.name, sec.path, sec.read_options, list(sec.columns)),
                    df=base.df,
                    perfil=base.perfil,
                    key_pairs=[(a, b) for a, b in sec.key_pairs],
                    how=sec.how,
                    mode=sec.mode,
                    summary=sec.summary,
                    prefix=sec.prefix,
                )
            )
        claves: list[str] = []
        for sec in plan.secondaries:
            for col, _ in sec.key_pairs:
                if col not in claves:
                    claves.append(col)
        columnas = {plan.principal.name: list(plan.principal.columns)}
        columnas.update({s.name: list(s.columns) for s in plan.secondaries})
        return cls(
            principal=principal,
            claves=claves,
            key_options=plan.key_options,
            secundarios=secundarios,
            columnas=columnas,
            orden=list(plan.output_column_order or []),
            resumen_final=plan.final_summary,
            salida=plan.output,
        )
