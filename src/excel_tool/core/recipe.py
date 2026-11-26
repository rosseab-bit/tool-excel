"""Recetas: guardar y cargar un `JoinPlan` para repetir el mismo cruce con archivos nuevos.

Contrato: docs/contratos.md § Fase 4 › Recetas.

Formato (JSON utf-8, `"version": 1`): `{"version": 1, "plan": {...}}`. Cada dataclass tiene su
serializador explícito (nada de pickle ni de nombres de clase en el archivo). Al cargar se valida
cada campo; lo que falta o tiene un tipo malo es una receta dañada, y los campos desconocidos se
ignoran con un aviso (compatibilidad hacia adelante).

Rutas: se guardan con `/`, relativas a la carpeta de la receta si el archivo está en esa carpeta o
debajo; si no, absolutas. Una ruta de Windows (`C:\\...`, `C:/...`, `\\\\servidor\\...`) cargada
en Linux queda como texto en un `Path` y figura como faltante si no existe.

Decisiones de nivel 1: `RecipeLoad.missing` lista los NOMBRES de las fuentes (`SourceSpec.name`)
cuyo archivo no existe, para que la UI sepa cuál pedirle al usuario; el detalle con la ruta va en
`warnings`.
"""

from __future__ import annotations

import json
import logging
import os
import re
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath
from typing import Any, get_args

from excel_tool.core.aggregate import AggFunc, Aggregation, AggregationSpec, GroupBy
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io import CsvOptions, ReadOptions, WriteFormat
from excel_tool.core.join import JoinPlan, SecondarySpec, SourceSpec
from excel_tool.core.limits import Destination
from excel_tool.core.normalize import KeyOptions
from excel_tool.core.output import OutputSpec
from excel_tool.core.split import SplitSpec
from excel_tool.core.types import ColumnType

log = logging.getLogger(__name__)

VERSION = 1
_MSG_DANADA = "La receta está dañada o es de una versión más nueva de la app."
_RE_WINDOWS = re.compile(r"^(?:[A-Za-z]:[\\/]|\\\\)")

_FORMATOS = set(get_args(WriteFormat))
_DESTINOS = set(get_args(Destination))
_FUNCS = set(get_args(AggFunc))
_TIPOS = set(get_args(ColumnType))
_MODOS_DIVISION = {"rows", "rows_keep_groups", "by_value", "sheets"}
_PERIODOS = {"month", "year"}
_HOW = {"left", "inner"}
_MODOS = {"detail", "summary", "first"}


@dataclass(frozen=True)
class RecipeLoad:
    plan: JoinPlan
    missing: list[str] = field(default_factory=list)  # nombres de fuentes cuyo archivo no existe
    warnings: list[str] = field(default_factory=list)


class _Invalida(Exception):
    pass


# ---------------------------------------------------------------------- rutas


def _ruta_a_texto(p: Path, base: Path) -> str:
    try:
        rel = p.resolve().relative_to(base.resolve())
        return PurePosixPath(*rel.parts).as_posix() if rel.parts else "."
    except (ValueError, OSError):
        return p.as_posix()


def _texto_a_ruta(texto: str, base: Path) -> Path:
    if _RE_WINDOWS.match(texto) or texto.startswith("/"):
        return Path(texto)  # absoluta (o de otro sistema): se deja como está
    return base.joinpath(*PurePosixPath(texto).parts)


# ---------------------------------------------------------------------- guardar


def _csv_a_json(c: CsvOptions | None) -> dict | None:
    return None if c is None else {"separator": c.separator, "encoding": c.encoding}


def _read_a_json(r: ReadOptions) -> dict:
    return {
        "sheet": r.sheet,
        "header_row": r.header_row,
        "csv": _csv_a_json(r.csv),
        "nrows": r.nrows,
    }


def _agg_a_json(a: AggregationSpec | None) -> dict | None:
    if a is None:
        return None
    return {
        "group_by": [{"column": g.column, "period": g.period} for g in a.group_by],
        "aggregations": [
            {"column": x.column, "func": x.func, "name": x.name} for x in a.aggregations
        ],
    }


def _split_a_json(s: SplitSpec | None) -> dict | None:
    if s is None:
        return None
    return {
        "mode": s.mode,
        "rows_per_part": s.rows_per_part,
        "group_column": s.group_column,
        "value_columns": None if s.value_columns is None else list(s.value_columns),
        "name_pattern": s.name_pattern,
    }


def _salida_a_json(o: OutputSpec | None, base: Path) -> dict | None:
    if o is None:
        return None
    return {
        "path": _ruta_a_texto(o.path, base),
        "format": o.format,
        "destination": o.destination,
        "split": _split_a_json(o.split),
        "write_index": o.write_index,
        "sheet_name": o.sheet_name,
        "column_types": None if o.column_types is None else dict(o.column_types),
    }


def _fuente_a_json(s: SourceSpec, base: Path) -> dict:
    return {
        "name": s.name,
        "path": _ruta_a_texto(s.path, base),
        "read_options": _read_a_json(s.read_options),
        "columns": list(s.columns),
    }


def _secundario_a_json(s: SecondarySpec, base: Path) -> dict:
    d = _fuente_a_json(s, base)
    d.update(
        {
            "key_pairs": [[a, b] for a, b in s.key_pairs],
            "how": s.how,
            "mode": s.mode,
            "summary": _agg_a_json(s.summary),
            "prefix": s.prefix,
        }
    )
    return d


def _plan_a_json(plan: JoinPlan, base: Path) -> dict:
    k = plan.key_options
    return {
        "principal": _fuente_a_json(plan.principal, base),
        "key_options": {
            "trim": k.trim,
            "ignore_case": k.ignore_case,
            "strip_leading_zeros": k.strip_leading_zeros,
        },
        "secondaries": [_secundario_a_json(s, base) for s in plan.secondaries],
        "output_column_order": (
            None if plan.output_column_order is None else list(plan.output_column_order)
        ),
        "final_summary": _agg_a_json(plan.final_summary),
        "output": _salida_a_json(plan.output, base),
    }


def save_recipe(plan: JoinPlan, path: Path) -> None:
    """Guarda `plan` en `path` (JSON). Escritura atómica: primero `.tmp`, después `os.replace`."""
    path = Path(path)
    datos = {"version": VERSION, "plan": _plan_a_json(plan, path.parent)}
    texto = json.dumps(datos, ensure_ascii=False, indent=2) + "\n"
    tmp = path.with_name(path.name + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            f.write(texto)
        os.replace(tmp, path)
    except OSError as e:
        log.exception("No se pudo guardar la receta %s", path)
        tmp.unlink(missing_ok=True)
        raise ExcelToolError(
            "No se pudo guardar la receta. Revisá que la carpeta exista y que puedas escribir.",
            str(e),
        ) from e


# ---------------------------------------------------------------------- cargar


class _Lector:
    def __init__(self, warnings: list[str]) -> None:
        self.warnings = warnings

    def obj(self, v: Any, ctx: str, requeridos: set[str], opcionales: set[str]) -> dict:
        if not isinstance(v, dict):
            raise _Invalida(f"{ctx}: se esperaba un objeto")
        faltan = requeridos - v.keys()
        if faltan:
            raise _Invalida(f"{ctx}: faltan campos {sorted(faltan)}")
        extra = v.keys() - requeridos - opcionales
        if extra:
            log.warning("Receta: campos desconocidos en %s, se ignoran: %s", ctx, sorted(extra))
            self.warnings.append(f"La receta trae datos que esta versión no conoce ({ctx}).")
        return v

    @staticmethod
    def texto(v: Any, ctx: str, *, nulo: bool = False) -> Any:
        if v is None and nulo:
            return None
        if not isinstance(v, str):
            raise _Invalida(f"{ctx}: se esperaba texto")
        return v

    @staticmethod
    def entero(v: Any, ctx: str) -> int | None:
        if v is None:
            return None
        if isinstance(v, bool) or not isinstance(v, int):
            raise _Invalida(f"{ctx}: se esperaba un entero")
        return v

    @staticmethod
    def booleano(v: Any, ctx: str) -> bool:
        if not isinstance(v, bool):
            raise _Invalida(f"{ctx}: se esperaba sí/no")
        return v

    @staticmethod
    def opcion(v: Any, ctx: str, validos: set[str]) -> Any:
        if not isinstance(v, str) or v not in validos:
            raise _Invalida(f"{ctx}: valor no válido {v!r}")
        return v

    @staticmethod
    def lista(v: Any, ctx: str) -> list:
        if not isinstance(v, list):
            raise _Invalida(f"{ctx}: se esperaba una lista")
        return v

    def textos(self, v: Any, ctx: str, *, nulo: bool = False) -> list[str] | None:
        if v is None and nulo:
            return None
        return [self.texto(x, f"{ctx}[]") for x in self.lista(v, ctx)]

    # -- dataclasses

    def csv(self, v: Any, ctx: str) -> CsvOptions | None:
        if v is None:
            return None
        d = self.obj(v, ctx, {"separator", "encoding"}, set())
        return CsvOptions(self.texto(d["separator"], ctx), self.texto(d["encoding"], ctx))

    def read(self, v: Any, ctx: str) -> ReadOptions:
        d = self.obj(v, ctx, set(), {"sheet", "header_row", "csv", "nrows"})
        return ReadOptions(
            sheet=self.texto(d.get("sheet"), ctx + ".sheet", nulo=True),
            header_row=self.entero(d.get("header_row"), ctx + ".header_row"),
            csv=self.csv(d.get("csv"), ctx + ".csv"),
            nrows=self.entero(d.get("nrows"), ctx + ".nrows"),
        )

    def agregado(self, v: Any, ctx: str) -> AggregationSpec | None:
        if v is None:
            return None
        d = self.obj(v, ctx, {"group_by", "aggregations"}, set())
        grupos = []
        for g in self.lista(d["group_by"], ctx + ".group_by"):
            gd = self.obj(g, ctx + ".group_by[]", {"column"}, {"period"})
            periodo = gd.get("period")
            if periodo is not None:
                self.opcion(periodo, ctx + ".period", _PERIODOS)
            grupos.append(GroupBy(self.texto(gd["column"], ctx + ".column"), periodo))
        aggs = []
        for a in self.lista(d["aggregations"], ctx + ".aggregations"):
            ad = self.obj(a, ctx + ".aggregations[]", {"column", "func"}, {"name"})
            aggs.append(
                Aggregation(
                    self.texto(ad["column"], ctx + ".column"),
                    self.opcion(ad["func"], ctx + ".func", _FUNCS),
                    self.texto(ad.get("name"), ctx + ".name", nulo=True),
                )
            )
        return AggregationSpec(grupos, aggs)

    def division(self, v: Any, ctx: str) -> SplitSpec | None:
        if v is None:
            return None
        d = self.obj(
            v,
            ctx,
            {"mode"},
            {"rows_per_part", "group_column", "value_columns", "name_pattern"},
        )
        return SplitSpec(
            mode=self.opcion(d["mode"], ctx + ".mode", _MODOS_DIVISION),
            rows_per_part=self.entero(d.get("rows_per_part"), ctx + ".rows_per_part"),
            group_column=self.texto(d.get("group_column"), ctx + ".group_column", nulo=True),
            value_columns=self.textos(d.get("value_columns"), ctx + ".value_columns", nulo=True),
            name_pattern=self.texto(d.get("name_pattern"), ctx + ".name_pattern", nulo=True),
        )

    def ruta(self, v: Any, ctx: str, base: Path) -> Path:
        texto = self.texto(v, ctx)
        if not texto.strip():
            raise _Invalida(f"{ctx}: ruta vacía")
        return _texto_a_ruta(texto, base)

    def salida(self, v: Any, ctx: str, base: Path) -> OutputSpec | None:
        if v is None:
            return None
        d = self.obj(
            v,
            ctx,
            {"path", "format", "destination"},
            {"split", "write_index", "sheet_name", "column_types"},
        )
        tipos = d.get("column_types")
        if tipos is not None:
            if not isinstance(tipos, dict):
                raise _Invalida(f"{ctx}.column_types: se esperaba un objeto")
            tipos = {
                self.texto(k, ctx): self.opcion(t, ctx + ".column_types", _TIPOS)
                for k, t in tipos.items()
            }
        return OutputSpec(
            path=self.ruta(d["path"], ctx + ".path", base),
            format=self.opcion(d["format"], ctx + ".format", _FORMATOS),
            destination=self.opcion(d["destination"], ctx + ".destination", _DESTINOS),
            split=self.division(d.get("split"), ctx + ".split"),
            write_index=self.booleano(d.get("write_index", True), ctx + ".write_index"),
            sheet_name=self.texto(d.get("sheet_name", "Datos"), ctx + ".sheet_name"),
            column_types=tipos,
        )

    def _campos_fuente(self, d: dict, ctx: str, base: Path) -> tuple:
        return (
            self.texto(d["name"], ctx + ".name"),
            self.ruta(d["path"], ctx + ".path", base),
            self.read(d.get("read_options", {}), ctx + ".read_options"),
            self.textos(d["columns"], ctx + ".columns"),
        )

    def fuente(self, v: Any, ctx: str, base: Path) -> SourceSpec:
        d = self.obj(v, ctx, {"name", "path", "columns"}, {"read_options"})
        return SourceSpec(*self._campos_fuente(d, ctx, base))

    def secundario(self, v: Any, ctx: str, base: Path) -> SecondarySpec:
        d = self.obj(
            v,
            ctx,
            {"name", "path", "columns", "key_pairs"},
            {"read_options", "how", "mode", "summary", "prefix"},
        )
        pares = []
        for p in self.lista(d["key_pairs"], ctx + ".key_pairs"):
            if not isinstance(p, list) or len(p) != 2:
                raise _Invalida(f"{ctx}.key_pairs: cada par debe tener 2 columnas")
            pares.append((self.texto(p[0], ctx), self.texto(p[1], ctx)))
        return SecondarySpec(
            *self._campos_fuente(d, ctx, base),
            key_pairs=pares,
            how=self.opcion(d.get("how", "left"), ctx + ".how", _HOW),
            mode=self.opcion(d.get("mode", "detail"), ctx + ".mode", _MODOS),
            summary=self.agregado(d.get("summary"), ctx + ".summary"),
            prefix=self.texto(d.get("prefix"), ctx + ".prefix", nulo=True),
        )

    def plan(self, v: Any, base: Path) -> JoinPlan:
        ctx = "plan"
        d = self.obj(
            v,
            ctx,
            {"principal", "secondaries"},
            {"key_options", "output_column_order", "final_summary", "output"},
        )
        k = self.obj(
            d.get("key_options", {}),
            "key_options",
            set(),
            {"trim", "ignore_case", "strip_leading_zeros"},
        )
        opciones = KeyOptions(
            trim=self.booleano(k.get("trim", True), "key_options.trim"),
            ignore_case=self.booleano(k.get("ignore_case", True), "key_options.ignore_case"),
            strip_leading_zeros=self.booleano(
                k.get("strip_leading_zeros", False), "key_options.strip_leading_zeros"
            ),
        )
        return JoinPlan(
            principal=self.fuente(d["principal"], "principal", base),
            key_options=opciones,
            secondaries=[
                self.secundario(s, f"secondaries[{i}]", base)
                for i, s in enumerate(self.lista(d["secondaries"], "secondaries"))
            ],
            output_column_order=self.textos(
                d.get("output_column_order"), "output_column_order", nulo=True
            ),
            final_summary=self.agregado(d.get("final_summary"), "final_summary"),
            output=self.salida(d.get("output"), "output", base),
        )


def load_recipe(path: Path) -> RecipeLoad:
    """Lee una receta. Los archivos que no existen NO impiden cargarla: quedan en `missing`.

    Lanza `ExcelToolError` si no se puede leer, está dañada o es de una versión desconocida.
    """
    path = Path(path)
    try:
        texto = path.read_text(encoding="utf-8-sig")
    except FileNotFoundError as e:
        raise ExcelToolError("No encontré la receta. Fijate si la moviste o la borraste.") from e
    except (OSError, UnicodeDecodeError) as e:
        log.exception("No se pudo leer la receta %s", path)
        raise ExcelToolError(_MSG_DANADA, str(e)) from e
    try:
        datos = json.loads(texto)
        if not isinstance(datos, dict) or "version" not in datos:
            raise _Invalida("falta la versión")
        version = datos["version"]
        if isinstance(version, bool) or version != VERSION:
            raise _Invalida(f"versión desconocida: {version!r}")
        advertencias: list[str] = []
        lector = _Lector(advertencias)
        extra = datos.keys() - {"version", "plan"}
        if extra:
            log.warning("Receta: campos desconocidos en la raíz, se ignoran: %s", sorted(extra))
            advertencias.append("La receta trae datos que esta versión no conoce.")
        if "plan" not in datos:
            raise _Invalida("falta el plan")
        plan = lector.plan(datos["plan"], path.parent)
    except (_Invalida, ValueError, RecursionError) as e:  # ValueError incluye JSONDecodeError
        log.error("Receta inválida %s: %s", path, e)
        raise ExcelToolError(_MSG_DANADA, str(e)) from e

    faltan: list[str] = []
    for fuente in [plan.principal, *plan.secondaries]:
        try:
            existe = fuente.path.is_file()
        except OSError:
            existe = False
        if not existe:
            faltan.append(fuente.name)
            advertencias.append(f"No encuentro el archivo «{fuente.path.name}» ({fuente.name}).")
    return RecipeLoad(plan, faltan, list(dict.fromkeys(advertencias)))
