"""Dividir una tabla en partes (por cantidad de filas, por grupos, por valor o en hojas).

Decisiones de diseño:
- Se trabaja con posiciones (índices) y nunca se copia la tabla entera: `plan_split` y
  `split_dataframe` comparten `_planificar`, así que dan exactamente las mismas partes.
- `split_dataframe` arma cada parte recién cuando se la pide (iterador).
- Tabla con 0 filas: 0 partes (no es un error). Quien guarda decide qué hacer con eso.
- `rows_keep_groups`: se REORDENA por grupo, manteniendo el orden de la primera aparición de cada
  grupo y el orden original de las filas dentro del grupo. Los grupos se van juntando mientras
  entren en `rows_per_part`; un grupo más grande que eso va entero en su propia parte (con aviso
  en el log). En ese modo `first_row` / `last_row` son la menor y la mayor fila del original
  que contiene la parte (no es un rango continuo).
- `by_value`: partes ordenadas por valor (alfabético sin distinguir mayúsculas, vacío al final).
  `PartPlan.value` conserva el valor tal cual ("" para el vacío); en el nombre se usa "vacio".
  Si una parte supera el límite se subdivide en tandas de a `límite` filas.
- Límite efectivo = el menor entre `rows_per_part` y `max_rows_per_part` (los que estén).
"""

from __future__ import annotations

import logging
from collections.abc import Iterator
from dataclasses import dataclass
from typing import Literal

import numpy as np
import pandas as pd

from excel_tool.core.errors import ExcelToolError
from excel_tool.core.io.writer import nombre_de_hoja
from excel_tool.platform_utils import safe_filename

logger = logging.getLogger(__name__)

_PATRON_FILAS = "{nombre}_parte_{n:02d}_de_{total:02d}"
_PATRON_VALOR = "{nombre}_{valor}"
_MAX_VALOR = 60  # largo máximo del valor dentro del nombre
_MAX_HOJA = 31
_VACIO = "vacio"


@dataclass(frozen=True)
class SplitSpec:
    mode: Literal["rows", "rows_keep_groups", "by_value", "sheets"]
    rows_per_part: int | None = None
    group_column: str | None = None
    value_columns: list[str] | None = None
    name_pattern: str | None = None


@dataclass(frozen=True)
class PartPlan:
    name: str
    rows: int
    value: tuple[str, ...] | None
    first_row: int
    last_row: int


# Selector de una parte: rango (inicio, fin) de filas del original o posiciones sueltas.
_Seleccion = tuple[int, int] | np.ndarray


def _limite(spec: SplitSpec, max_rows: int | None) -> int:
    candidatos = [v for v in (spec.rows_per_part, max_rows) if v is not None]
    if not candidatos or min(candidatos) < 1:
        raise ExcelToolError(
            "Indicá cuántas filas tiene que tener cada parte (un número mayor que cero)."
        )
    return int(min(candidatos))


def _columnas(df: pd.DataFrame, nombres: list[str] | None, que: str) -> list[str]:
    if not nombres:
        raise ExcelToolError(f"Elegí {que}.")
    faltan = [c for c in nombres if c not in df.columns]
    if faltan:
        raise ExcelToolError(
            "No encuentro estas columnas en la tabla: " + ", ".join(faltan) + ".", str(faltan)
        )
    return list(nombres)


def _serie(df: pd.DataFrame, columna: str) -> pd.Series:
    s = df[columna]
    if isinstance(s, pd.DataFrame):  # columna repetida: se usa la primera
        s = s.iloc[:, 0]
    return s.fillna("")


# ----------------------------------------------------------------------------- nombres


class _Nombres:
    """Arma nombres válidos y sin repetir (sin distinguir mayúsculas, como Windows)."""

    def __init__(self, base: str, spec: SplitSpec, hojas: bool):
        self.base = base
        self.spec = spec
        self.hojas = hojas
        self.usados: set[str] = set()

    def _formatear(self, patron: str, **datos) -> str:
        try:
            return patron.format(**datos)
        except (KeyError, IndexError, ValueError) as e:
            raise ExcelToolError(
                "El modelo de nombre de las partes no es válido. Podés usar {nombre}, {n}, "
                "{total} y {valor}.",
                repr(e),
            ) from e

    def nombre(
        self,
        n: int,
        total: int,
        valor: tuple[str, ...] | None = None,
        sub: tuple[int, int] | None = None,
    ) -> str:
        base = self.base
        texto_valor = ""
        if valor is not None:
            texto_valor = " - ".join((v if v != "" else _VACIO)[:_MAX_VALOR] for v in valor)
        if valor is None:
            patron = self.spec.name_pattern or _PATRON_FILAS
            crudo = self._formatear(patron, nombre=base, n=n, total=total, valor="")
            if self.hojas and len(crudo) > _MAX_HOJA and not self.spec.name_pattern:
                sobra = len(crudo) - _MAX_HOJA
                crudo = self._formatear(
                    patron, nombre=base[: max(1, len(base) - sobra)], n=n, total=total, valor=""
                )
        else:
            patron = self.spec.name_pattern or _PATRON_VALOR
            crudo = self._formatear(patron, nombre=base, n=n, total=total, valor=texto_valor)
            if sub is not None and not (self.spec.name_pattern and "{n" in self.spec.name_pattern):
                crudo += f"_parte_{sub[0]}_de_{sub[1]}"
        return self._unico(crudo)

    def _unico(self, crudo: str) -> str:
        if self.hojas:
            return nombre_de_hoja(crudo, self.usados)
        limpio = safe_filename(crudo)
        final, k = limpio, 2
        while final.lower() in self.usados:
            sufijo = f" ({k})"
            final = limpio[: 150 - len(sufijo)] + sufijo
            k += 1
        self.usados.add(final.lower())
        return final


# ------------------------------------------------------------------------------ modos


def _por_filas(n: int, limite: int, nombres: _Nombres) -> list[tuple[PartPlan, _Seleccion]]:
    total = -(-n // limite)
    partes = []
    for i in range(total):
        ini, fin = i * limite, min((i + 1) * limite, n)
        plan = PartPlan(nombres.nombre(i + 1, total), fin - ini, None, ini + 1, fin)
        partes.append((plan, (ini, fin)))
    return partes


def _por_grupos(
    df: pd.DataFrame, spec: SplitSpec, limite: int, nombres: _Nombres
) -> list[tuple[PartPlan, _Seleccion]]:
    (col,) = _columnas(
        df, [spec.group_column] if spec.group_column else None, "la columna de grupos"
    )
    codigos, _ = pd.factorize(_serie(df, col), sort=False)  # códigos en orden de primera aparición
    orden = np.argsort(codigos, kind="stable")
    cuentas = np.bincount(codigos)
    acum = np.concatenate(([0], np.cumsum(cuentas)))  # filas acumuladas antes de cada grupo
    n_grupos = len(cuentas)

    rangos: list[tuple[int, int]] = []
    g = 0
    while g < n_grupos:
        # último grupo g2 tal que las filas de g..g2 entren en el límite
        g2 = int(np.searchsorted(acum, acum[g] + limite, side="right")) - 1
        if g2 <= g:
            g2 = g + 1
            logger.warning(
                "Un grupo de la columna %r tiene %d filas y supera las %d por parte: "
                "va entero en su propia parte.",
                col,
                int(cuentas[g]),
                limite,
            )
        rangos.append((int(acum[g]), int(acum[g2])))
        g = g2
    partes = []
    total = len(rangos)
    for i, (ini, fin) in enumerate(rangos):
        pos = orden[ini:fin]
        plan = PartPlan(
            nombres.nombre(i + 1, total), fin - ini, None, int(pos.min()) + 1, int(pos.max()) + 1
        )
        partes.append((plan, pos))
    return partes


def _rangos_por_valor(df: pd.DataFrame, columnas: list[str]):
    """(orden, límites de grupo, valores de cada grupo). Orden estable por valor."""
    n = len(df)
    ranks: list[np.ndarray] = []
    uniques: list[np.ndarray] = []
    for c in columnas:
        cod, unicos = pd.factorize(_serie(df, c), sort=False)
        u = [str(x) for x in unicos]
        clave = sorted(range(len(u)), key=lambda i: (u[i] == "", u[i].casefold(), u[i]))
        rango = np.empty(len(u), dtype=np.int64)
        rango[clave] = np.arange(len(u))
        ranks.append(rango[cod])
        uniques.append(np.array(u, dtype=object)[np.array(clave, dtype=np.int64)])
    if len(columnas) == 1:
        orden = np.argsort(ranks[0], kind="stable")
    else:
        orden = np.lexsort(tuple(reversed(ranks)))
    cambio = np.zeros(n, dtype=bool)
    for r in ranks:
        s = r[orden]
        cambio[1:] |= s[1:] != s[:-1]
    cambio[:1] = True
    inicios = np.flatnonzero(cambio)
    limites = np.concatenate((inicios, [n]))
    valores = [
        tuple(str(uniques[k][ranks[k][orden[ini]]]) for k in range(len(columnas)))
        for ini in inicios
    ]
    return orden, limites, valores


def _por_valor(
    df: pd.DataFrame, spec: SplitSpec, limite: int | None, nombres: _Nombres
) -> list[tuple[PartPlan, _Seleccion]]:
    columnas = _columnas(df, spec.value_columns, "la columna por la que dividir")
    orden, limites, valores = _rangos_por_valor(df, columnas)
    total = len(valores)
    partes = []
    for i, valor in enumerate(valores):
        ini, fin = int(limites[i]), int(limites[i + 1])
        tandas = 1 if limite is None else -(-(fin - ini) // limite)
        for t in range(tandas):
            a = ini + t * (limite or 0)
            b = fin if tandas == 1 else min(a + limite, fin)  # type: ignore[operator]
            pos = orden[a:b]
            sub = (t + 1, tandas) if tandas > 1 else None
            plan = PartPlan(
                nombres.nombre(i + 1, total, valor, sub),
                b - a,
                valor,
                int(pos[0]) + 1,  # orden estable: la primera es la menor
                int(pos.max()) + 1,
            )
            partes.append((plan, pos))
    return partes


def _planificar(
    df: pd.DataFrame, spec: SplitSpec, base_name: str, max_rows: int | None
) -> list[tuple[PartPlan, _Seleccion]]:
    n = len(df)
    if spec.mode not in ("rows", "rows_keep_groups", "by_value", "sheets"):
        raise ExcelToolError("No conozco esa forma de dividir.", str(spec.mode))
    nombres = _Nombres(base_name, spec, hojas=spec.mode == "sheets")
    if spec.mode == "by_value":
        candidatos = [v for v in (spec.rows_per_part, max_rows) if v is not None]
        limite = int(min(candidatos)) if candidatos else None
        if limite is not None and limite < 1:
            raise ExcelToolError("El máximo de filas por parte tiene que ser mayor que cero.")
        if n == 0:
            return []
        return _por_valor(df, spec, limite, nombres)
    limite = _limite(spec, max_rows)
    if n == 0:
        return []
    if spec.mode == "rows_keep_groups":
        return _por_grupos(df, spec, limite, nombres)
    return _por_filas(n, limite, nombres)


def plan_split(
    df: pd.DataFrame, spec: SplitSpec, *, base_name: str, max_rows_per_part: int | None = None
) -> list[PartPlan]:
    """Cómo quedaría dividida la tabla, sin armar los datos."""
    return [plan for plan, _ in _planificar(df, spec, base_name, max_rows_per_part)]


def split_dataframe(
    df: pd.DataFrame, spec: SplitSpec, *, base_name: str, max_rows_per_part: int | None = None
) -> Iterator[tuple[PartPlan, pd.DataFrame]]:
    """Las mismas partes que `plan_split`, una por una (cada parte se arma al pedirla)."""
    for plan, sel in _planificar(df, spec, base_name, max_rows_per_part):
        if isinstance(sel, tuple):
            parte = df.iloc[sel[0] : sel[1]]
        else:
            parte = df.iloc[sel]
        yield plan, parte.reset_index(drop=True)
