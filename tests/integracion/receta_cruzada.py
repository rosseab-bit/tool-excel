"""Prueba de recetas entre Linux y Windows (no es un test; la corre el PM a mano).

    python tests/integracion/receta_cruzada.py crear CARPETA [--datos OTRA_CARPETA]
    python tests/integracion/receta_cruzada.py verificar CARPETA

`crear` copia los 3 archivos del escenario (clientes.xlsx, ventas.csv, cobranzas.xlsx) a CARPETA (o
a OTRA_CARPETA, por ejemplo en otra unidad: la receta guarda entonces rutas absolutas), guarda
`receta.json` en CARPETA, corre la unión y anota en `esperado.json` las filas, las columnas y una
huella del resultado. `verificar` carga la receta (que puede haberse hecho en el OTRO sistema),
exige que no falte ningún archivo, corre la unión y compara con `esperado.json`.
Termina con código 0 si todo coincide. Hay que correrla desde la carpeta del repo, con el entorno
(`.venv/bin/python` en Linux, `.venv\\Scripts\\python` en Windows).
"""

from __future__ import annotations

import hashlib
import json
import shutil
import sys
from pathlib import Path

from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy
from excel_tool.core.io import ReadOptions, read_table
from excel_tool.core.join import JoinPlan, SecondarySpec, SourceSpec, run_join
from excel_tool.core.normalize import KeyOptions
from excel_tool.core.output import OutputSpec
from excel_tool.core.recipe import load_recipe, save_recipe
from excel_tool.core.split import SplitSpec

ORIGEN = Path(__file__).resolve().parent.parent / "fixtures" / "generados" / "escenario_3"
ARCHIVOS = {"clientes": "clientes.xlsx", "ventas": "ventas.csv", "cobranzas": "cobranzas.xlsx"}


def _huella(df) -> str:
    h = hashlib.sha256()
    for fila in df.fillna("").itertuples(index=False):
        h.update("\x1f".join(str(x) for x in fila).encode("utf-8") + b"\n")
    return h.hexdigest()


def _correr(plan: JoinPlan) -> dict:
    ro = ReadOptions()
    tablas = {s.name: read_table(s.path, ro).df for s in [plan.principal, *plan.secondaries]}
    res = run_join(plan, tablas)
    return {
        "filas": len(res.df),
        "columnas": list(res.df.columns),
        "huella": _huella(res.df),
        "lleva_division": plan.output is not None and plan.output.split is not None,
    }


def crear(carpeta: Path, datos: Path | None) -> None:
    carpeta.mkdir(parents=True, exist_ok=True)
    datos = datos or carpeta
    datos.mkdir(parents=True, exist_ok=True)
    for nombre in ARCHIVOS.values():
        shutil.copy2(ORIGEN / nombre, datos / nombre)
    ro = ReadOptions()
    plan = JoinPlan(
        SourceSpec("clientes", datos / "clientes.xlsx", ro, ["ID cliente", "Nombre", "Región"]),
        KeyOptions(strip_leading_zeros=True),
        [
            SecondarySpec(
                "ventas",
                datos / "ventas.csv",
                ro,
                ["Fecha", "Artículo", "Importe"],
                [("ID cliente", "ID cliente")],
                mode="summary",
                summary=AggregationSpec(
                    [GroupBy("ID cliente")],
                    [Aggregation("Importe", "sum", "Total comprado"), Aggregation("", "count")],
                ),
            ),
            SecondarySpec(
                "cobranzas",
                datos / "cobranzas.xlsx",
                ro,
                ["Monto"],
                [("ID cliente", "Nro Cliente")],
                mode="summary",
                summary=AggregationSpec(
                    [GroupBy("Nro Cliente")], [Aggregation("Monto", "sum", "Total cobrado")]
                ),
            ),
        ],
        output=OutputSpec(
            carpeta / "salida" / "resultado.csv",
            "csv_excel",
            "excel",
            SplitSpec("rows", rows_per_part=12),
        ),
    )
    save_recipe(plan, carpeta / "receta.json")
    esperado = _correr(plan)
    (carpeta / "esperado.json").write_text(
        json.dumps(esperado, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"OK: receta en {carpeta / 'receta.json'} ({esperado['filas']} filas)")
    if "\\" in (carpeta / "receta.json").read_text(encoding="utf-8"):
        print("AVISO: la receta tiene barras invertidas; las rutas tienen que guardarse con /")


def verificar(carpeta: Path) -> int:
    carga = load_recipe(carpeta / "receta.json")
    if carga.missing:
        print(f"FALLA: faltan archivos: {carga.missing}\n{carga.warnings}")
        return 1
    esperado = json.loads((carpeta / "esperado.json").read_text(encoding="utf-8"))
    real = _correr(carga.plan)
    if real != esperado:
        print(f"FALLA: esperado {esperado}\nreal {real}")
        return 1
    division = carga.plan.output.split
    if division is None or division.mode != "rows" or division.rows_per_part != 12:
        print(f"FALLA: la división no se conservó: {division}")
        return 1
    print(f"OK: la receta de {carpeta} da el mismo resultado ({real['filas']} filas, con división)")
    return 0


def main(argv: list[str]) -> int:
    if len(argv) >= 3 and argv[1] == "crear":
        datos = Path(argv[argv.index("--datos") + 1]) if "--datos" in argv else None
        crear(Path(argv[2]), datos)
        return 0
    if len(argv) == 3 and argv[1] == "verificar":
        return verificar(Path(argv[2]))
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
