"""Medición de la Fase 3 (no es un test). Se corre con /usr/bin/time -v (pico de memoria):

    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase3.py save_csv
    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase3.py save_xlsx
    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase3.py save_valor_csv
    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase3.py split_file
    /usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase3.py split_file_valor

Todos parten del CSV de 2M de filas. `save_*` imitan a la UI: leer, perfilar (para los tipos) y
`save_result` en 7 partes de 300.000 filas (`save_valor_csv`: un archivo por Región, subdividiendo
a 1.000.000 de filas). `split_file*` no carga el archivo en memoria.
Las salidas van a una carpeta temporal que se borra al final. El presupuesto es 1,5 GB (ADR-023).
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import time
from pathlib import Path

from excel_tool.core.io import ReadOptions, read_table
from excel_tool.core.output import OutputSpec, save_result, split_file
from excel_tool.core.profile import profile_table
from excel_tool.core.split import SplitSpec

GRANDES = Path(__file__).resolve().parent.parent / "fixtures" / "generados" / "grandes"
ORIGEN = GRANDES / "rendimiento_2M.csv"


def _paso(nombre: str, t0: float) -> float:
    ahora = time.monotonic()
    print(f"  {nombre}: {ahora - t0:.1f} s", flush=True)
    return ahora


def main(escenario: str) -> None:
    salida = Path(tempfile.mkdtemp(prefix="medicion_fase3_"))
    try:
        t0 = time.monotonic()
        formato = "xlsx" if escenario == "save_xlsx" else "csv_excel"
        if escenario.startswith("split_file"):
            if escenario == "split_file":
                dividir = SplitSpec("rows", rows_per_part=300_000)
            else:
                dividir = SplitSpec("by_value", value_columns=["Región"], rows_per_part=1_000_000)
            spec = OutputSpec(salida / "partes.csv", "csv_excel", "excel", dividir)
            res = split_file(ORIGEN, ReadOptions(), spec)
            t = _paso("dividir el archivo", t0)
        else:
            leido = read_table(ORIGEN)
            t = _paso("leer", t0)
            perfil = profile_table(leido)
            t = _paso("perfilar", t)
            tipos = {c.name: c.detected_type for c in perfil.columns_profile}
            if escenario == "save_valor_csv":
                dividir = SplitSpec("by_value", value_columns=["Región"], rows_per_part=1_000_000)
            else:
                dividir = SplitSpec("rows", rows_per_part=300_000)
            spec = OutputSpec(
                salida / f"partes{'.xlsx' if formato == 'xlsx' else '.csv'}",
                formato,
                "excel",
                dividir,
                column_types=tipos,
            )
            res = save_result(leido.df, spec)
            t = _paso("guardar", t)
        print(
            f"  total: {time.monotonic() - t0:.1f} s; {res.rows_total} filas en "
            f"{len(res.parts)} partes; avisos: {len(res.warnings)}",
            flush=True,
        )
        assert res.rows_total == sum(p.rows for p in res.parts)
    finally:
        shutil.rmtree(salida, ignore_errors=True)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "save_csv")
