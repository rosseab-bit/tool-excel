"""Control de memoria antes de cargar un archivo (ADR-025)."""

import logging
from dataclasses import dataclass
from pathlib import Path

from excel_tool import platform_utils

log = logging.getLogger(__name__)

# Medido en F2-05 (ADR-028): base fija de la app con pandas y pyarrow cargados.
BASE_BYTES = 114 * 1024 * 1024
# Factores medidos en F2-05 con margen para el análisis (neto: CSV 4,4; xlsx 21; xls 16).
FACTOR_CSV = 6
FACTOR_XLSX = 25  # también xlsm, xlsb (sin medir, se usa el de xlsx) y ods
FACTOR_XLS = 20

_FACTORES = {
    "csv": FACTOR_CSV,
    "txt": FACTOR_CSV,
    "xlsx": FACTOR_XLSX,
    "xlsm": FACTOR_XLSX,
    "xlsb": FACTOR_XLSX,
    "ods": FACTOR_XLSX,
    "xls": FACTOR_XLS,
}
_GB = 1024**3
_MB = 1024**2


@dataclass(frozen=True)
class MemoryCheck:
    needed: int
    available: int | None
    ok: bool
    message: str


def _monto(n: int) -> str:
    """Monto en GB o MB con coma decimal."""
    if n >= _GB:
        texto = f"{n / _GB:.1f} GB"
    else:
        texto = f"{n / _MB:.1f} MB"
    return texto.replace(".", ",")


def check_memory(path: Path, *, format: str | None = None, margin: float = 0.8) -> MemoryCheck:
    """Estima la memoria necesaria para cargar el archivo y la compara con la libre."""
    path = Path(path)
    ext = (format or path.suffix).lower().lstrip(".")
    factor = _FACTORES.get(ext, FACTOR_XLSX)
    needed = int(BASE_BYTES + path.stat().st_size * factor)
    available = platform_utils.available_memory()
    if available is None:
        log.info("Memoria libre desconocida; se deja cargar %s", path.name)
        return MemoryCheck(needed, None, True, "")
    ok = needed <= available * margin
    message = ""
    if not ok:
        message = (
            "Este archivo es muy grande para la memoria libre de esta computadora: "
            f"necesita alrededor de {_monto(needed)} y hay {_monto(available)} libres. "
            "Cerrá otros programas y probá de nuevo."
        )
        log.warning("Memoria insuficiente para %s: %s", path.name, message)
    return MemoryCheck(needed, available, ok, message)


# Bytes por celda de un resultado ya armado (texto pyarrow: ~10 caracteres + offsets + el
# vector auxiliar de la unión). Medido en F4-03 con el escenario grande, con margen.
BYTES_POR_CELDA_RESULTADO = 32


def check_result_memory(
    rows: int,
    columns: int,
    *,
    extra_bytes: int = 0,
    bytes_per_cell: int = BYTES_POR_CELDA_RESULTADO,
    margin: float = 0.8,
) -> MemoryCheck:
    """Estima la memoria de un resultado de `rows` x `columns` y la compara con la libre.

    No mira archivos: los datos de entrada ya están cargados. `extra_bytes` suma lo que se
    necesita aparte de las celdas (por ejemplo, los vectores de posiciones de la unión).
    """
    needed = int(rows) * int(columns) * bytes_per_cell + int(extra_bytes)
    available = platform_utils.available_memory()
    if available is None:
        log.info("Memoria libre desconocida; se deja armar el resultado")
        return MemoryCheck(needed, None, True, "")
    ok = needed <= available * margin
    message = ""
    if not ok:
        message = (
            "El resultado es muy grande para la memoria libre de esta computadora: "
            f"necesita alrededor de {_monto(needed)} y hay {_monto(available)} libres. "
            "Probá con menos columnas, resumí alguno de los archivos o cerrá otros programas."
        )
        log.warning("Memoria insuficiente para el resultado: %s", message)
    return MemoryCheck(needed, available, ok, message)
