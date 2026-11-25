"""Preferencias del usuario (JSON en la carpeta de datos). Contrato: docs/contratos.md § Fase 2."""

from __future__ import annotations

import contextlib
import json
import logging
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path

from excel_tool import platform_utils
from excel_tool.core.io import CsvOptions

log = logging.getLogger(__name__)

_FILE_NAME = "preferencias.json"


@dataclass
class Preferences:
    last_folder: Path | None = None
    last_csv_options: CsvOptions | None = None
    last_destination: str | None = None  # se usa en la Fase 3


def _path() -> Path:
    return platform_utils.user_data_dir() / _FILE_NAME


def load_preferences() -> Preferences:
    """Lee las preferencias. Si el archivo falta o está dañado, devuelve los valores por defecto."""
    try:
        path = _path()
        if not path.exists():
            return Preferences()
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("el JSON no es un objeto")
        folder = data.get("last_folder")
        csv_raw = data.get("last_csv_options")
        csv = None
        if isinstance(csv_raw, dict):
            csv = CsvOptions(separator=str(csv_raw["separator"]), encoding=str(csv_raw["encoding"]))
        dest = data.get("last_destination")
        return Preferences(
            last_folder=Path(folder) if isinstance(folder, str) and folder else None,
            last_csv_options=csv,
            last_destination=dest if isinstance(dest, str) else None,
        )
    except Exception:
        log.warning(
            "No se pudieron leer las preferencias; se usan los valores por defecto.", exc_info=True
        )
        return Preferences()


def save_preferences(prefs: Preferences) -> None:
    """Guarda las preferencias de forma atómica (temporal + reemplazo). Nunca rompe la app."""
    data = {
        "last_folder": str(prefs.last_folder) if prefs.last_folder else None,
        "last_csv_options": (
            {
                "separator": prefs.last_csv_options.separator,
                "encoding": prefs.last_csv_options.encoding,
            }
            if prefs.last_csv_options
            else None
        ),
        "last_destination": prefs.last_destination,
    }
    tmp_name: str | None = None
    try:
        path = _path()
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp_name = tempfile.mkstemp(dir=path.parent, prefix="preferencias-", suffix=".tmp")
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        os.replace(tmp_name, path)
        tmp_name = None
    except Exception:
        log.warning("No se pudieron guardar las preferencias.", exc_info=True)
    finally:
        if tmp_name:
            with contextlib.suppress(OSError):
                os.unlink(tmp_name)
