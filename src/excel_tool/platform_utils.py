"""Todo lo que depende del sistema operativo vive acá."""

import os
import re
import subprocess
import sys
from pathlib import Path

from excel_tool import APP_DIR_NAME

_PROHIBIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f\x7f]')
_RESERVADOS = {"CON", "PRN", "AUX", "NUL"}
_RESERVADOS |= {f"COM{i}" for i in range(1, 10)} | {f"LPT{i}" for i in range(1, 10)}
_MAX_LARGO = 150


def user_data_dir() -> Path:
    """Carpeta de datos de la app (se crea si no existe)."""
    if sys.platform == "win32":
        base_env = os.environ.get("LOCALAPPDATA")
        base = Path(base_env) if base_env else Path.home() / "AppData" / "Local"
    else:
        xdg = os.environ.get("XDG_DATA_HOME")
        base = Path(xdg) if xdg else Path.home() / ".local" / "share"
    carpeta = base / APP_DIR_NAME
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


def log_dir() -> Path:
    """Carpeta de logs (se crea si no existe)."""
    carpeta = user_data_dir() / "logs"
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta


def open_folder(path: Path) -> None:
    """Abre la carpeta en el explorador de archivos del sistema."""
    try:
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]  # noqa: S606
        else:
            subprocess.Popen(  # noqa: S603
                ["xdg-open", str(path)],
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
    except OSError as exc:
        raise OSError(f"No se pudo abrir la carpeta: {path}") from exc


def safe_filename(name: str, replacement: str = "_") -> str:
    """Nombre de archivo válido en Linux y Windows."""
    limpio = _PROHIBIDOS.sub(replacement, name).rstrip(" .")
    if not limpio:
        return "sin_nombre"
    base, punto, ext = limpio.partition(".")
    if base.upper() in _RESERVADOS:
        limpio = base + replacement + punto + ext
    if len(limpio) > _MAX_LARGO:
        stem, punto, ext = limpio.rpartition(".")
        if punto and 0 < len(ext) < _MAX_LARGO:
            limpio = stem[: _MAX_LARGO - len(ext) - 1] + "." + ext
        else:
            limpio = limpio[:_MAX_LARGO]
        limpio = limpio.rstrip(" .") or "sin_nombre"
    return limpio
