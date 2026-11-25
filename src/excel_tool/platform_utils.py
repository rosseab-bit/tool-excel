"""Todo lo que depende del sistema operativo vive acá."""

import ctypes
import logging
import os
import re
import subprocess
import sys
from pathlib import Path

from excel_tool import APP_DIR_NAME

log = logging.getLogger(__name__)

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


def app_dir() -> Path | None:
    """Carpeta del ejecutable si la app está compilada; None si corre desde el código fuente."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return None


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


class _MemoryStatusEx(ctypes.Structure):
    _fields_ = [  # noqa: RUF012
        ("dwLength", ctypes.c_uint32),
        ("dwMemoryLoad", ctypes.c_uint32),
        ("ullTotalPhys", ctypes.c_uint64),
        ("ullAvailPhys", ctypes.c_uint64),
        ("ullTotalPageFile", ctypes.c_uint64),
        ("ullAvailPageFile", ctypes.c_uint64),
        ("ullTotalVirtual", ctypes.c_uint64),
        ("ullAvailVirtual", ctypes.c_uint64),
        ("sullAvailExtendedVirtual", ctypes.c_uint64),
    ]


def _available_memory_windows() -> int:
    estado = _MemoryStatusEx()
    estado.dwLength = ctypes.sizeof(_MemoryStatusEx)
    if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(estado)):  # type: ignore[attr-defined]
        raise OSError("GlobalMemoryStatusEx falló")
    return int(estado.ullAvailPhys)


def _available_memory_linux() -> int:
    with Path("/proc/meminfo").open(encoding="utf-8") as f:
        for linea in f:
            if linea.startswith("MemAvailable:"):
                return int(linea.split()[1]) * 1024  # viene en kB
    raise OSError("no hay MemAvailable en /proc/meminfo")


def available_memory() -> int | None:
    """Bytes de memoria libre, o None si no se puede saber."""
    try:
        if sys.platform == "win32":
            return _available_memory_windows()
        if sys.platform.startswith("linux"):
            return _available_memory_linux()
        log.info("Memoria libre: sistema no soportado (%s)", sys.platform)
    except (OSError, ValueError, AttributeError, IndexError) as exc:
        log.info("No se pudo leer la memoria libre: %s", exc)
    return None
