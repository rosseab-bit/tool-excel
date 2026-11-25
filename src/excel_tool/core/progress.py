"""Progreso y cancelación de operaciones largas. Contrato: docs/contratos.md § Fase 2.

El callback se llama desde el hilo de trabajo: acá no se sabe nada de tkinter.
"""

from __future__ import annotations

import threading
import time
from collections.abc import Callable
from dataclasses import dataclass

from excel_tool.core.errors import Cancelado


@dataclass(frozen=True)
class ProgressEvent:
    message: str  # castellano, para mostrar tal cual
    done: int  # pasos hechos
    total: int | None  # None = indeterminado


ProgressCallback = Callable[[ProgressEvent], None]


class CancelToken:
    """Bandera de cancelación, segura entre hilos."""

    def __init__(self) -> None:
        self._event = threading.Event()

    def cancel(self) -> None:
        self._event.set()

    @property
    def cancelled(self) -> bool:
        return self._event.is_set()

    def check(self) -> None:
        """Si se canceló, lanza `Cancelado`."""
        if self._event.is_set():
            raise Cancelado()


class ThrottledProgress:
    """Limita la frecuencia de un `ProgressCallback` (por defecto <= 10 eventos por segundo).

    Uso: `report = ThrottledProgress(progress)` y después `report(ProgressEvent(...))`.
    - Con `progress=None` no hace nada.
    - Se descarta un evento si el anterior salió hace menos de `min_interval` segundos,
      salvo el primero y el último (`done == total`), que siempre pasan.
    - `report.force(event)` emite sin mirar el tiempo (ej. al cambiar de columna).
    - `clock` se puede inyectar para los tests.
    """

    def __init__(
        self,
        progress: ProgressCallback | None,
        min_interval: float = 0.1,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self._progress = progress
        self._min_interval = min_interval
        self._clock = clock
        self._last: float | None = None

    def __call__(self, event: ProgressEvent) -> None:
        if self._progress is None:
            return
        now = self._clock()
        is_last = event.total is not None and event.done >= event.total
        if self._last is None or is_last or now - self._last >= self._min_interval:
            self._last = now
            self._progress(event)

    def force(self, event: ProgressEvent) -> None:
        if self._progress is None:
            return
        self._last = self._clock()
        self._progress(event)
