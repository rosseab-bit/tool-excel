"""Operaciones largas en segundo plano con ventana de progreso (docs/contratos.md § Fase 2).

La función corre en un hilo; el progreso y el resultado viajan por una `queue.Queue` y el hilo de
tkinter los lee con `after()`. Ningún widget se toca desde el hilo de trabajo.
"""

from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
from collections.abc import Callable
from tkinter import messagebox, ttk

from excel_tool.core.errors import Cancelado, ExcelToolError
from excel_tool.core.progress import CancelToken, ProgressCallback, ProgressEvent

log = logging.getLogger(__name__)

INTERVALO_MS = 100
RETARDO_VENTANA_S = 0.3
MENSAJE_INESPERADO = "Ocurrió un problema inesperado. El detalle quedó guardado en el registro."


class _Tarea:
    def __init__(
        self,
        root: tk.Misc,
        func: Callable[[ProgressCallback, CancelToken], object],
        title: str,
        on_success: Callable[[object], None],
        on_error: Callable[[BaseException], None] | None,
        cancellable: bool,
        modal: bool,
        on_cancel: Callable[[], None] | None = None,
    ) -> None:
        self.root = root
        self.on_cancel = on_cancel
        self.cancelada = False
        self.title = title
        self.on_success = on_success
        self.on_error = on_error
        self.cancellable = cancellable
        self.modal = modal
        self.cancel = CancelToken()
        self.cola: queue.Queue[tuple[str, object]] = queue.Queue()
        self.ventana: tk.Toplevel | None = None
        self.etiqueta: ttk.Label | None = None
        self.barra: ttk.Progressbar | None = None
        self.boton: ttk.Button | None = None
        self.determinada: bool | None = None
        self.ultimo_evento: ProgressEvent | None = None
        self.terminada = False
        self.abandonada = False
        self.transcurrido_ms = 0
        self._func = func
        self._after_id: str | None = None

        try:
            root.bind("<Destroy>", self._al_destruir_raiz, add="+")
        except tk.TclError:
            pass
        self.hilo = threading.Thread(target=self._correr, name="worker", daemon=True)
        self.hilo.start()
        self._programar()

    # --- hilo de trabajo (sin tocar widgets) ---
    def _correr(self) -> None:
        def progress(event: ProgressEvent) -> None:
            self.cola.put(("progreso", event))

        try:
            resultado = self._func(progress, self.cancel)
        except BaseException as exc:  # noqa: BLE001 - se informa en el hilo de tkinter
            if self.cancelada:
                log.info("'%s' terminó con error después de cancelar: %r", self.title, exc)
            self.cola.put(("error", exc))
        else:
            self.cola.put(("ok", resultado))

    # --- hilo de tkinter ---
    def _programar(self) -> None:
        try:
            self._after_id = self.root.after(INTERVALO_MS, self._sondear)
        except tk.TclError:
            self._abandonar()

    def _al_destruir_raiz(self, event: tk.Event) -> None:
        if event.widget is self.root:
            self._abandonar()

    def _abandonar(self) -> None:
        """La ventana principal se cerró: cancelar y no llamar a nada más."""
        self.abandonada = True
        self.cancel.cancel()

    def _sondear(self) -> None:
        if self.abandonada or self.cancelada:
            return
        self.transcurrido_ms += INTERVALO_MS
        final: tuple[str, object] | None = None
        try:
            while True:
                tipo, dato = self.cola.get_nowait()
                if tipo == "progreso":
                    self.ultimo_evento = dato  # type: ignore[assignment]
                else:
                    final = (tipo, dato)
                    break
        except queue.Empty:
            pass

        if final is not None:
            self.terminada = True
            self._cerrar_ventana()
            self._terminar(*final)
            return

        if self.ventana is None and self.transcurrido_ms >= RETARDO_VENTANA_S * 1000:
            self._crear_ventana()
        if self.ventana is not None:
            self._actualizar_ventana()
        self._programar()

    def _crear_ventana(self) -> None:
        win = tk.Toplevel(self.root)
        win.title(self.title)
        win.resizable(False, False)
        marco = ttk.Frame(win, padding=16)
        marco.pack(fill="both", expand=True)
        self.etiqueta = ttk.Label(marco, text=self.title, width=50, anchor="w")
        self.etiqueta.pack(fill="x")
        self.barra = ttk.Progressbar(marco, length=380, mode="indeterminate")
        self.barra.pack(fill="x", pady=(10, 10))
        if self.cancellable:
            self.boton = ttk.Button(marco, text="Cancelar", command=self._pedir_cancelar)
            self.boton.pack(anchor="e")
        win.protocol("WM_DELETE_WINDOW", self._pedir_cancelar if self.cancellable else _nada)
        self.ventana = win
        if self.modal:
            try:
                win.transient(self.root.winfo_toplevel())  # type: ignore[arg-type]
                win.grab_set()
            except tk.TclError:
                log.debug("No se pudo tomar el foco exclusivo de la ventana de progreso")

    def _actualizar_ventana(self) -> None:
        if self.etiqueta is None or self.barra is None:
            return
        ev = self.ultimo_evento
        if ev is not None:
            self.etiqueta.configure(text=ev.message)
        total = ev.total if ev is not None else None
        if total is not None:
            if self.determinada is not True:
                self.barra.stop()
                self.barra.configure(mode="determinate")
                self.determinada = True
            self.barra.configure(maximum=max(total, 1), value=ev.done if ev else 0)
        else:
            if self.determinada is not False:
                self.barra.configure(mode="indeterminate")
                self.barra.start(30)
                self.determinada = False

    def _pedir_cancelar(self) -> None:
        """Cancelar responde al instante: se cierra la ventana y se descarta el resultado."""
        if self.cancelada or self.terminada:
            return
        self.cancelada = True
        self.cancel.cancel()
        if self._after_id is not None:
            try:
                self.root.after_cancel(self._after_id)
            except tk.TclError:
                pass
        self._cerrar_ventana()
        self._avisar_cancelada()

    def _avisar_cancelada(self) -> None:
        log.info("Operación cancelada por el usuario: %s", self.title)
        callback, self.on_cancel = self.on_cancel, None
        if callback is None:
            return
        try:
            callback()
        except Exception:
            log.exception("Falló on_cancel de la operación '%s'", self.title)

    def _cerrar_ventana(self) -> None:
        win, self.ventana = self.ventana, None
        if win is None:
            return
        try:
            if self.barra is not None:
                self.barra.stop()
            win.grab_release()
        except tk.TclError:
            pass
        try:
            win.destroy()
        except tk.TclError:
            pass

    def _terminar(self, tipo: str, dato: object) -> None:
        try:
            if tipo == "ok":
                self.on_success(dato)
            elif isinstance(dato, Cancelado):
                self.cancelada = True
                self._avisar_cancelada()
            elif isinstance(dato, BaseException):
                self._error(dato)
        except Exception:
            log.exception("Falló el callback de la operación '%s'", self.title)
            if tipo == "ok":
                self._mostrar(MENSAJE_INESPERADO)

    def _error(self, exc: BaseException) -> None:
        if isinstance(exc, ExcelToolError):
            log.warning("%s: %s | %s", self.title, exc.mensaje, exc.detalle or "")
            mensaje = exc.mensaje
        else:
            log.error(
                "Error inesperado en '%s'", self.title, exc_info=(type(exc), exc, exc.__traceback__)
            )
            mensaje = MENSAJE_INESPERADO
        if self.on_error is not None:
            self.on_error(exc)
        else:
            self._mostrar(mensaje)

    def _mostrar(self, mensaje: str) -> None:
        try:
            messagebox.showerror(self.title, mensaje, parent=self.root)
        except tk.TclError:
            log.exception("No se pudo mostrar el mensaje de error")


def _nada() -> None:
    return None


def run_in_background[T](
    root: tk.Misc,
    func: Callable[[ProgressCallback, CancelToken], T],
    *,
    title: str,
    on_success: Callable[[T], None],
    on_error: Callable[[BaseException], None] | None = None,
    cancellable: bool = True,
    modal: bool = True,
    on_cancel: Callable[[], None] | None = None,
) -> CancelToken:
    """Corre `func(progress, cancel)` en un hilo y devuelve el `CancelToken`.

    `on_success`, `on_error` y `on_cancel` se llaman siempre en el hilo de tkinter. Al cancelar
    desde la ventana se cierra enseguida y no se llama a `on_success` ni `on_error`, aunque `func`
    siga corriendo; `on_cancel` se llama una sola vez.
    """
    tarea = _Tarea(
        root,
        func,
        title,
        on_success,  # type: ignore[arg-type]
        on_error,
        cancellable,
        modal,
        on_cancel,
    )
    return tarea.cancel
