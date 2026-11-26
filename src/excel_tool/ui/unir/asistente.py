"""Ventana del asistente para unir archivos: barra de pasos, cuerpo con scroll y pie fijo."""

from __future__ import annotations

import logging
import tkinter as tk
from tkinter import messagebox, ttk

from excel_tool.ui.desplazable import CuerpoDesplazable
from excel_tool.ui.unir.estado import EstadoUnion
from excel_tool.ui.unir.pasos import PASOS, Paso

log = logging.getLogger(__name__)

ALTO_MAXIMO_VENTANA = 768
ALTO_CUERPO = 420  # el pie y la barra de pasos quedan siempre a la vista
ANCHO_VENTANA = 900
TEXTO_CANCELAR = "Si cancelás, se pierde lo que cargaste. ¿Querés salir?"


class Asistente(tk.Toplevel):
    def __init__(
        self, parent: tk.Misc, estado: EstadoUnion | None = None, paso_inicial: int = 0
    ) -> None:
        super().__init__(parent)
        self.estado = estado if estado is not None else EstadoUnion()
        self.cerrada = False
        self.guardado = False
        self.indice = -1
        self.completados: set[int] = set()
        self.title("Unir archivos")
        self.transient(parent.winfo_toplevel())  # type: ignore[arg-type]
        self.protocol("WM_DELETE_WINDOW", self.cancelar)

        self._barra = ttk.Frame(self, padding=(12, 10, 12, 6))
        self._barra.pack(side="top", fill="x")
        self._etiquetas: list[ttk.Label] = []
        self.cuerpo = CuerpoDesplazable(self, alto_maximo=ALTO_CUERPO)
        self._armar_pie()
        self.pasos: list[Paso] = [clase(self.cuerpo.marco, self.estado, self) for clase in PASOS]
        self._armar_barra()
        self.ir_a(max(0, min(paso_inicial, len(self.pasos) - 1)))
        self._ubicar()

    # ---- armado ----
    def _armar_barra(self) -> None:
        for i, paso in enumerate(self.pasos):
            self._barra.columnconfigure(i, weight=1, uniform="pasos")
            etiqueta = ttk.Label(self._barra, anchor="center", justify="center", wraplength=110)
            etiqueta.grid(row=0, column=i, sticky="ew", padx=2)
            etiqueta.paso_titulo = paso.titulo  # type: ignore[attr-defined]
            self._etiquetas.append(etiqueta)
        ttk.Separator(self).pack(side="top", fill="x", before=self.cuerpo.lienzo.master)

    def _armar_pie(self) -> None:
        pie = self.cuerpo.pie
        self.estado_texto = ttk.Label(pie, text="", wraplength=ANCHO_VENTANA - 360)
        self.estado_texto.pack(side="left", fill="x", expand=True)
        self.boton_cancelar = ttk.Button(pie, text="Cancelar", command=self.cancelar)
        self.boton_cancelar.pack(side="right", padx=(6, 0))
        self.boton_siguiente = ttk.Button(pie, text="Siguiente", command=self.siguiente)
        self.boton_siguiente.pack(side="right", padx=(6, 0))
        self.boton_anterior = ttk.Button(pie, text="Anterior", command=self.anterior)
        self.boton_anterior.pack(side="right")

    def _ubicar(self) -> None:
        self.update_idletasks()
        ancho = min(max(self.winfo_reqwidth(), ANCHO_VENTANA), self.winfo_screenwidth())
        alto = min(self.winfo_reqheight(), ALTO_MAXIMO_VENTANA, self.winfo_screenheight())
        x = max((self.winfo_screenwidth() - ancho) // 2, 0)
        y = max((self.winfo_screenheight() - alto) // 2, 0)
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.minsize(min(ancho, 640), min(alto, 400))

    # ---- navegación ----
    @property
    def paso_actual(self) -> Paso:
        return self.pasos[self.indice]

    def avisar(self, texto: str) -> None:
        """Muestra un texto en el pie (los pasos lo usan para sus mensajes)."""
        self.estado_texto.configure(text=texto)

    def marcar_guardado(self, guardado: bool) -> None:
        """Después de guardar, "Cancelar" pasa a ser "Terminar" (cierra sin preguntar)."""
        self.guardado = guardado
        self.boton_cancelar.configure(text="Terminar" if guardado else "Cancelar")

    def _mostrar(self, indice: int) -> None:
        if 0 <= self.indice < len(self.pasos):
            self.pasos[self.indice].pack_forget()
        self.indice = indice
        self.pasos[indice].pack(fill="both", expand=True)
        self.avisar("")
        self.boton_anterior.state(["!disabled"] if indice > 0 else ["disabled"])
        ultimo = indice == len(self.pasos) - 1
        if ultimo:
            self.boton_siguiente.pack_forget()  # en el último paso no hay "Siguiente"
        elif not self.boton_siguiente.winfo_manager():
            self.boton_siguiente.pack(side="right", padx=(6, 0), before=self.boton_anterior)
        self.marcar_guardado(False)
        self._refrescar_barra()

    def _refrescar_barra(self) -> None:
        total = len(self.pasos)
        for i, etiqueta in enumerate(self._etiquetas):
            base = f"{i + 1}. {etiqueta.paso_titulo}"  # type: ignore[attr-defined]
            if i == self.indice:
                etiqueta.configure(text=base, font="TkHeadingFont", foreground="")
            elif i in self.completados:
                etiqueta.configure(text=f"✓ {base}", font="TkDefaultFont", foreground="#2e7d32")
            else:
                etiqueta.configure(text=base, font="TkDefaultFont", foreground="#666666")
        self.title(f"Unir archivos - paso {self.indice + 1} de {total}")

    def ir_a(self, indice: int) -> None:
        """Salta a un paso (sin validar). Los anteriores quedan como completados."""
        self.completados = set(range(indice))
        self._mostrar(indice)
        self.paso_actual.al_entrar()

    def siguiente(self) -> None:
        if self.indice >= len(self.pasos) - 1:
            return
        mensaje = self.paso_actual.validar()
        if mensaje:
            self.avisar(mensaje)
            return
        self.paso_actual.al_salir()
        self.completados.add(self.indice)
        self._mostrar(self.indice + 1)
        self.paso_actual.al_entrar()

    def anterior(self) -> None:
        if self.indice <= 0:
            return
        self.paso_actual.al_salir()
        self._mostrar(self.indice - 1)
        self.paso_actual.al_entrar()

    # ---- cierre ----
    def cancelar(self) -> None:
        if self.guardado:
            self.cerrar()
            return
        if self.estado.hay_datos and not messagebox.askyesno(
            "Unir archivos", TEXTO_CANCELAR, parent=self
        ):
            return
        self.cerrar()

    def cerrar(self) -> None:
        self.cerrada = True
        self.cuerpo.desactivar_rueda()
        try:
            self.destroy()
        except tk.TclError:
            pass


def abrir_asistente(
    parent: tk.Misc, estado: EstadoUnion | None = None, paso_inicial: int = 0
) -> Asistente:
    """Abre el asistente. Con `estado` (por ejemplo de una receta) arranca en `paso_inicial`."""
    return Asistente(parent, estado, paso_inicial)
