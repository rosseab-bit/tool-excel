"""Paso 1: elegir el archivo principal (el que manda: una fila del resultado por cada suya)."""

from __future__ import annotations

from tkinter import ttk

from excel_tool.ui.unir.estado import Fuente
from excel_tool.ui.unir.pasos.base import Paso
from excel_tool.ui.unir.widgets_archivos import pedir_fuente, resumen_fuente, ver_informacion

TEXTO_AYUDA = (
    "Este es el archivo base: el resultado va a tener una fila por cada fila de este archivo "
    "(salvo que traigas el detalle de otro)."
)
TEXTO_SIN_ARCHIVO = "Todavía no elegiste ningún archivo."
MENSAJE_FALTA = "Elegí el archivo principal para seguir."


class PasoPrincipal(Paso):
    titulo = "Archivo principal"

    def __init__(self, parent, estado, asistente) -> None:
        super().__init__(parent, estado, asistente)
        ttk.Label(self, text="¿Cuál es el archivo principal?", font="TkHeadingFont").pack(
            anchor="w", pady=(0, 6)
        )
        ttk.Label(self, text=TEXTO_AYUDA, wraplength=760, justify="left").pack(
            anchor="w", pady=(0, 12)
        )
        fila = ttk.Frame(self)
        fila.pack(anchor="w", fill="x")
        self.boton_elegir = ttk.Button(fila, text="Elegir archivo…", command=self._elegir)
        self.boton_elegir.pack(side="left")
        self.boton_info = ttk.Button(fila, text="🔍 Ver información", command=self._ver)
        self.boton_info.pack(side="left", padx=(8, 0))
        self.etiqueta_nombre = ttk.Label(self, text=TEXTO_SIN_ARCHIVO, font="TkHeadingFont")
        self.etiqueta_nombre.pack(anchor="w", pady=(14, 0))
        self.etiqueta_detalle = ttk.Label(self, text="")
        self.etiqueta_detalle.pack(anchor="w")
        self.al_entrar()

    # ---- pantalla ----
    def al_entrar(self) -> None:
        principal = self.estado.principal
        if principal is None:
            self.etiqueta_nombre.configure(text=TEXTO_SIN_ARCHIVO)
            self.etiqueta_detalle.configure(text="")
            self.boton_info.state(["disabled"])
            return
        if principal.falta:
            self.etiqueta_nombre.configure(text=principal.nombre)
            self.etiqueta_detalle.configure(
                text="Falta el archivo: elegí el de este mes con «Elegir archivo…»."
            )
            self.boton_info.state(["disabled"])
            return
        self.etiqueta_nombre.configure(text=principal.nombre)
        self.etiqueta_detalle.configure(text=resumen_fuente(principal))
        self.boton_info.state(["!disabled"])

    # ---- acciones ----
    def _elegir(self) -> None:
        pedir_fuente(self, self._nombres_usados, self._cargado)

    def _nombres_usados(self) -> list[str]:
        return [s.nombre for s in self.estado.secundarios]

    def _cargado(self, fuente: Fuente) -> None:
        if not self.winfo_exists():
            return
        estado = self.estado
        actual = estado.principal
        if actual is not None and actual.falta:
            estado.completar(actual.nombre, fuente)  # receta: se conserva todo lo demás
        else:
            if actual is not None:
                estado.columnas.pop(actual.nombre, None)
            estado.principal = fuente
            estado.claves = []
            estado.orden = []
        self.al_entrar()
        self.asistente.avisar("")

    def _ver(self) -> None:
        if self.estado.principal is not None and not self.estado.principal.falta:
            ver_informacion(self, self.estado.principal)

    # ---- validación ----
    def validar(self) -> str | None:
        principal = self.estado.principal
        if principal is None or principal.falta:
            return MENSAJE_FALTA
        return None
