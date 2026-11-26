"""Paso 2: elegir la o las columnas clave del archivo principal."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from excel_tool.ui.unir.pasos.base import Paso

TEXTO_AYUDA = "La columna que identifica a cada cliente, producto, etc."
TEXTO_MAS_DE_UNA = "Si hace falta más de una (por ejemplo, cliente y artículo), marcá todas."
TEXTO_SIN_ARCHIVO = "Primero elegí el archivo principal (paso anterior)."
MENSAJE_FALTA = "Marcá al menos una columna clave para seguir."
MARCA_CLAVE = "🔑"


class PasoClave(Paso):
    titulo = "Columna clave"

    def __init__(self, parent, estado, asistente) -> None:
        super().__init__(parent, estado, asistente)
        self.casillas: dict[str, ttk.Checkbutton] = {}
        self._variables: dict[str, tk.BooleanVar] = {}
        self._archivo: object = None
        ttk.Label(self, text="¿Cuál es la columna clave?", font="TkHeadingFont").pack(
            anchor="w", pady=(0, 6)
        )
        ttk.Label(self, text=TEXTO_AYUDA).pack(anchor="w")
        ttk.Label(self, text=TEXTO_MAS_DE_UNA).pack(anchor="w", pady=(0, 10))
        self.lista = ttk.Frame(self)
        self.lista.pack(anchor="w", fill="x")
        self.al_entrar()

    # ---- pantalla ----
    def _columnas_ordenadas(self) -> list[tuple[str, bool]]:
        """(nombre, es posible clave), con las posibles claves primero y el resto en su orden."""
        principal = self.estado.principal
        if principal is None or principal.df is None:
            return []
        unicas = {}
        if principal.perfil is not None:
            unicas = {c.name: bool(c.is_unique) for c in principal.perfil.columns_profile}
        nombres = [str(c) for c in principal.df.columns]
        claves = [(n, unicas.get(n, False)) for n in nombres]
        return [p for p in claves if p[1]] + [p for p in claves if not p[1]]

    def al_entrar(self) -> None:
        principal = self.estado.principal
        if principal is not self._archivo:
            self._archivo = principal
            self._dibujar()
        elif principal is not None:
            self._marcar_desde_estado()

    def _dibujar(self) -> None:
        for hijo in self.lista.winfo_children():
            hijo.destroy()
        self.casillas = {}
        self._variables = {}
        columnas = self._columnas_ordenadas()
        if not columnas:
            ttk.Label(self.lista, text=TEXTO_SIN_ARCHIVO).pack(anchor="w")
            return
        for nombre, posible in columnas:
            var = tk.BooleanVar(master=self.lista, value=nombre in self.estado.claves)
            texto = f"{MARCA_CLAVE} {nombre}" if posible else nombre
            casilla = ttk.Checkbutton(self.lista, text=texto, variable=var, command=self._cambio)
            casilla.pack(anchor="w", pady=1)
            self.casillas[nombre] = casilla
            self._variables[nombre] = var

    def _marcar_desde_estado(self) -> None:
        for nombre, var in self._variables.items():
            var.set(nombre in self.estado.claves)

    # ---- estado ----
    def _elegidas(self) -> list[str]:
        """Las marcadas, en el orden de las columnas del archivo."""
        principal = self.estado.principal
        if principal is None or principal.df is None:
            return []
        return [
            str(c)
            for c in principal.df.columns
            if self._variables.get(str(c)) is not None and self._variables[str(c)].get()
        ]

    def _cambio(self) -> None:
        self.estado.claves = self._elegidas()
        self.asistente.avisar("")

    def al_salir(self) -> None:
        self.estado.claves = self._elegidas()

    def validar(self) -> str | None:
        if not self._elegidas():
            return MENSAJE_FALTA
        return None
