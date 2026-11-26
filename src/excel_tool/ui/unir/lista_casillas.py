"""Lista de columnas con casilla, buscador y "Marcar todas / Quitar todas"."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from tkinter import ttk

ALTO_MAXIMO = 150
ALTO_FILA = 24


class ListaCasillas(ttk.Frame):
    def __init__(
        self,
        parent,
        columnas: list[str],
        marcadas: set[str],
        *,
        fijas: frozenset[str] = frozenset(),
        on_change: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self.columnas = list(columnas)
        self.fijas = fijas
        self._on_change = on_change
        self.busqueda = tk.StringVar(master=self)
        cabecera = ttk.Frame(self)
        cabecera.pack(fill="x")
        ttk.Label(cabecera, text="Buscar:").pack(side="left")
        self.entrada = ttk.Entry(cabecera, textvariable=self.busqueda, width=18)
        self.entrada.pack(side="left", padx=4)
        self.boton_todas = ttk.Button(cabecera, text="Marcar todas", command=self.marcar_todas)
        self.boton_todas.pack(side="left", padx=2)
        self.boton_ninguna = ttk.Button(cabecera, text="Quitar todas", command=self.quitar_todas)
        self.boton_ninguna.pack(side="left", padx=2)
        self.cuenta = ttk.Label(cabecera)
        self.cuenta.pack(side="right")

        cuerpo = ttk.Frame(self)
        cuerpo.pack(fill="x", pady=(4, 0))
        self.lienzo = tk.Canvas(cuerpo, highlightthickness=0, borderwidth=0, height=ALTO_MAXIMO)
        self.barra = ttk.Scrollbar(cuerpo, orient="vertical", command=self.lienzo.yview)
        self.lienzo.configure(yscrollcommand=self.barra.set)
        self.lienzo.pack(side="left", fill="x", expand=True)
        self.barra.pack(side="right", fill="y")
        self.marco = ttk.Frame(self.lienzo)
        self.lienzo.create_window((0, 0), window=self.marco, anchor="nw")
        self.marco.bind("<Configure>", self._ajustar)

        self.variables: dict[str, tk.BooleanVar] = {}
        self.casillas: dict[str, ttk.Checkbutton] = {}
        for c in self.columnas:
            var = tk.BooleanVar(master=self.marco, value=c in marcadas or c in fijas)
            texto = f"{c}  (clave)" if c in fijas else c
            casilla = ttk.Checkbutton(self.marco, text=texto, variable=var, command=self._cambio)
            if c in fijas:
                casilla.state(["disabled"])
            self.variables[c] = var
            self.casillas[c] = casilla
        self.busqueda.trace_add("write", lambda *_: self._filtrar())
        self._filtrar()

    def _ajustar(self, _evento: object = None) -> None:
        self.lienzo.configure(scrollregion=self.lienzo.bbox("all"))
        alto = min(max(self.marco.winfo_reqheight(), ALTO_FILA), ALTO_MAXIMO)
        self.lienzo.configure(height=alto)

    # ---- consultas ----
    def marcadas(self) -> list[str]:
        """Las columnas marcadas, en el orden del archivo."""
        return [c for c in self.columnas if self.variables[c].get()]

    def visibles(self) -> list[str]:
        texto = self.busqueda.get().strip().casefold()
        return [c for c in self.columnas if not texto or texto in c.casefold()]

    # ---- acciones ----
    def _filtrar(self) -> None:
        for casilla in self.casillas.values():
            casilla.pack_forget()
        for c in self.visibles():
            self.casillas[c].pack(anchor="w", fill="x")
        self._cuenta()
        self.lienzo.yview_moveto(0)

    def _cuenta(self) -> None:
        self.cuenta.configure(text=f"{len(self.marcadas())} de {len(self.columnas)}")

    def _cambio(self) -> None:
        self._cuenta()
        if self._on_change:
            self._on_change()

    def _poner(self, valor: bool) -> None:
        for c in self.visibles():
            if c not in self.fijas:
                self.variables[c].set(valor)
        self._cambio()

    def marcar_todas(self) -> None:
        self._poner(True)

    def quitar_todas(self) -> None:
        self._poner(False)
