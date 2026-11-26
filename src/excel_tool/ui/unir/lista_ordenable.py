"""Lista de textos que se puede reordenar con botones (↑ ↓, arriba de todo, abajo de todo)."""

from __future__ import annotations

from collections.abc import Callable, Hashable
from tkinter import ttk


class ListaOrdenable(ttk.Frame):
    def __init__(
        self,
        parent,
        *,
        on_change: Callable[[], None] | None = None,
        alto: int = 12,
    ) -> None:
        super().__init__(parent)
        self._on_change = on_change
        self._datos: list[tuple[Hashable, str]] = []
        self.arbol = ttk.Treeview(self, columns=(), show="tree", selectmode="browse", height=alto)
        self.arbol.column("#0", width=240, stretch=True)
        barra = ttk.Scrollbar(self, orient="vertical", command=self.arbol.yview)
        self.arbol.configure(yscrollcommand=barra.set)
        self.arbol.grid(row=0, column=0, sticky="nsew")
        barra.grid(row=0, column=1, sticky="ns")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        botones = ttk.Frame(self)
        botones.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self.botones: dict[str, ttk.Button] = {}
        for clave, texto, accion in (
            ("arriba_todo", "⤒ Arriba de todo", self.arriba_de_todo),
            ("subir", "↑ Subir", self.subir),
            ("bajar", "↓ Bajar", self.bajar),
            ("abajo_todo", "⤓ Abajo de todo", self.abajo_de_todo),
        ):
            boton = ttk.Button(botones, text=texto, command=accion)
            boton.pack(side="left", padx=(0, 4))
            self.botones[clave] = boton

    # ---- datos ----
    def poner(self, elementos: list[tuple[Hashable, str]], seleccionado: int | None = None) -> None:
        """Cada elemento es (clave, texto que se ve); las claves identifican aunque se repita
        el texto."""
        self._datos = list(elementos)
        self.arbol.delete(*self.arbol.get_children())
        for i, (_, texto) in enumerate(self._datos):
            self.arbol.insert("", "end", iid=str(i), text=texto)
        if seleccionado is not None and 0 <= seleccionado < len(self._datos):
            self.seleccionar(seleccionado)

    def claves(self) -> list[Hashable]:
        return [c for c, _ in self._datos]

    def textos(self) -> list[str]:
        return [t for _, t in self._datos]

    @property
    def indice(self) -> int | None:
        actual = self.arbol.selection()
        return int(actual[0]) if actual else None

    def seleccionar(self, indice: int) -> None:
        self.arbol.selection_set(str(indice))
        self.arbol.see(str(indice))

    # ---- movimientos ----
    def _mover(self, destino: int) -> None:
        origen = self.indice
        if origen is None or not 0 <= destino < len(self._datos) or destino == origen:
            return
        elemento = self._datos.pop(origen)
        self._datos.insert(destino, elemento)
        self.poner(self._datos, destino)
        if self._on_change:
            self._on_change()

    def subir(self) -> None:
        if self.indice is not None:
            self._mover(self.indice - 1)

    def bajar(self) -> None:
        if self.indice is not None:
            self._mover(self.indice + 1)

    def arriba_de_todo(self) -> None:
        self._mover(0)

    def abajo_de_todo(self) -> None:
        self._mover(len(self._datos) - 1)
