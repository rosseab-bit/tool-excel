"""Clase base de los pasos del asistente (docs/contratos.md § Fase 5 › Asistente)."""

from __future__ import annotations

from tkinter import ttk
from typing import TYPE_CHECKING

from excel_tool.ui.unir.estado import EstadoUnion

if TYPE_CHECKING:
    from excel_tool.ui.unir.asistente import Asistente


class Paso(ttk.Frame):
    """Un paso del asistente. Se crea una sola vez y se muestra/oculta, así al volver atrás no
    se pierde nada."""

    titulo: str = ""

    def __init__(self, parent, estado: EstadoUnion, asistente: Asistente) -> None:
        super().__init__(parent)
        self.estado = estado
        self.asistente = asistente

    def al_entrar(self) -> None:
        """Se llama al mostrar el paso: refresca desde el estado."""

    def validar(self) -> str | None:
        """None = puede seguir; un texto = por qué no (se muestra en el pie)."""
        return None

    def al_salir(self) -> None:
        """Se llama al avanzar (validado): vuelca lo elegido al estado."""


class PasoProvisorio(Paso):
    """Reemplazar en las tareas F5-11 a F5-14: solo muestra un texto y deja seguir."""

    def __init__(self, parent, estado: EstadoUnion, asistente: Asistente) -> None:
        super().__init__(parent, estado, asistente)
        ttk.Label(self, text=self.titulo, font="TkHeadingFont").pack(anchor="w", pady=(0, 8))
        ttk.Label(self, text="Este paso todavía no está listo.").pack(anchor="w")
