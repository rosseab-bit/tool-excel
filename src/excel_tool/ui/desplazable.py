"""Cuerpo con desplazamiento vertical y pie fijo para ventanas que pueden quedar altas.

El pie (botones) queda siempre a la vista, abajo. El cuerpo va dentro de un lienzo con barra
vertical que aparece solo si hace falta y responde a la rueda del mouse (Windows y Linux).
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class CuerpoDesplazable:
    def __init__(
        self,
        ventana: tk.Misc,
        *,
        alto_maximo: int,
        padding_cuerpo: int = 12,
        padding_pie: tuple[int, int, int, int] = (12, 0, 12, 10),
    ) -> None:
        self.ventana = ventana
        self.alto_maximo = alto_maximo
        self.pie = ttk.Frame(ventana, padding=padding_pie)
        self.pie.pack(side="bottom", fill="x")
        contenedor = ttk.Frame(ventana)
        contenedor.pack(side="top", fill="both", expand=True)
        self.lienzo = tk.Canvas(contenedor, highlightthickness=0, borderwidth=0)
        self.barra_v = ttk.Scrollbar(contenedor, orient="vertical", command=self.lienzo.yview)
        self.lienzo.configure(yscrollcommand=self.barra_v.set)
        self.barra_v.pack(side="right", fill="y")
        self.lienzo.pack(side="left", fill="both", expand=True)
        self.marco = ttk.Frame(self.lienzo, padding=padding_cuerpo)
        self._id_marco = self.lienzo.create_window((0, 0), window=self.marco, anchor="nw")
        self.marco.bind("<Configure>", self._al_cambiar_marco)
        self.lienzo.bind("<Configure>", self._al_cambiar_lienzo)
        self.lienzo.bind("<Enter>", self._activar_rueda)
        self.lienzo.bind("<Leave>", self._desactivar_rueda)

    def _al_cambiar_marco(self, _evento: object = None) -> None:
        self.lienzo.configure(scrollregion=self.lienzo.bbox("all"))
        pedido = self.marco.winfo_reqheight()
        self.lienzo.configure(
            height=min(pedido, self.alto_maximo), width=self.marco.winfo_reqwidth()
        )
        hay_barra = pedido > self.alto_maximo
        if hay_barra and not self.barra_v.winfo_ismapped():
            self.barra_v.pack(side="right", fill="y", before=self.lienzo)
        elif not hay_barra and self.barra_v.winfo_ismapped():
            self.barra_v.pack_forget()

    def _al_cambiar_lienzo(self, evento: tk.Event) -> None:
        self.lienzo.itemconfigure(self._id_marco, width=evento.width)

    def _activar_rueda(self, _evento: object = None) -> None:
        self.ventana.bind_all("<MouseWheel>", self._rueda)
        self.ventana.bind_all("<Button-4>", self._rueda)
        self.ventana.bind_all("<Button-5>", self._rueda)

    def desactivar_rueda(self, _evento: object = None) -> None:
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            try:
                self.ventana.unbind_all(ev)
            except tk.TclError:
                pass

    _desactivar_rueda = desactivar_rueda

    def _rueda(self, evento: tk.Event) -> None:
        if not self.barra_v.winfo_ismapped():
            return
        if getattr(evento, "num", 0) == 4:
            paso = -1
        elif getattr(evento, "num", 0) == 5:
            paso = 1
        else:
            paso = -1 if evento.delta > 0 else 1
        self.lienzo.yview_scroll(paso, "units")
