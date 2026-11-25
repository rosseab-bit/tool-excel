"""Ventana principal de la aplicación."""

import logging
import tkinter as tk
from importlib import resources
from tkinter import ttk

from excel_tool import APP_NAME, __version__
from excel_tool.ui.explorador import abrir_explorador
from excel_tool.ui.selector import Seleccion, abrir_selector

logger = logging.getLogger(__name__)

ANCHO_MINIMO = 480
ALTO_MINIMO = 320
PROXIMAMENTE = "Próximamente"
OPCIONES_INICIO = (
    ("Ver información de un archivo", "Mirá qué tiene: columnas, tipos de datos y valores.", True),
    ("Unir archivos", "Combiná varios archivos a partir de una columna en común.", False),
    (
        "Juntar archivos uno debajo del otro",
        "Armá un solo archivo con las filas de varios.",
        False,
    ),
    ("Dividir un archivo", "Partí un archivo grande en varios más chicos.", False),
)


def _cargar_icono(ventana: tk.Tk) -> None:
    """Pone el ícono de la ventana. Si falla, lo registra y sigue."""
    try:
        ruta = resources.files("excel_tool") / "assets" / "icono.png"
        imagen = tk.PhotoImage(data=ruta.read_bytes())
        ventana.iconphoto(True, imagen)
        ventana._icono = imagen  # type: ignore[attr-defined]  # evita que se libere
    except Exception:
        logger.warning("No se pudo cargar el ícono", exc_info=True)


def _centrar(ventana: tk.Tk) -> None:
    ventana.update_idletasks()
    ancho = max(ventana.winfo_reqwidth(), ANCHO_MINIMO)
    alto = max(ventana.winfo_reqheight(), ALTO_MINIMO)
    x = max((ventana.winfo_screenwidth() - ancho) // 2, 0)
    y = max((ventana.winfo_screenheight() - alto) // 2, 0)
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")


class PantallaInicio(ttk.Frame):
    """Pantalla de inicio con las cuatro acciones principales."""

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, padding=24)
        self.pack(fill="both", expand=True)
        ttk.Label(self, text=APP_NAME, font="TkHeadingFont").pack(pady=(8, 2))
        ttk.Label(self, text=f"Versión {__version__}", font="TkSmallCaptionFont").pack()
        ttk.Label(self, text="¿Qué querés hacer?", font="TkTextFont").pack(pady=(16, 8))
        self.botones: list[ttk.Button] = []
        for titulo, explicacion, habilitado in OPCIONES_INICIO:
            marco = ttk.Frame(self)
            marco.pack(fill="x", pady=6)
            boton = ttk.Button(
                marco, text=titulo, command=self._ver_informacion if habilitado else None
            )
            boton.pack(fill="x", ipady=8)
            texto = explicacion if habilitado else f"{PROXIMAMENTE}. {explicacion}"
            ttk.Label(marco, text=texto, font="TkSmallCaptionFont", wraplength=420).pack(anchor="w")
            if not habilitado:
                boton.state(["disabled"])
            self.botones.append(boton)
        self.estado = ttk.Label(self, text="", wraplength=420)
        self.estado.pack(pady=(12, 0))
        self.selector = None
        self.explorador = None

    def _ver_informacion(self) -> None:
        self.selector = abrir_selector(self.winfo_toplevel(), self._cargar)

    def _cargar(self, eleccion: Seleccion) -> None:
        self.estado.configure(text="")
        self.explorador = abrir_explorador(self.winfo_toplevel(), eleccion)


def crear_ventana() -> tk.Tk:
    """Arma la ventana principal sin iniciar el bucle de eventos."""
    ventana = tk.Tk()
    try:
        ventana.title(APP_NAME)
        ventana.minsize(ANCHO_MINIMO, ALTO_MINIMO)
        _cargar_icono(ventana)

        PantallaInicio(ventana)
        _centrar(ventana)
    except Exception:
        ventana.destroy()
        raise
    return ventana


def run() -> None:
    """Abre la ventana y espera hasta que se cierre."""
    ventana = crear_ventana()
    ventana.mainloop()
