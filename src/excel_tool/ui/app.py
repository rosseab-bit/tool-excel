"""Ventana principal de la aplicación."""

import logging
import tkinter as tk
from importlib import resources
from tkinter import ttk

from excel_tool import APP_NAME, __version__

logger = logging.getLogger(__name__)

ANCHO_MINIMO = 480
ALTO_MINIMO = 320
TEXTO_PREPARACION = "Estamos preparando la aplicación. Muy pronto vas a poder usarla desde acá."


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


def crear_ventana() -> tk.Tk:
    """Arma la ventana principal sin iniciar el bucle de eventos."""
    ventana = tk.Tk()
    try:
        ventana.title(APP_NAME)
        ventana.minsize(ANCHO_MINIMO, ALTO_MINIMO)
        _cargar_icono(ventana)

        marco = ttk.Frame(ventana, padding=24)
        marco.pack(fill="both", expand=True)
        ttk.Label(marco, text=APP_NAME, font="TkHeadingFont").pack(pady=(40, 4))
        ttk.Label(marco, text=f"Versión {__version__}", font="TkSmallCaptionFont").pack()
        ttk.Label(
            marco, text=TEXTO_PREPARACION, font="TkTextFont", wraplength=400, justify="center"
        ).pack(pady=24)
        _centrar(ventana)
    except Exception:
        ventana.destroy()
        raise
    return ventana


def run() -> None:
    """Abre la ventana y espera hasta que se cierre."""
    ventana = crear_ventana()
    ventana.mainloop()
