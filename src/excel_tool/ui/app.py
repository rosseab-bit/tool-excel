"""Ventana principal de la aplicación."""

import logging
import tkinter as tk
from functools import partial
from importlib import resources
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from excel_tool import APP_NAME, __version__
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.preferences import load_preferences
from excel_tool.core.recipe import load_recipe
from excel_tool.ui.explorador import abrir_explorador
from excel_tool.ui.selector import Seleccion, abrir_selector
from excel_tool.ui.unir.asistente import abrir_asistente
from excel_tool.ui.unir.estado import EstadoUnion, leer_fuente
from excel_tool.ui.worker import MENSAJE_INESPERADO, run_in_background

logger = logging.getLogger(__name__)

ANCHO_MINIMO = 480
ALTO_MINIMO = 320
PROXIMAMENTE = "Próximamente"
PASO_RECETA = 2  # al abrir una receta se arranca en "Archivos a sumar"
# (título, explicación, nombre del método que se ejecuta; None = todavía no está)
OPCIONES_INICIO = (
    (
        "Ver información de un archivo",
        "Mirá qué tiene: columnas, tipos de datos y valores.",
        "_ver_informacion",
    ),
    (
        "Unir archivos",
        "Combiná varios archivos a partir de una columna en común.",
        "_unir_archivos",
    ),
    (
        "Abrir una receta…",
        "Repetí una unión que ya guardaste, por ejemplo con los archivos del mes siguiente.",
        "_abrir_receta",
    ),
    (
        "Juntar archivos uno debajo del otro",
        "Armá un solo archivo con las filas de varios.",
        None,
    ),
    ("Dividir un archivo", "Partí un archivo grande en varios más chicos.", None),
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
        for titulo, explicacion, metodo in OPCIONES_INICIO:
            habilitado = metodo is not None
            marco = ttk.Frame(self)
            marco.pack(fill="x", pady=6)
            boton = ttk.Button(
                marco, text=titulo, command=getattr(self, metodo) if metodo else None
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
        self.asistente = None

    def _ver_informacion(self) -> None:
        self.selector = abrir_selector(self.winfo_toplevel(), self._cargar)

    def _unir_archivos(self) -> None:
        self.estado.configure(text="")
        self.asistente = abrir_asistente(self.winfo_toplevel())

    def _abrir_receta(self) -> None:
        carpeta = load_preferences().last_folder
        opciones = {
            "parent": self.winfo_toplevel(),
            "title": "Elegí la receta",
            "filetypes": [("Recetas", "*.json"), ("Todos los archivos", "*.*")],
        }
        if carpeta is not None and Path(carpeta).is_dir():
            opciones["initialdir"] = str(carpeta)
        ruta = filedialog.askopenfilename(**opciones)
        if not ruta:
            return
        self.estado.configure(text="")

        def trabajo(progress, cancel):
            carga = load_recipe(Path(ruta))
            estado = EstadoUnion.desde_receta(
                carga, leer=partial(leer_fuente, progress=progress, cancel=cancel)
            )
            return estado, carga.warnings

        run_in_background(
            self.winfo_toplevel(),
            trabajo,
            title="Abriendo la receta",
            on_success=self._receta_lista,
            on_error=self._receta_fallo,
        )

    def _receta_lista(self, resultado: tuple[EstadoUnion, list[str]]) -> None:
        estado, avisos = resultado
        for aviso in avisos:
            logger.info("Receta: %s", aviso)
        faltan = estado.faltantes
        if faltan:
            lista = "\n".join(f"• {nombre}" for nombre in faltan)
            messagebox.showinfo(
                "Faltan archivos",
                "No encontramos estos archivos de la receta. Vas a poder elegir otros "
                f"en el asistente:\n\n{lista}",
                parent=self.winfo_toplevel(),
            )
        self.asistente = abrir_asistente(self.winfo_toplevel(), estado, PASO_RECETA)

    def _receta_fallo(self, exc: BaseException) -> None:
        if isinstance(exc, ExcelToolError):
            logger.warning("No se pudo abrir la receta: %s", exc)
            texto = str(exc)
        else:
            logger.error("Falló al abrir la receta", exc_info=exc)
            texto = MENSAJE_INESPERADO
        messagebox.showerror("No se pudo abrir la receta", texto, parent=self.winfo_toplevel())

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
