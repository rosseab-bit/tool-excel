"""Selector de archivo: devuelve la elección del usuario (ruta, hoja, texto, encabezado).

No devuelve datos: quien lo usa carga el archivo después con ``read_table``.
"""

from __future__ import annotations

import logging
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from excel_tool.core.io import CsvOptions, ReadOptions, SheetSummary, list_sheets, read_table
from excel_tool.core.preferences import Preferences, load_preferences, save_preferences
from excel_tool.ui.grilla import Grilla, formatear_numero
from excel_tool.ui.worker import run_in_background

log = logging.getLogger(__name__)

FILAS_VISTA_PREVIA = 10
EXTENSIONES_CSV = (".csv", ".txt")
TIPOS_ARCHIVO = [
    ("Archivos de Excel y de texto", "*.csv *.txt *.xlsx *.xlsm *.xlsb *.xls *.ods"),
    ("Todos los archivos", "*.*"),
]

SEPARADORES = {
    ";": "Punto y coma ( ; )",
    ",": "Coma ( , )",
    "\t": "Tabulación",
    "|": "Barra vertical ( | )",
}
CODIFICACIONES = {
    "utf-8-sig": "Recomendado (con acentos y ñ)",
    "utf-8": "Estándar (con acentos y ñ)",
    "cp1252": "Windows antiguo (con acentos y ñ)",
}

TEXTO_ERROR_VISTA = "No se pudo leer el archivo con estas opciones."
TEXTO_MEMORIA_NO_DISPONIBLE = "No se pudo comprobar la memoria libre; se sigue igual."


@dataclass(frozen=True)
class Seleccion:
    """Lo que eligió el usuario. ``header_row`` va en base 0 (0 = primera fila)."""

    path: Path
    sheet: str | None
    csv_options: CsvOptions | None
    header_row: int

    def read_options(self) -> ReadOptions:
        return ReadOptions(sheet=self.sheet, header_row=self.header_row, csv=self.csv_options)


def comprobar_memoria(path: Path) -> tuple[bool, str]:
    """Devuelve (se puede cargar, mensaje). El import va acá para poder reemplazarla en tests."""
    try:
        from excel_tool.core.memory import check_memory
    except ImportError:
        log.warning("core.memory todavía no existe; no se controla la memoria.")
        return True, ""
    try:
        resultado = check_memory(path)
    except Exception:
        log.warning("Falló el control de memoria", exc_info=True)
        return True, ""
    return bool(resultado.ok), str(resultado.message)


def pedir_archivo(parent: tk.Misc) -> Path | None:
    """Abre el diálogo de archivos en la última carpeta usada."""
    carpeta = load_preferences().last_folder
    opciones: dict[str, Any] = {
        "parent": parent,
        "title": "Elegí un archivo",
        "filetypes": TIPOS_ARCHIVO,
    }
    if carpeta is not None and carpeta.is_dir():
        opciones["initialdir"] = str(carpeta)
    elegido = filedialog.askopenfilename(**opciones)
    return Path(elegido) if elegido else None


def es_csv(path: Path) -> bool:
    return path.suffix.lower() in EXTENSIONES_CSV


class SelectorArchivo(tk.Toplevel):
    """Ventana de elección de archivo. Llama a ``on_accept(Seleccion)`` al aceptar.

    Si se pasa ``path`` no abre el diálogo de archivos (útil para reabrir con otro archivo).
    """

    def __init__(
        self,
        parent: tk.Misc,
        on_accept: Callable[[Seleccion], None],
        path: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.on_accept = on_accept
        self._padre = parent
        self.path: Path | None = None
        self.cerrada = False
        self._hojas: list[SheetSummary] = []
        self._hoja: str | None = None
        self._csv: CsvOptions | None = None
        self._encabezado: int | None = None  # base 0; None = detectar
        self._encabezado_usado = 0
        self._df: Any = None
        self._pendientes = 0
        self._version = 0
        self._arrancado = False

        self.title("Elegir archivo")
        self.minsize(640, 460)
        self.transient(parent.winfo_toplevel())  # type: ignore[arg-type]
        self.protocol("WM_DELETE_WINDOW", self.cancelar)
        self._armar()
        self.withdraw()
        self.after_idle(lambda: self._arrancar(path))

    # ---- estado ----
    @property
    def ocupado(self) -> bool:
        return self._pendientes > 0

    @property
    def listo(self) -> bool:
        return self.path is not None and self._df is not None and not self.ocupado

    # ---- armado ----
    def _armar(self) -> None:
        marco = ttk.Frame(self, padding=16)
        marco.pack(fill="both", expand=True)
        self.etiqueta_archivo = ttk.Label(marco, text="", font="TkHeadingFont", wraplength=600)
        self.etiqueta_archivo.pack(anchor="w")

        self.opciones = ttk.Frame(marco)
        self.opciones.pack(fill="x", pady=(12, 8))
        self.opciones.columnconfigure(1, weight=1)
        self._etiquetas: dict[ttk.Frame, ttk.Label] = {}

        self.var_hoja = tk.StringVar()
        self.var_sep = tk.StringVar()
        self.var_cod = tk.StringVar()
        self.var_encabezado = tk.StringVar(value="1")

        self.fila_hoja = self._fila(0, "Hoja")
        self.combo_hoja = ttk.Combobox(
            self.fila_hoja, textvariable=self.var_hoja, state="readonly", width=48
        )
        self.combo_hoja.pack(side="left")
        self.combo_hoja.bind("<<ComboboxSelected>>", lambda _e: self._eligio_hoja())

        self.fila_sep = self._fila(1, "Separador de columnas")
        self.combo_sep = ttk.Combobox(
            self.fila_sep,
            textvariable=self.var_sep,
            state="readonly",
            values=list(SEPARADORES.values()),
            width=30,
        )
        self.combo_sep.pack(side="left")
        self.combo_sep.bind("<<ComboboxSelected>>", lambda _e: self._eligio_separador())

        self.fila_cod = self._fila(2, "Tipo de texto (si ves letras raras, probá otra opción)")
        self.combo_cod = ttk.Combobox(
            self.fila_cod,
            textvariable=self.var_cod,
            state="readonly",
            values=list(CODIFICACIONES.values()),
            width=34,
        )
        self.combo_cod.pack(side="left")
        self.combo_cod.bind("<<ComboboxSelected>>", lambda _e: self._eligio_codificacion())

        self.fila_enc = self._fila(3, "Fila donde están los títulos de las columnas")
        self.spin_encabezado = ttk.Spinbox(
            self.fila_enc,
            from_=1,
            to=1000,
            width=6,
            textvariable=self.var_encabezado,
            command=self._eligio_encabezado,
        )
        self.spin_encabezado.pack(side="left")
        for evento in ("<Return>", "<FocusOut>"):
            self.spin_encabezado.bind(evento, lambda _e: self._eligio_encabezado(), add="+")

        self.etiqueta_estado = ttk.Label(marco, text="", wraplength=600)
        self.etiqueta_estado.pack(anchor="w", pady=(0, 4))
        ttk.Label(marco, text="Vista previa (primeras filas)").pack(anchor="w")
        self.grilla = Grilla(marco, proveedor=None, page_size=FILAS_VISTA_PREVIA)
        self.grilla.pack(fill="both", expand=True, pady=(4, 8))

        botones = ttk.Frame(marco)
        botones.pack(fill="x")
        self.boton_cancelar = ttk.Button(botones, text="Cancelar", command=self.cancelar)
        self.boton_cancelar.pack(side="right")
        self.boton_aceptar = ttk.Button(botones, text="Aceptar", command=self.aceptar)
        self.boton_aceptar.pack(side="right", padx=8)
        self.boton_otro = ttk.Button(botones, text="Elegir otro archivo", command=self.otro_archivo)
        self.boton_otro.pack(side="left")
        self.boton_aceptar.state(["disabled"])

    def _fila(self, n: int, texto: str) -> ttk.Frame:
        etiqueta = ttk.Label(self.opciones, text=texto)
        etiqueta.grid(row=n, column=0, sticky="w", pady=3, padx=(0, 12))
        fila = ttk.Frame(self.opciones)
        fila.grid(row=n, column=1, sticky="w", pady=3)
        self._etiquetas[fila] = etiqueta
        return fila

    def _mostrar_fila(self, nombre: str, visible: bool) -> None:
        marco: ttk.Frame = getattr(self, nombre)
        for w in (marco, self._etiquetas[marco]):
            if visible:
                w.grid()
            else:
                w.grid_remove()

    # ---- arranque ----
    def _arrancar(self, path: Path | None) -> None:
        if self.cerrada:
            return
        if path is None:
            path = pedir_archivo(self._padre)
        if path is None:
            self.cancelar()
            return
        self._empezar_con(path)

    def _empezar_con(self, path: Path) -> None:
        ok, mensaje = comprobar_memoria(path)
        if not ok:
            messagebox.showwarning("Archivo muy grande", mensaje, parent=self._padre)
            self.cancelar()
            return
        self.path = path
        self._arrancado = True
        self._hojas, self._hoja, self._csv, self._encabezado, self._df = [], None, None, None, None
        self.var_encabezado.set("1")
        self.etiqueta_archivo.configure(text=path.name)
        self.boton_aceptar.state(["disabled"])
        self.deiconify()
        self.lift()
        csv = es_csv(path)
        self._mostrar_fila("fila_hoja", False)
        self._mostrar_fila("fila_sep", csv)
        self._mostrar_fila("fila_cod", csv)
        if csv:
            self._pedir_vista_previa()
        else:
            self._pedir_hojas()

    def otro_archivo(self) -> None:
        nuevo = pedir_archivo(self)
        if nuevo is not None:
            self._empezar_con(nuevo)

    # ---- trabajo en segundo plano ----
    def _pedir_hojas(self) -> None:
        assert self.path is not None
        ruta = self.path
        self._pendientes += 1
        self.etiqueta_estado.configure(text="Mirando las hojas del archivo…")

        def terminar(hojas: list[SheetSummary]) -> None:
            self._pendientes -= 1
            if self.cerrada or ruta != self.path:
                return
            self._hojas = hojas
            self._llenar_hojas()
            self._pedir_vista_previa()

        def fallar(exc: BaseException) -> None:
            self._pendientes -= 1
            self._error(exc, "No se pudieron leer las hojas del archivo.")

        run_in_background(
            self,
            lambda _p, _c: list_sheets(ruta),
            title="Leyendo el archivo",
            on_success=terminar,
            on_error=fallar,
            cancellable=False,
        )

    def _llenar_hojas(self) -> None:
        textos = [
            f"{h.name} ({formatear_numero(h.rows)} filas, {formatear_numero(h.columns)} columnas)"
            for h in self._hojas
        ]
        self.combo_hoja.configure(values=textos)
        if self._hojas:
            self._hoja = self._hojas[0].name
            self.var_hoja.set(textos[0])
        self._mostrar_fila("fila_hoja", len(self._hojas) > 1)

    def _pedir_vista_previa(self) -> None:
        assert self.path is not None
        ruta = self.path
        opciones = ReadOptions(
            sheet=self._hoja,
            header_row=self._encabezado,
            csv=self._csv,
            nrows=FILAS_VISTA_PREVIA,
        )
        self._version += 1
        version = self._version
        self._pendientes += 1
        self._df = None
        self.boton_aceptar.state(["disabled"])
        self.etiqueta_estado.configure(text="Cargando la vista previa…")

        def terminar(res: Any) -> None:
            self._pendientes -= 1
            if self.cerrada or version != self._version:
                return
            self._mostrar_vista(res)

        def fallar(exc: BaseException) -> None:
            self._pendientes -= 1
            if version == self._version:
                self._error(exc, TEXTO_ERROR_VISTA)

        run_in_background(
            self,
            lambda _p, _c: read_table(ruta, opciones),
            title="Leyendo el archivo",
            on_success=terminar,
            on_error=fallar,
            cancellable=False,
            modal=False,
        )

    def _mostrar_vista(self, res: Any) -> None:
        self._df = res.df
        self._encabezado_usado = res.header_row
        if res.csv_options is not None:
            self._csv = res.csv_options
            self.var_sep.set(SEPARADORES.get(res.csv_options.separator, res.csv_options.separator))
            self.var_cod.set(CODIFICACIONES.get(res.csv_options.encoding, res.csv_options.encoding))
        self.var_encabezado.set(str(res.header_row + 1))
        df = res.df
        self.grilla.set_proveedor(lambda _pagina: (df, len(df)))
        columnas = len(df.columns)
        self.etiqueta_estado.configure(
            text=f"Se detectaron {formatear_numero(columnas)} columnas." if columnas else ""
        )
        if res.warnings:
            self.etiqueta_estado.configure(text=" ".join(res.warnings))
        self.boton_aceptar.state(["!disabled"])

    def _error(self, exc: BaseException, generico: str) -> None:
        mensaje = getattr(exc, "mensaje", None) or generico
        self.etiqueta_estado.configure(text=mensaje)
        self.grilla.set_proveedor(None)
        self.boton_aceptar.state(["disabled"])
        messagebox.showerror("No se pudo leer", mensaje, parent=self)

    # ---- cambios del usuario ----
    def seleccionar_hoja(self, nombre: str) -> None:
        for h in self._hojas:
            if h.name == nombre:
                self.var_hoja.set(self.combo_hoja.cget("values")[self._hojas.index(h)])
        self._eligio_hoja()

    def _eligio_hoja(self) -> None:
        indice = self.combo_hoja.current()
        if indice < 0:
            valores = list(self.combo_hoja.cget("values"))
            indice = valores.index(self.var_hoja.get()) if self.var_hoja.get() in valores else -1
        if indice < 0 or self._hojas[indice].name == self._hoja:
            return
        self._hoja = self._hojas[indice].name
        self._encabezado = None  # cada hoja tiene su propio encabezado
        self._pedir_vista_previa()

    def seleccionar_separador(self, separador: str) -> None:
        self.var_sep.set(SEPARADORES[separador])
        self._eligio_separador()

    def _eligio_separador(self) -> None:
        por_texto = {v: k for k, v in SEPARADORES.items()}
        sep = por_texto.get(self.var_sep.get())
        if sep is None or self._csv is None or sep == self._csv.separator:
            return
        self._csv = CsvOptions(separator=sep, encoding=self._csv.encoding)
        self._pedir_vista_previa()

    def seleccionar_codificacion(self, codificacion: str) -> None:
        self.var_cod.set(CODIFICACIONES[codificacion])
        self._eligio_codificacion()

    def _eligio_codificacion(self) -> None:
        por_texto = {v: k for k, v in CODIFICACIONES.items()}
        cod = por_texto.get(self.var_cod.get())
        if cod is None or self._csv is None or cod == self._csv.encoding:
            return
        self._csv = CsvOptions(separator=self._csv.separator, encoding=cod)
        self._pedir_vista_previa()

    def poner_encabezado(self, fila: int) -> None:
        """``fila`` en base 1, como la ve el usuario."""
        self.var_encabezado.set(str(fila))
        self._eligio_encabezado()

    def _eligio_encabezado(self) -> None:
        try:
            fila = int(self.var_encabezado.get()) - 1
        except ValueError:
            self.var_encabezado.set(str(self._encabezado_usado + 1))
            return
        if fila < 0:
            self.var_encabezado.set(str(self._encabezado_usado + 1))
            return
        if fila == self._encabezado_usado or fila == self._encabezado:
            return
        self._encabezado = fila
        self._pedir_vista_previa()

    # ---- cierre ----
    def seleccion(self) -> Seleccion | None:
        if self.path is None:
            return None
        return Seleccion(
            path=self.path,
            sheet=self._hoja,
            csv_options=self._csv if es_csv(self.path) else None,
            header_row=self._encabezado_usado,
        )

    def aceptar(self) -> None:
        elegido = self.seleccion()
        if elegido is None or not self.listo:
            return
        previas = load_preferences()
        save_preferences(
            Preferences(
                last_folder=elegido.path.parent,
                last_csv_options=elegido.csv_options or previas.last_csv_options,
                last_destination=previas.last_destination,
            )
        )
        self._cerrar()
        self.on_accept(elegido)

    def cancelar(self) -> None:
        self._cerrar()

    def _cerrar(self) -> None:
        self.cerrada = True
        self._version += 1
        try:
            self.destroy()
        except tk.TclError:
            pass


def abrir_selector(
    parent: tk.Misc, on_accept: Callable[[Seleccion], None], path: Path | None = None
) -> SelectorArchivo:
    """Abre el selector. ``on_accept`` recibe la elección cuando el usuario acepta."""
    return SelectorArchivo(parent, on_accept, path)
