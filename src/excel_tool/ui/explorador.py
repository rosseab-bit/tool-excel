"""Ventana del explorador: pestañas General y Columnas (PLAN §3.1).

Lee el archivo y analiza el perfil en segundo plano (``ui/worker.py``). La UI no calcula:
solo muestra lo que trae el perfil del core.
"""

from __future__ import annotations

import logging
import queue
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from excel_tool.core.io import ReadOptions, ReadResult, list_sheets, read_table
from excel_tool.core.progress import CancelToken, ProgressCallback
from excel_tool.ui import exportar, guardar
from excel_tool.ui.datos import TIPOS_QUE_SIRVEN, PanelDatos
from excel_tool.ui.distintos import PanelDistintos, mensaje_de
from excel_tool.ui.formato import (
    MAX_AVISOS,
    datos_generales,
    detalle_columna,
    fila_columna,
    formatear_numero,
    texto_hoja,
)
from excel_tool.ui.selector import Seleccion
from excel_tool.ui.worker import run_in_background

log = logging.getLogger(__name__)

INTERVALO_COLUMNAS_MS = 100
TEXTO_ANALIZANDO = "Analizando el archivo…"
TEXTO_ERROR = "No se pudo analizar el archivo."
TEXTO_SIN_AVISOS = "No hay avisos: no se encontró nada raro."

ENCABEZADOS = {
    "pos": ("Columna", 80),
    "nombre": ("Nombre", 200),
    "tipo": ("Tipo", 120),
    "datos": ("Con datos", 90),
    "vacias": ("Vacías", 80),
    "distintos": ("Valores distintos", 120),
    "marcas": ("Marcas", 260),
}
COLUMNAS_NUMERICAS = {"datos", "vacias", "distintos"}


def perfilar(leido: Any, **kwargs: Any) -> Any:
    """Llama a ``profile_table`` del core. El import va acá para poder reemplazarla en tests."""
    from excel_tool.core.profile import profile_table

    return profile_table(leido, **kwargs)


def perfil_de_tabla(df: Any, titulo: str, **kwargs: Any) -> Any:
    """Perfil de una tabla que ya está en memoria y no viene de un archivo (por ejemplo, el
    resultado de unir). El perfil queda con el nombre `titulo` y sin datos de archivo."""
    leido = ReadResult(df, Path(titulo), "csv", None, [], None, 0, {}, [])
    return perfilar(leido, display_name=titulo, **kwargs)


DESTINOS_SEMAFORO = (("Excel", "excel", "xlsx"), ("Google Sheets", "google_sheets", "csv_google"))
TITULO_SEMAFORO = "¿Entra en Excel o en Google Sheets?"
TEXTO_DIVIDIR = "Dividir y guardar…"


def _limites() -> Any:
    """Módulo `core.limits` (import perezoso, reemplazable en tests). None si todavía no existe."""
    try:
        from excel_tool.core import limits
    except ImportError:
        log.warning("core.limits todavía no existe: no se muestra el semáforo.")
        return None
    return limits


def calcular_semaforo(df: Any) -> list[tuple[str, Any]] | None:
    """Lista de (destino, resultado de `check`) para Excel y Google Sheets, o None si no se pudo."""
    if df is None:
        return None
    limits = _limites()
    if limits is None:
        return None
    try:
        est = limits.estimate(df)
        return [
            (nombre, limits.check(est, destino, formato))
            for nombre, destino, formato in DESTINOS_SEMAFORO
        ]
    except Exception:
        log.exception("No se pudo calcular el semáforo de tamaño")
        return None


def _clave_orden(clave: str, col: Any) -> Any:
    """Clave para ordenar la tabla por un encabezado (solo ordena, no calcula)."""
    if clave == "pos":
        return col.position
    if clave == "nombre":
        return col.name.lower()
    if clave == "tipo":
        return fila_columna(col)[2].lower()
    if clave == "datos":
        return col.non_empty
    if clave == "vacias":
        return col.empty
    if clave == "distintos":
        return col.distinct
    return fila_columna(col)[6].lower()


class Explorador(tk.Toplevel):
    """Ventana con las pestañas del explorador. Se cierra sola si se cancela el análisis."""

    def __init__(
        self,
        parent: tk.Misc,
        seleccion: Seleccion | None,
        on_cerrar: Callable[[], None] | None = None,
        *,
        tabla: Any = None,
        titulo: str | None = None,
        perfil: Any = None,
    ) -> None:
        super().__init__(parent)
        if seleccion is None and tabla is None:
            raise ValueError("Hace falta un archivo o una tabla.")
        self.seleccion = seleccion
        self.tabla_dada = tabla  # tabla ya cargada: no se vuelve a leer el archivo
        self.perfil_dado = perfil
        self.nombre = titulo or (seleccion.path.name if seleccion is not None else "Tabla")
        self.hay_archivo = seleccion is not None
        self.on_cerrar = on_cerrar
        self.on_ver_distintos: Callable[[str], None] | None = None
        self.perfil: Any = None
        self.cerrada = False
        self.analizando = False
        self._version = 0
        self._token: CancelToken | None = None
        self._cola: queue.Queue[tuple[int, Any]] = queue.Queue()
        self._columnas: dict[str, Any] = {}
        self._orden: tuple[str, bool] | None = None
        self._hoja_actual: str | None = seleccion.sheet if seleccion else None
        self._hoja_pedida: str | None = self._hoja_actual
        self._header_row: int | None = seleccion.header_row if seleccion else None
        self._hojas: list[Any] = []
        self.tabla_frecuentes: ttk.Treeview | None = None
        self._perfil_en_paneles: Any = None
        self.df: Any = None  # la tabla ya leída (no se copia)
        self.semaforo: list[tuple[str, Any]] | None = None

        self.title(f"Información de {self.nombre}")
        self.minsize(820, 560)
        self.protocol("WM_DELETE_WINDOW", self.cerrar)
        self._armar()
        self.after_idle(lambda: self.analizar(self._hoja_actual))

    # ---- armado ----
    def _armar(self) -> None:
        marco = ttk.Frame(self, padding=12)
        marco.pack(fill="both", expand=True)
        arriba = ttk.Frame(marco)
        arriba.pack(fill="x")
        self.etiqueta_archivo = ttk.Label(arriba, text=self.nombre, font="TkHeadingFont")
        self.etiqueta_archivo.pack(side="left")
        self.var_hoja = tk.StringVar(master=self.etiqueta_archivo)
        self.combo_hoja = ttk.Combobox(
            arriba, textvariable=self.var_hoja, state="readonly", width=44
        )
        self.combo_hoja.bind("<<ComboboxSelected>>", lambda _e: self._eligio_hoja())
        self.etiqueta_hoja = ttk.Label(arriba, text="Hoja")
        self.etiqueta_estado = ttk.Label(marco, text=TEXTO_ANALIZANDO)
        self.etiqueta_estado.pack(anchor="w", pady=(4, 6))

        self.pestanas = ttk.Notebook(marco)
        self.pestanas.pack(fill="both", expand=True)
        self.pestana_general = ttk.Frame(self.pestanas, padding=12)
        self.pestana_columnas = ttk.Frame(self.pestanas, padding=8)
        self.pestanas.add(self.pestana_general, text="General")
        self.pestanas.add(self.pestana_columnas, text="Columnas")
        self.pestana_distintos = ttk.Frame(self.pestanas, padding=8)
        self.pestana_datos = ttk.Frame(self.pestanas, padding=8)
        self.pestanas.add(self.pestana_distintos, text="Valores distintos")
        self.pestanas.add(self.pestana_datos, text="Datos")

        self._armar_general()
        self._armar_columnas()
        self.panel_datos = PanelDatos(self.pestana_datos, self)
        self.panel_datos.pack(fill="both", expand=True)
        self.panel_distintos = PanelDistintos(
            self.pestana_distintos, self, on_ver_filas=self._ver_filas
        )
        self.panel_distintos.pack(fill="both", expand=True)
        self.on_ver_distintos = self.ir_a_distintos

    def agregar_pestana(self, marco: ttk.Frame, texto: str) -> None:
        """Para F2-14: suma una de las pestañas reservadas."""
        self.pestanas.add(marco, text=texto)

    def _armar_general(self) -> None:
        self.contenido_general = ttk.Frame(self.pestana_general)
        self.contenido_general.pack(fill="both", expand=True)
        self.boton_informe = ttk.Button(
            self.pestana_general,
            text="Guardar informe del archivo…",
            command=self.guardar_informe,
            state="disabled",
        )
        self.boton_informe.pack(side="bottom", anchor="e", pady=(8, 0))

    def _armar_columnas(self) -> None:
        panel = ttk.PanedWindow(self.pestana_columnas, orient="vertical")
        panel.pack(fill="both", expand=True)
        arriba = ttk.Frame(panel)
        panel.add(arriba, weight=3)
        self.tabla = ttk.Treeview(
            arriba, columns=list(ENCABEZADOS), show="headings", selectmode="browse"
        )
        for clave, (texto, ancho) in ENCABEZADOS.items():
            self.tabla.heading(clave, text=texto, command=lambda c=clave: self._ordenar(c))
            self.tabla.column(
                clave,
                width=ancho,
                anchor="e" if clave in COLUMNAS_NUMERICAS else "w",
                stretch=clave == "marcas",
            )
        barra = ttk.Scrollbar(arriba, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=barra.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.tabla.bind("<<TreeviewSelect>>", lambda _e: self._mostrar_detalle())

        self.detalle = ttk.LabelFrame(panel, text="Detalle de la columna", padding=8)
        panel.add(self.detalle, weight=2)
        self.contenido_detalle = ttk.Frame(self.detalle)
        self.contenido_detalle.pack(fill="both", expand=True)
        self.boton_distintos = ttk.Button(
            self.detalle,
            text="Ver valores distintos de esta columna",
            command=self._ver_distintos,
        )
        self._limpiar_detalle()

    # ---- análisis en segundo plano ----
    def analizar(self, hoja: str | None) -> None:
        """Lee y analiza el archivo (la hoja pedida) en segundo plano."""
        if self.cerrada:
            return
        if self._token is not None:
            self._token.cancel()
        self._version += 1
        version = self._version
        self._hoja_pedida = hoja
        self.analizando = True
        self.etiqueta_estado.configure(text=TEXTO_ANALIZANDO)
        self.combo_hoja.state(["disabled"])
        self._vaciar_columnas()
        self._cola = queue.Queue()
        cola = self._cola
        tabla: dict[str, Any] = {}
        if self.tabla_dada is not None:
            trabajo = self._trabajo_con_tabla(tabla, cola, version)
        else:
            trabajo = self._trabajo_leyendo(tabla, cola, version, hoja)

        self._token = run_in_background(
            self,
            trabajo,
            title="Analizando el archivo",
            on_success=lambda perfil: self._termino(
                version, perfil, tabla.get("df"), tabla.get("semaforo")
            ),
            on_error=lambda exc: self._fallo(version, exc),
            on_cancel=lambda: self._cancelo(version),
            modal=False,
        )
        self.after(INTERVALO_COLUMNAS_MS, lambda: self._drenar(version))

    def _trabajo_leyendo(
        self,
        tabla: dict[str, Any],
        cola: queue.Queue[tuple[int, Any]],
        version: int,
        hoja: str | None,
    ) -> Callable[[ProgressCallback, CancelToken], Any]:
        assert self.seleccion is not None
        opciones = ReadOptions(
            sheet=hoja,
            header_row=self._header_row if hoja == self._hoja_actual else None,
            csv=self.seleccion.csv_options,
        )
        ruta = self.seleccion.path

        def trabajo(progress: ProgressCallback, cancel: CancelToken) -> Any:
            leido = read_table(ruta, opciones)
            tabla["df"] = leido.df
            cancel.check()
            tabla["semaforo"] = calcular_semaforo(leido.df)
            hojas = list_sheets(ruta) if leido.sheets else None
            return perfilar(
                leido,
                sheets=hojas,
                on_column=lambda col: cola.put((version, col)),
                progress=progress,
                cancel=cancel,
            )

        return trabajo

    def _trabajo_con_tabla(
        self, tabla: dict[str, Any], cola: queue.Queue[tuple[int, Any]], version: int
    ) -> Callable[[ProgressCallback, CancelToken], Any]:
        """La tabla ya está cargada: no se lee nada. Si no vino el perfil, se arma acá."""
        df, perfil_dado, titulo = self.tabla_dada, self.perfil_dado, self.nombre

        def trabajo(progress: ProgressCallback, cancel: CancelToken) -> Any:
            tabla["df"] = df
            cancel.check()
            tabla["semaforo"] = calcular_semaforo(df)
            if perfil_dado is not None:
                return perfil_dado
            return perfil_de_tabla(
                df,
                titulo,
                on_column=lambda col: cola.put((version, col)),
                progress=progress,
                cancel=cancel,
            )

        return trabajo

    def cancelar_analisis(self) -> None:
        if self._token is not None:
            self._token.cancel()

    def _drenar(self, version: int) -> None:
        """Pasa a la tabla las columnas que fue terminando el hilo de trabajo."""
        if self.cerrada or version != self._version or not self.analizando:
            return
        self._tomar_columnas(version)
        self.after(INTERVALO_COLUMNAS_MS, lambda: self._drenar(version))

    def _tomar_columnas(self, version: int) -> None:
        try:
            while True:
                v, col = self._cola.get_nowait()
                if v == version:
                    self._agregar_columna(col)
        except queue.Empty:
            pass

    def _termino(self, version: int, perfil: Any, df: Any = None, semaforo: Any = None) -> None:
        if self.cerrada or version != self._version:
            return
        self.analizando = False
        self.perfil = perfil
        self.df = df
        self.semaforo = semaforo
        self._hoja_actual = perfil.sheet
        self._header_row = perfil.header_row
        self.mostrar_perfil(perfil)

    def _cancelo(self, version: int) -> None:
        if self.cerrada or version != self._version:
            return
        self.analizando = False
        if self.perfil is None:
            self.cerrar()
        else:
            self._volver_a_hoja_anterior()

    def _fallo(self, version: int, exc: BaseException) -> None:
        if self.cerrada or version != self._version:
            return
        self.analizando = False
        mensaje = getattr(exc, "mensaje", None) or TEXTO_ERROR
        messagebox.showerror("No se pudo analizar", mensaje, parent=self)
        if self.perfil is None:
            self.cerrar()
        else:
            self._volver_a_hoja_anterior()

    def _volver_a_hoja_anterior(self) -> None:
        self.mostrar_perfil(self.perfil)

    # ---- mostrar el perfil ----
    def mostrar_perfil(self, perfil: Any) -> None:
        self.perfil = perfil
        self.etiqueta_estado.configure(
            text=(
                f"{formatear_numero(perfil.rows)} filas y "
                f"{formatear_numero(perfil.columns)} columnas"
            )
        )
        self._llenar_hojas(perfil)
        self._llenar_general(perfil)
        self._vaciar_columnas()
        for col in perfil.columns_profile:
            self._agregar_columna(col)
        self.boton_informe.state(["!disabled"])
        if perfil is not self._perfil_en_paneles:
            self._perfil_en_paneles = perfil
            self.panel_distintos.cargar(self.df, perfil)
            self.panel_datos.cargar(self.df, perfil)

    def _llenar_hojas(self, perfil: Any) -> None:
        self._hojas = list(perfil.sheets or [])
        textos = [texto_hoja(h) for h in self._hojas]
        self.combo_hoja.configure(values=textos)
        self.combo_hoja.state(["!disabled"])
        self.etiqueta_hoja.pack_forget()
        self.combo_hoja.pack_forget()
        if len(self._hojas) > 1:
            self.combo_hoja.pack(side="right")
            self.etiqueta_hoja.pack(side="right", padx=(0, 6))
        for h, texto in zip(self._hojas, textos, strict=False):
            if h.name == perfil.sheet:
                self.var_hoja.set(texto)

    def _llenar_general(self, perfil: Any) -> None:
        for hijo in self.contenido_general.winfo_children():
            hijo.destroy()
        secciones = datos_generales(perfil)
        if not self.hay_archivo:
            secciones = secciones[1:]  # una tabla en memoria no tiene datos de archivo
        for titulo, filas in secciones:
            caja = ttk.LabelFrame(self.contenido_general, text=titulo, padding=8)
            caja.pack(fill="x", pady=(0, 8))
            caja.columnconfigure(1, weight=1)
            for n, (etiqueta, valor) in enumerate(filas):
                ttk.Label(caja, text=etiqueta).grid(row=n, column=0, sticky="nw", padx=(0, 16))
                ttk.Label(caja, text=valor, wraplength=520).grid(row=n, column=1, sticky="w")
        self._llenar_semaforo()
        avisos = ttk.LabelFrame(self.contenido_general, text="Avisos", padding=8)
        avisos.pack(fill="x")
        lista = list(perfil.warnings)
        if not lista:
            ttk.Label(avisos, text=TEXTO_SIN_AVISOS).pack(anchor="w")
        for aviso in lista[:MAX_AVISOS]:
            ttk.Label(avisos, text=f"• {aviso}", wraplength=640, justify="left").pack(anchor="w")
        if len(lista) > MAX_AVISOS:
            resto = formatear_numero(len(lista) - MAX_AVISOS)
            ttk.Label(avisos, text=f"…y {resto} avisos más.").pack(anchor="w")

    def _llenar_semaforo(self) -> None:
        if not self.semaforo:
            return
        caja = ttk.LabelFrame(self.contenido_general, text=TITULO_SEMAFORO, padding=8)
        caja.pack(fill="x", pady=(0, 8))
        caja.columnconfigure(1, weight=1)
        self.lineas_semaforo: list[tuple[tk.Label, ttk.Label]] = []
        no_entra = False
        for n, (nombre, chk) in enumerate(self.semaforo):
            etiqueta, fondo, frente = guardar.NIVELES[chk.level]
            marca = tk.Label(
                caja,
                text=f"●  {nombre}: {etiqueta}",
                bg=fondo,
                fg=frente,
                anchor="w",
                padx=8,
                width=44,
            )
            marca.grid(row=n, column=0, sticky="w", pady=2)
            motivos = list(chk.reasons)
            motivo = ttk.Label(
                caja,
                text=motivos[0] if motivos else guardar.TEXTO_OK,
                wraplength=420,
                justify="left",
            )
            motivo.grid(row=n, column=1, sticky="w", padx=(12, 0))
            self.lineas_semaforo.append((marca, motivo))
            no_entra = no_entra or chk.level == "obligatorio"
        self.boton_dividir = None
        if no_entra:
            self.boton_dividir = ttk.Button(
                caja, text=TEXTO_DIVIDIR, command=self.dividir_y_guardar
            )
            self.boton_dividir.grid(row=len(self.semaforo), column=0, sticky="w", pady=(6, 0))

    def dividir_y_guardar(self) -> None:
        """Abre el paso Guardar con toda la tabla (sin copiarla)."""
        if self.df is None or self.perfil is None:
            return
        datos = self.panel_datos
        tipos = {
            c.name: c.detected_type
            for c in self.perfil.columns_profile
            if c.detected_type in TIPOS_QUE_SIRVEN
        }
        guardar.abrir_guardar(
            self,
            self.df,
            nombre_sugerido=Path(self.perfil.file.path).stem,
            claves_sugeridas=list(datos.claves),
            column_types=tipos,
        )

    def _vaciar_columnas(self) -> None:
        self.tabla.delete(*self.tabla.get_children())
        self._columnas = {}
        self._limpiar_detalle()

    def _agregar_columna(self, col: Any) -> None:
        iid = str(col.position)
        if iid in self._columnas:
            return
        self._columnas[iid] = col
        self.tabla.insert("", "end", iid=iid, values=fila_columna(col))

    def _ordenar(self, clave: str) -> None:
        descendente = self._orden == (clave, True)
        self._orden = (clave, not descendente)
        ids = sorted(
            self._columnas,
            key=lambda i: _clave_orden(clave, self._columnas[i]),
            reverse=descendente,
        )
        for n, iid in enumerate(ids):
            self.tabla.move(iid, "", n)

    # ---- detalle de una columna ----
    def _limpiar_detalle(self) -> None:
        for hijo in self.contenido_detalle.winfo_children():
            hijo.destroy()
        self.boton_distintos.pack_forget()
        ttk.Label(self.contenido_detalle, text="Elegí una columna para ver su detalle.").pack(
            anchor="w"
        )

    def columna_elegida(self) -> Any:
        seleccion = self.tabla.selection()
        return self._columnas.get(seleccion[0]) if seleccion else None

    def seleccionar_columna(self, nombre_o_posicion: str | int) -> None:
        for iid, col in self._columnas.items():
            if col.name == nombre_o_posicion or col.position == nombre_o_posicion:
                self.tabla.selection_set(iid)
                self.tabla.see(iid)
                self._mostrar_detalle()
                return

    def _mostrar_detalle(self) -> None:
        col = self.columna_elegida()
        if col is None:
            return
        for hijo in self.contenido_detalle.winfo_children():
            hijo.destroy()
        izquierda = ttk.Frame(self.contenido_detalle)
        izquierda.pack(side="left", fill="both", expand=True)
        ttk.Label(izquierda, text=col.name, font="TkHeadingFont").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )
        fila = 1
        for etiqueta, valor in detalle_columna(col):
            ttk.Label(izquierda, text=etiqueta).grid(row=fila, column=0, sticky="nw", padx=(0, 12))
            ttk.Label(izquierda, text=valor, wraplength=380, justify="left").grid(
                row=fila, column=1, sticky="w"
            )
            fila += 1
        for nota in col.notes:
            ttk.Label(izquierda, text=f"• {nota}", wraplength=480, justify="left").grid(
                row=fila, column=0, columnspan=2, sticky="w"
            )
            fila += 1
        if col.top_values:
            derecha = ttk.Frame(self.contenido_detalle)
            derecha.pack(side="right", fill="y", padx=(12, 0))
            ttk.Label(derecha, text="Valores más frecuentes").pack(anchor="w")
            top = ttk.Treeview(derecha, columns=("valor", "cantidad"), show="headings", height=6)
            top.heading("valor", text="Valor")
            top.heading("cantidad", text="Veces")
            top.column("valor", width=180, anchor="w")
            top.column("cantidad", width=80, anchor="e")
            for valor, cantidad in col.top_values:
                top.insert(
                    "",
                    "end",
                    values=(valor if valor != "" else "(vacío)", formatear_numero(cantidad)),
                )
            top.pack(fill="y")
            self.tabla_frecuentes = top
        self.boton_distintos.pack_forget()
        if self.on_ver_distintos is not None:
            self.boton_distintos.pack(
                side="bottom", anchor="w", pady=(6, 0), before=self.contenido_detalle
            )

    def _ver_distintos(self) -> None:
        col = self.columna_elegida()
        if col is not None and self.on_ver_distintos is not None:
            self.on_ver_distintos(col.name)

    # ---- valores distintos y datos ----
    def ir_a_distintos(self, columna: str) -> None:
        """Va a la pestaña Valores distintos con la columna elegida y calcula."""
        self.pestanas.select(self.pestana_distintos)
        self.panel_distintos.elegir_columnas([columna])
        self.panel_distintos.calcular()

    def _ver_filas(self, columnas: list[str], valores: tuple[str, ...], normalizar: bool) -> None:
        self.panel_datos.mostrar_valor(columnas, valores, normalizar)
        self.pestanas.select(self.pestana_datos)

    def guardar_informe(self) -> None:
        """Guarda el informe del archivo con los valores distintos calculados hasta ahora."""
        if self.perfil is None:
            return
        perfil = self.perfil
        elegido = filedialog.asksaveasfilename(
            parent=self,
            title="Guardar informe del archivo",
            defaultextension=".xlsx",
            filetypes=[("Excel (.xlsx)", "*.xlsx")],
            initialdir=str(Path(perfil.file.path).parent),
            initialfile=f"Informe de {Path(perfil.file.path).stem}"[:80],
        )
        if not elegido:
            return
        ruta = Path(elegido)
        distintos = list(self.panel_distintos.calculados.values())

        def trabajo(progress: ProgressCallback, cancel: CancelToken) -> Any:
            return exportar.guardar_informe(perfil, distintos, ruta)

        run_in_background(
            self,
            trabajo,
            title="Guardando el informe",
            on_success=lambda _r: messagebox.showinfo(
                "Listo", f"El informe quedó guardado en:\n{ruta}", parent=self
            ),
            on_error=lambda exc: messagebox.showerror(
                "No se pudo guardar", mensaje_de(exc), parent=self
            ),
            cancellable=False,
        )

    # ---- hoja ----
    def cambiar_hoja(self, nombre: str) -> None:
        for h in self._hojas:
            if h.name == nombre:
                self.var_hoja.set(texto_hoja(h))
        self._eligio_hoja()

    def _eligio_hoja(self) -> None:
        textos = [texto_hoja(h) for h in self._hojas]
        actual = self.var_hoja.get()
        if actual not in textos:
            return
        nombre = self._hojas[textos.index(actual)].name
        if nombre == self._hoja_actual and not self.analizando:
            return
        self.analizar(nombre)

    # ---- cierre ----
    def cerrar(self) -> None:
        if self.cerrada:
            return
        self.cerrada = True
        self._version += 1
        self.cancelar_analisis()
        try:
            self.destroy()
        except tk.TclError:
            pass
        if self.on_cerrar is not None:
            self.on_cerrar()


def abrir_explorador(
    parent: tk.Misc,
    seleccion: Seleccion | None = None,
    on_cerrar: Callable[[], None] | None = None,
    *,
    tabla: Any = None,
    titulo: str | None = None,
    perfil: Any = None,
) -> Explorador:
    """Abre la ventana del explorador y empieza a analizar.

    Con `tabla` (un DataFrame ya cargado) no se vuelve a leer ningún archivo: se usa esa tabla y
    el `perfil` si viene (si no, se arma en segundo plano). Sin `seleccion` la tabla no viene de
    un archivo (por ejemplo, el resultado de unir) y `titulo` es el nombre que se muestra.
    """
    return Explorador(parent, seleccion, on_cerrar, tabla=tabla, titulo=titulo, perfil=perfil)
