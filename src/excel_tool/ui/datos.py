"""Pestaña "Datos" del explorador: grilla paginada con filtros y orden (PLAN §3.1).

No calcula nada: arma las condiciones y llama al core (``filter_rows``, ``sort_rows``,
``rows_with_value``, ``get_page``). Filtrar y ordenar corren en segundo plano.
"""

from __future__ import annotations

import logging
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any

import numpy as np
import pandas as pd

from excel_tool.core.explore import (
    FilterCondition,
    filter_rows,
    get_page,
    rows_with_value,
    sort_rows,
)
from excel_tool.ui import guardar
from excel_tool.ui.distintos import mensaje_de
from excel_tool.ui.formato import formatear_numero
from excel_tool.ui.grilla import COLUMNA_FILA, Grilla
from excel_tool.ui.worker import run_in_background

log = logging.getLogger(__name__)

OPERACIONES = {
    "es igual a": "igual",
    "es distinto de": "distinto",
    "contiene": "contiene",
    "está vacío": "vacio",
    "no está vacío": "no_vacio",
    "es mayor que": "mayor",
    "es menor que": "menor",
}
SIN_VALOR = {"vacio", "no_vacio"}
TIPOS_QUE_SIRVEN = {"entero", "decimal", "fecha", "fecha_hora", "si_no"}
TEXTO_EXPORTAR = "Exportar lo filtrado…"
TEXTO_GUARDAR_TODO = "Guardar estas filas…"
PREGUNTA_OCULTAS = (
    "Hay columnas ocultas en la tabla.\n\n"
    "¿Querés guardar todas las columnas?\n\n"
    "Sí: se guardan todas, también las ocultas.\n"
    "No: se guardan solo las que están a la vista."
)


class FilaCondicion:
    """Una condición del filtro: columna, operación y valor."""

    def __init__(self, panel: PanelDatos, nombres: list[str]) -> None:
        self.marco = ttk.Frame(panel.contenedor_condiciones)
        self.marco.pack(fill="x", pady=1)
        self.var_columna = tk.StringVar()
        self.var_operacion = tk.StringVar(value="es igual a")
        self.var_valor = tk.StringVar()
        self.combo_columna = ttk.Combobox(
            self.marco, textvariable=self.var_columna, values=nombres, state="readonly", width=24
        )
        self.combo_operacion = ttk.Combobox(
            self.marco,
            textvariable=self.var_operacion,
            values=list(OPERACIONES),
            state="readonly",
            width=16,
        )
        self.entrada = ttk.Entry(self.marco, textvariable=self.var_valor, width=24)
        self.boton_quitar = ttk.Button(
            self.marco, text="Quitar", width=8, command=lambda: panel.quitar_condicion(self)
        )
        self.combo_columna.pack(side="left", padx=(0, 4))
        self.combo_operacion.pack(side="left", padx=(0, 4))
        self.entrada.pack(side="left", padx=(0, 4), fill="x", expand=True)
        self.boton_quitar.pack(side="left")
        self.combo_operacion.bind("<<ComboboxSelected>>", lambda _e: self._ajustar())
        self.entrada.bind("<Return>", lambda _e: panel.aplicar())
        if nombres:
            self.var_columna.set(nombres[0])

    def _ajustar(self) -> None:
        sin_valor = OPERACIONES.get(self.var_operacion.get()) in SIN_VALOR
        self.entrada.configure(state="disabled" if sin_valor else "normal")

    def condicion(self) -> FilterCondition | None:
        columna = self.var_columna.get()
        if not columna:
            return None
        op = OPERACIONES[self.var_operacion.get()]
        valor = "" if op in SIN_VALOR else self.var_valor.get()
        return FilterCondition(columna, op, valor)  # type: ignore[arg-type]


class PanelDatos(ttk.Frame):
    def __init__(self, master: tk.Misc, ventana: tk.Misc, page_size: int | None = None) -> None:
        super().__init__(master)
        self.ventana = ventana
        self.df: Any = None
        self.tipos: dict[str, str] = {}
        self.nombres: list[str] = []
        self.filas: pd.Index = pd.RangeIndex(0)
        self.condiciones: list[FilaCondicion] = []
        self.aviso_valor: str | None = None
        self.nombre_archivo = "resultado"
        self.claves: list[str] = []
        self._orden: tuple[str, bool] | None = None
        self._version = 0
        self._page_size = page_size
        self._armar()

    def _armar(self) -> None:
        filtros = ttk.LabelFrame(self, text="Filtrar filas", padding=6)
        filtros.pack(fill="x")
        self.contenedor_condiciones = ttk.Frame(filtros)
        self.contenedor_condiciones.pack(fill="x")
        botones = ttk.Frame(filtros)
        botones.pack(fill="x", pady=(4, 0))
        self.boton_agregar = ttk.Button(
            botones, text="Agregar condición", command=self.agregar_condicion
        )
        self.boton_agregar.pack(side="left")
        self.boton_aplicar = ttk.Button(botones, text="Aplicar", command=self.aplicar)
        self.boton_aplicar.pack(side="left", padx=6)
        self.boton_exportar = ttk.Button(
            botones, text=TEXTO_GUARDAR_TODO, command=self.exportar, state="disabled"
        )
        self.boton_exportar.pack(side="right", padx=(8, 0))
        self.etiqueta_conteo = ttk.Label(botones, text="")
        self.etiqueta_conteo.pack(side="right")

        self.marco_aviso = ttk.Frame(self)
        self.etiqueta_aviso = ttk.Label(self.marco_aviso, text="", wraplength=640)
        self.etiqueta_aviso.pack(side="left")
        self.boton_quitar_filtro = ttk.Button(
            self.marco_aviso, text="Quitar filtro", command=self.quitar_filtro_valor
        )
        self.boton_quitar_filtro.pack(side="left", padx=8)

        kwargs = {"page_size": self._page_size} if self._page_size else {}
        self.grilla = Grilla(self, None, on_sort=self._ordenar, **kwargs)
        self.grilla.pack(fill="both", expand=True, pady=(6, 0))

    # ---- datos ----
    def cargar(self, df: Any, perfil: Any) -> None:
        """Tabla nueva: se muestran todas las filas, sin filtros ni orden."""
        self._version += 1
        self.df = df
        self.nombres = [c.name for c in perfil.columns_profile]
        self.tipos = {c.name: c.detected_type for c in perfil.columns_profile}
        self._orden = None
        self.claves = [c.name for c in perfil.columns_profile if getattr(c, "is_unique", False)]
        try:
            self.nombre_archivo = Path(perfil.file.path).stem or "resultado"
        except (AttributeError, TypeError):
            self.nombre_archivo = "resultado"
        for fila in list(self.condiciones):
            self.quitar_condicion(fila)
        self._poner_aviso(None)
        if df is None:
            self.filas = pd.RangeIndex(0)
            self.grilla.set_proveedor(None)
            self._actualizar_conteo()
            return
        self._poner_filas(pd.RangeIndex(len(df)))

    def _poner_filas(self, filas: pd.Index) -> None:
        self.filas = filas
        df = self.df
        self.grilla.set_proveedor(
            lambda pagina: (get_page(df, filas, pagina, self.grilla.page_size), len(filas))
        )
        self._actualizar_conteo()

    def _actualizar_conteo(self) -> None:
        total = 0 if self.df is None else len(self.df)
        self.etiqueta_conteo.configure(
            text=(
                f"Mostrando {formatear_numero(len(self.filas))} de {formatear_numero(total)} filas"
            )
        )
        self.boton_exportar.configure(
            text=TEXTO_EXPORTAR if len(self.filas) < total else TEXTO_GUARDAR_TODO
        )
        self.boton_exportar.state(["!disabled"] if self.df is not None else ["disabled"])

    def _poner_aviso(self, texto: str | None) -> None:
        self.aviso_valor = texto
        if texto is None:
            self.marco_aviso.pack_forget()
            return
        self.etiqueta_aviso.configure(text=texto)
        self.marco_aviso.pack(fill="x", pady=(6, 0), before=self.grilla)

    # ---- condiciones ----
    def agregar_condicion(
        self, columna: str | None = None, operacion: str | None = None, valor: str = ""
    ) -> FilaCondicion:
        fila = FilaCondicion(self, self.nombres)
        if columna is not None:
            fila.var_columna.set(columna)
        if operacion is not None:
            fila.var_operacion.set(operacion)
            fila._ajustar()
        fila.var_valor.set(valor)
        self.condiciones.append(fila)
        return fila

    def quitar_condicion(self, fila: FilaCondicion) -> None:
        if fila in self.condiciones:
            self.condiciones.remove(fila)
        fila.marco.destroy()

    def lista_condiciones(self) -> list[FilterCondition]:
        return [c for c in (f.condicion() for f in self.condiciones) if c is not None]

    # ---- filtrar / ordenar (segundo plano) ----
    def _correr(
        self,
        titulo: str,
        calcular: Callable[[], pd.Index],
        al_terminar: Callable[[pd.Index], None],
    ) -> None:
        version = self._version
        self.boton_aplicar.state(["disabled"])

        def trabajo(progress, cancel):
            return calcular()

        def listo(filas):
            self.boton_aplicar.state(["!disabled"])
            if version == self._version:
                al_terminar(filas)

        def fallo(exc):
            self.boton_aplicar.state(["!disabled"])
            if version == self._version:
                messagebox.showerror("Datos", mensaje_de(exc), parent=self.ventana)

        run_in_background(
            self.ventana,
            trabajo,
            title=titulo,
            on_success=listo,
            on_error=fallo,
            on_cancel=lambda: self.boton_aplicar.state(["!disabled"]),
            cancellable=False,
        )

    def _con_orden(self, filas: pd.Index) -> pd.Index:
        """Reaplica el orden elegido (si hay) sobre las filas. Corre en el hilo de trabajo."""
        if self._orden is None:
            return filas
        columna, ascendente = self._orden
        return sort_rows(self.df, filas, columna, ascendente, self.tipos.get(columna, "texto"))

    def aplicar(self) -> None:
        if self.df is None:
            return
        df, tipos, condiciones = self.df, dict(self.tipos), self.lista_condiciones()

        def calcular() -> pd.Index:
            return self._con_orden(filter_rows(df, condiciones, tipos))

        def terminar(filas: pd.Index) -> None:
            self._poner_aviso(None)
            self._poner_filas(filas)

        self._correr("Filtrando las filas", calcular, terminar)

    def _ordenar(self, columna: str, ascendente: bool) -> None:
        if self.df is None:
            return
        df, filas = self.df, self.filas
        tipo = self.tipos.get(columna, "texto")

        def calcular() -> pd.Index:
            return sort_rows(df, filas, columna, ascendente, tipo)

        def terminar(ordenadas: pd.Index) -> None:
            self._orden = (columna, ascendente)
            self._poner_filas(ordenadas)

        self._correr("Ordenando las filas", calcular, terminar)

    def mostrar_valor(
        self, columnas: list[str], valores: tuple[str, ...], normalizar: bool
    ) -> None:
        """Muestra solo las filas que tienen esos valores (desde Valores distintos)."""
        if self.df is None:
            return
        df = self.df
        texto = "Estás viendo solo las filas donde " + " y ".join(
            f"{c} es «{v}»" for c, v in zip(columnas, valores, strict=True)
        )

        def calcular() -> pd.Index:
            return self._con_orden(rows_with_value(df, columnas, valores, normalizar))

        def terminar(filas: pd.Index) -> None:
            for fila in list(self.condiciones):
                self.quitar_condicion(fila)
            self._poner_aviso(texto + ".")
            self._poner_filas(filas)

        self._correr("Buscando las filas", calcular, terminar)

    def quitar_filtro_valor(self) -> None:
        if self.df is None:
            return
        df = self.df

        def terminar(filas: pd.Index) -> None:
            self._poner_aviso(None)
            self._poner_filas(filas)

        self._correr("Mostrando todas las filas", lambda: self._con_orden_total(df), terminar)

    def _con_orden_total(self, df: Any) -> pd.Index:
        return self._con_orden(pd.RangeIndex(len(df)))

    # ---- exportar lo que se ve ----
    def _son_todas_en_orden(self) -> bool:
        """Las filas son todas y en el orden del archivo: no hace falta copiar nada."""
        return len(self.filas) == len(self.df) and self.filas.equals(pd.RangeIndex(len(self.df)))

    def _columnas_a_guardar(self) -> list[str] | None:
        """None = todas. Si hay columnas ocultas, pregunta si se guardan o no."""
        visibles = [n for n in self.grilla.columnas_visibles if n != COLUMNA_FILA]
        ocultas = [n for n in self.nombres if n not in visibles]
        if not ocultas or not visibles:
            return None
        if messagebox.askyesno("Guardar las filas", PREGUNTA_OCULTAS, parent=self.ventana):
            return None
        return [n for n in self.nombres if n in visibles]

    def exportar(self) -> None:
        """Abre el paso Guardar con las filas filtradas y ordenadas que se ven en la grilla."""
        if self.df is None:
            return
        df, filas = self.df, self.filas
        columnas = self._columnas_a_guardar()
        filtrado = len(filas) < len(df)
        nombre = f"{self.nombre_archivo}_filtrado" if filtrado else self.nombre_archivo
        todas = self._son_todas_en_orden()
        tipos_perfil, claves = dict(self.tipos), list(self.claves)

        def abrir(tabla: Any) -> None:
            nombres = [str(c) for c in tabla.columns]
            tipos = {n: tipos_perfil[n] for n in nombres if tipos_perfil.get(n) in TIPOS_QUE_SIRVEN}
            guardar.abrir_guardar(
                self.ventana,
                tabla,
                nombre_sugerido=nombre,
                claves_sugeridas=[c for c in claves if c in nombres],
                column_types=tipos,
            )

        if todas and columnas is None:
            abrir(df)
            return

        def trabajo(progress, cancel):
            base = df if todas else df.iloc[np.asarray(filas, dtype=np.int64)]
            if columnas is not None:
                base = base[columnas]
            base.index = pd.RangeIndex(len(base))  # copia nueva: se puede renumerar
            return base

        def fallo(exc):
            messagebox.showerror("Guardar las filas", mensaje_de(exc), parent=self.ventana)

        run_in_background(
            self.ventana,
            trabajo,
            title="Preparando las filas",
            on_success=abrir,
            on_error=fallo,
            cancellable=False,
            modal=False,
        )
