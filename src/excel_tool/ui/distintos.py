"""Pestaña "Valores distintos" del explorador (PLAN §3.1).

Solo arma las opciones y llama al core (``value_counts`` y ``rows_with_value``); el cálculo corre
en segundo plano y solo cuando el usuario aprieta "Calcular".
"""

from __future__ import annotations

import logging
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from excel_tool.core.explore import DistinctOptions, value_counts
from excel_tool.ui import exportar
from excel_tool.ui.formato import formatear_numero, formatear_porcentaje
from excel_tool.ui.worker import MENSAJE_INESPERADO, run_in_background

log = logging.getLogger(__name__)

MAX_FILAS_TABLA = 5000
MAX_GRUPOS_VARIANTES = 20
TIPOS_ARCHIVO = [("Excel (.xlsx)", "*.xlsx"), ("CSV (.csv)", "*.csv")]


def formato_de_ruta(ruta: Path) -> str:
    """Formato de salida según la extensión elegida."""
    return "csv_excel" if ruta.suffix.lower() == ".csv" else "xlsx"


def mensaje_de(exc: BaseException) -> str:
    """Mensaje amigable para mostrar (el detalle ya lo registró el worker)."""
    return getattr(exc, "mensaje", None) or MENSAJE_INESPERADO


class PanelDistintos(ttk.Frame):
    def __init__(
        self,
        master: tk.Misc,
        ventana: tk.Misc,
        on_ver_filas: Callable[[list[str], tuple[str, ...], bool], None] | None = None,
    ) -> None:
        super().__init__(master)
        self.ventana = ventana
        self.on_ver_filas = on_ver_filas
        self.df: Any = None
        self.perfil: Any = None
        self.resultado: Any = None
        self.calculados: dict[tuple[str, ...], Any] = {}
        self._normalizado = False
        self._nombres: dict[str, str] = {}
        self._version = 0

        self.var_orden = tk.StringVar(master=self, value="cantidad")
        self.var_vacios = tk.BooleanVar(master=self, value=True)
        self.var_juntar = tk.BooleanVar(master=self, value=False)
        self.var_buscar = tk.StringVar(master=self)
        self._armar()

    def _armar(self) -> None:
        izq = ttk.Frame(self)
        izq.pack(side="left", fill="y", padx=(0, 10))
        ttk.Label(izq, text="Columnas (podés elegir varias)").pack(anchor="w")
        caja = ttk.Frame(izq)
        caja.pack(fill="y", expand=True, pady=(2, 6))
        self.lista = ttk.Treeview(caja, show="tree", selectmode="extended", height=10)
        self.lista.column("#0", width=200)
        barra = ttk.Scrollbar(caja, orient="vertical", command=self.lista.yview)
        self.lista.configure(yscrollcommand=barra.set)
        self.lista.pack(side="left", fill="y", expand=True)
        barra.pack(side="right", fill="y")

        opciones = ttk.LabelFrame(izq, text="Opciones", padding=6)
        opciones.pack(fill="x")
        ttk.Radiobutton(
            opciones, text="Ordenar por cantidad", variable=self.var_orden, value="cantidad"
        ).pack(anchor="w")
        ttk.Radiobutton(
            opciones, text="Ordenar de la A a la Z", variable=self.var_orden, value="alfabetico"
        ).pack(anchor="w")
        ttk.Checkbutton(opciones, text="Incluir celdas vacías", variable=self.var_vacios).pack(
            anchor="w"
        )
        ttk.Checkbutton(
            opciones,
            text="Juntar escrituras parecidas (Norte / NORTE / Norte )",
            variable=self.var_juntar,
        ).pack(anchor="w")
        ttk.Label(opciones, text="Buscar un texto").pack(anchor="w", pady=(4, 0))
        self.entrada_buscar = ttk.Entry(opciones, textvariable=self.var_buscar, width=28)
        self.entrada_buscar.pack(fill="x")
        self.entrada_buscar.bind("<Return>", lambda _e: self.calcular())

        self.boton_calcular = ttk.Button(izq, text="Calcular", command=self.calcular)
        self.boton_calcular.pack(fill="x", pady=(8, 0))

        der = ttk.Frame(self)
        der.pack(side="left", fill="both", expand=True)
        self.etiqueta_resumen = ttk.Label(
            der, text="Elegí una o más columnas y apretá «Calcular».", wraplength=560
        )
        self.etiqueta_resumen.pack(anchor="w")
        self.etiqueta_aviso = ttk.Label(der, text="", wraplength=560, justify="left")
        self.etiqueta_aviso.pack(anchor="w")
        self.etiqueta_variantes = ttk.Label(der, text="", wraplength=560, justify="left")
        self.etiqueta_variantes.pack(anchor="w")

        pie = ttk.Frame(der)
        pie.pack(side="bottom", fill="x", pady=(6, 0))
        self.boton_ver = ttk.Button(
            pie, text="Ver filas con este valor", command=self.ver_filas, state="disabled"
        )
        self.boton_ver.pack(side="left")
        self.boton_exportar = ttk.Button(
            pie, text="Exportar esta lista…", command=self.exportar, state="disabled"
        )
        self.boton_exportar.pack(side="right")

        cuerpo = ttk.Frame(der)
        cuerpo.pack(fill="both", expand=True, pady=(6, 0))
        self.tabla = ttk.Treeview(cuerpo, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(cuerpo, orient="vertical", command=self.tabla.yview)
        hsb = ttk.Scrollbar(cuerpo, orient="horizontal", command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        cuerpo.rowconfigure(0, weight=1)
        cuerpo.columnconfigure(0, weight=1)
        self.tabla.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.tabla.bind("<Double-1>", lambda _e: self.ver_filas())
        self.tabla.bind("<<TreeviewSelect>>", lambda _e: self._actualizar_botones())

    # ---- datos ----
    def cargar(self, df: Any, perfil: Any) -> None:
        """Se llama cuando hay una tabla nueva (otra hoja u otro archivo): limpia todo."""
        self._version += 1
        self.df = df
        self.perfil = perfil
        self.calculados = {}
        self._nombres = {}
        self.lista.delete(*self.lista.get_children())
        for col in perfil.columns_profile:
            iid = str(col.position)
            self._nombres[iid] = col.name
            self.lista.insert("", "end", iid=iid, text=col.name)
        self._limpiar_resultado()

    def _limpiar_resultado(self) -> None:
        self.resultado = None
        self.tabla.delete(*self.tabla.get_children())
        self.etiqueta_resumen.configure(text="Elegí una o más columnas y apretá «Calcular».")
        self.etiqueta_aviso.configure(text="")
        self.etiqueta_variantes.configure(text="")
        self._actualizar_botones()

    def elegir_columnas(self, nombres: list[str]) -> None:
        iids = [i for i, n in self._nombres.items() if n in nombres]
        self.lista.selection_set(iids)
        if iids:
            self.lista.see(iids[0])

    def columnas_elegidas(self) -> list[str]:
        return [self._nombres[i] for i in self.lista.selection() if i in self._nombres]

    def opciones(self) -> DistinctOptions:
        buscar = self.var_buscar.get().strip()
        return DistinctOptions(
            sort=self.var_orden.get(),  # type: ignore[arg-type]
            include_empty=bool(self.var_vacios.get()),
            normalize=bool(self.var_juntar.get()),
            search=buscar or None,
        )

    # ---- calcular ----
    def calcular(self) -> None:
        if self.df is None:
            return
        columnas = self.columnas_elegidas()
        if not columnas:
            messagebox.showinfo(
                "Valores distintos", "Elegí al menos una columna.", parent=self.ventana
            )
            return
        opciones = self.opciones()
        df = self.df
        version = self._version
        self.boton_calcular.state(["disabled"])

        def trabajo(progress, cancel):
            return value_counts(df, columnas, opciones, progress=progress, cancel=cancel)

        def listo(resultado):
            self.boton_calcular.state(["!disabled"])
            if version == self._version:
                self._mostrar(resultado, opciones.normalize)

        def fallo(exc):
            self.boton_calcular.state(["!disabled"])
            if version == self._version:
                messagebox.showerror("Valores distintos", mensaje_de(exc), parent=self.ventana)

        run_in_background(
            self.ventana,
            trabajo,
            title="Contando los valores",
            on_success=listo,
            on_error=fallo,
            on_cancel=lambda: self.boton_calcular.state(["!disabled"]),
        )

    def _mostrar(self, resultado: Any, normalizado: bool) -> None:
        self.resultado = resultado
        self._normalizado = normalizado
        self.calculados[tuple(resultado.columns)] = resultado
        tabla = resultado.table
        columnas = list(resultado.columns)
        self.tabla.delete(*self.tabla.get_children())
        ids = [f"c{i}" for i in range(len(columnas) + 2)]
        self.tabla.configure(columns=ids, displaycolumns="#all")
        titulos = [*columnas, resultado.count_column, resultado.percent_column]
        for i, (cid, titulo) in enumerate(zip(ids, titulos, strict=True)):
            numerica = i >= len(columnas)
            self.tabla.heading(cid, text=titulo)
            self.tabla.column(
                cid,
                width=110 if numerica else 180,
                anchor="e" if numerica else "w",
                stretch=not numerica and i == len(columnas) - 1,
            )
        visibles = tabla.iloc[:MAX_FILAS_TABLA]
        n = len(columnas)
        for pos, fila in enumerate(visibles.itertuples(index=False)):
            valores = [str(v) for v in fila[:n]]
            valores.append(formatear_numero(int(fila[n])))
            valores.append(formatear_porcentaje(float(fila[n + 1])))
            self.tabla.insert("", "end", iid=str(pos), values=valores)
        self.etiqueta_resumen.configure(
            text=(
                f"{formatear_numero(resultado.distinct_total)} valores distintos en "
                f"{formatear_numero(resultado.rows_total)} filas."
            )
        )
        avisos = []
        if resultado.truncated:
            avisos.append(
                f"La lista es muy larga: se calcularon solo los primeros "
                f"{formatear_numero(len(tabla))} valores."
            )
        if len(tabla) > MAX_FILAS_TABLA:
            avisos.append(
                f"En pantalla se ven los primeros {formatear_numero(MAX_FILAS_TABLA)}; "
                "para verlos todos, exportá la lista."
            )
        self.etiqueta_aviso.configure(text=" ".join(avisos))
        grupos = resultado.variant_groups
        if grupos:
            partes = [" / ".join(f"«{v}»" if v else "«(vacío)»" for v in g) for g in grupos]
            texto = "Estos valores parecen el mismo escrito distinto: " + "; ".join(
                partes[:MAX_GRUPOS_VARIANTES]
            )
            if len(partes) > MAX_GRUPOS_VARIANTES:
                texto += f" … y {formatear_numero(len(partes) - MAX_GRUPOS_VARIANTES)} grupos más."
            texto += ". Podés tildar «Juntar escrituras parecidas» y calcular de nuevo."
            self.etiqueta_variantes.configure(text=texto)
        else:
            self.etiqueta_variantes.configure(text="")
        self._actualizar_botones()

    def _actualizar_botones(self) -> None:
        hay = self.resultado is not None
        self.boton_exportar.state(["!disabled"] if hay else ["disabled"])
        elegido = bool(self.tabla.selection())
        self.boton_ver.state(["!disabled"] if hay and elegido else ["disabled"])

    # ---- ver filas ----
    def valores_elegidos(self) -> tuple[str, ...] | None:
        seleccion = self.tabla.selection()
        if self.resultado is None or not seleccion:
            return None
        fila = self.resultado.table.iloc[int(seleccion[0])]
        return tuple(str(fila[c]) for c in self.resultado.columns)

    def ver_filas(self) -> None:
        valores = self.valores_elegidos()
        if valores is None or self.on_ver_filas is None:
            return
        self.on_ver_filas(list(self.resultado.columns), valores, self._normalizado)

    # ---- exportar ----
    def exportar(self) -> None:
        if self.resultado is None:
            return
        opciones: dict[str, Any] = {
            "parent": self.ventana,
            "title": "Exportar esta lista",
            "defaultextension": ".xlsx",
            "filetypes": TIPOS_ARCHIVO,
        }
        if self.perfil is not None:
            opciones["initialdir"] = str(Path(self.perfil.file.path).parent)
            opciones["initialfile"] = f"Valores de {' y '.join(self.resultado.columns)}"[:80]
        elegido = filedialog.asksaveasfilename(**opciones)
        if not elegido:
            return
        ruta = Path(elegido)
        resultado = self.resultado
        formato = formato_de_ruta(ruta)

        def trabajo(progress, cancel):
            return exportar.guardar_distintos(resultado, ruta, formato)

        run_in_background(
            self.ventana,
            trabajo,
            title="Guardando la lista",
            on_success=lambda _r: messagebox.showinfo(
                "Listo", f"La lista quedó guardada en:\n{ruta}", parent=self.ventana
            ),
            on_error=lambda exc: messagebox.showerror(
                "No se pudo guardar", mensaje_de(exc), parent=self.ventana
            ),
            cancellable=False,
        )
