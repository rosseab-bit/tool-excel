"""Editor de resumen reutilizable (PLAN §3.2 pasos 4 y 7, docs/contratos.md § Fase 4).

La UI no calcula: arma `AggregationSpec` del contrato y llama a `core.aggregate.summarize` en el
worker. El core se importa de forma perezosa (`_core`) para poder reemplazarlo en tests.
"""

from __future__ import annotations

import importlib
import logging
import tkinter as tk
from collections.abc import Callable
from tkinter import messagebox, ttk
from types import SimpleNamespace
from typing import TYPE_CHECKING, Any

from excel_tool.core.errors import ExcelToolError
from excel_tool.ui import guardar
from excel_tool.ui.desplazable import CuerpoDesplazable
from excel_tool.ui.formato import NOMBRES_TIPO
from excel_tool.ui.grilla import Grilla, formatear_numero
from excel_tool.ui.worker import MENSAJE_INESPERADO, run_in_background

if TYPE_CHECKING:
    import pandas as pd

log = logging.getLogger(__name__)

TEXTO_NO_DISPONIBLE = "Esta opción todavía no está disponible en esta versión."
MAX_FILAS_VISTA = 200_000
TEXTO_MUESTRA = f"Vista previa con las primeras {formatear_numero(MAX_FILAS_VISTA)} filas."
TEXTO_VACIO = "Elegí por qué columnas agrupar o agregá al menos un cálculo."
TEXTO_MIXTA = "Algunas celdas no son números y no se van a sumar."
TEXTO_CLAVE = "(clave)"
MUCHAS_COLUMNAS = 12
ALTO_GRUPOS = 130
ALTO_MAXIMO_CUERPO = 520  # píxeles; con el pie entra en una pantalla de 768 de alto

PERIODOS = {"por día": None, "por mes": "month", "por año": "year"}
TIPOS_FECHA = {"fecha", "fecha_hora"}
TIPOS_NUMERICOS = {"entero", "decimal", "mixto"}
OPERACIONES = {
    "Sumar": "sum",
    "Contar filas": "count",
    "Contar distintos": "count_distinct",
    "Promedio": "mean",
    "Mínimo": "min",
    "Máximo": "max",
    "Primero": "first",
    "Último": "last",
    "Listar valores": "join_text",
}
ETIQUETA_DE = {func: texto for texto, func in OPERACIONES.items()}
NUMERICAS = {"sum", "mean"}
PERIODO_TEXTO = {periodo: texto for texto, periodo in PERIODOS.items()}


def _core() -> Any:
    """Tipos y funciones del core de la Fase 4 (import perezoso, reemplazable en tests)."""
    try:
        aggregate = importlib.import_module("excel_tool.core.aggregate")
    except ImportError as exc:
        log.warning("El core de la Fase 4 todavía no existe: %s", exc)
        raise ExcelToolError(TEXTO_NO_DISPONIBLE) from exc
    return SimpleNamespace(
        GroupBy=aggregate.GroupBy,
        Aggregation=aggregate.Aggregation,
        AggregationSpec=aggregate.AggregationSpec,
        summarize=aggregate.summarize,
        default_name=aggregate.default_name,
    )


def _mensaje_de(exc: BaseException) -> str:
    return exc.mensaje if isinstance(exc, ExcelToolError) else MENSAJE_INESPERADO


def avisos_de(resultado: Any) -> list[str]:
    """Frases para el usuario con lo que el core no pudo convertir y sus avisos."""
    frases = []
    for columna, cantidad in (getattr(resultado, "not_numeric", None) or {}).items():
        if cantidad == 1:
            frases.append(f"1 celda de {columna} no es un número y no se sumó.")
        else:
            frases.append(
                f"{formatear_numero(cantidad)} celdas de {columna} no son números y no se sumaron."
            )
    frases.extend(getattr(resultado, "warnings", None) or [])
    return frases


class _FilaCalculo:
    """Un cálculo: columna, operación y nombre del resultado."""

    def __init__(self, editor: EditorResumen) -> None:
        self.editor = editor
        self.marco = ttk.Frame(editor.contenedor_calculos)
        self.marco.pack(fill="x", pady=1)
        self.var_columna = tk.StringVar()
        self.var_operacion = tk.StringVar(value=ETIQUETA_DE["count"])
        self.var_nombre = tk.StringVar()
        self.editado = False
        self._poniendo = False
        self.combo_operacion = ttk.Combobox(
            self.marco,
            textvariable=self.var_operacion,
            values=list(OPERACIONES),
            state="readonly",
            width=16,
        )
        self.combo_columna = ttk.Combobox(
            self.marco,
            textvariable=self.var_columna,
            values=editor.nombres,
            state="readonly",
            width=22,
        )
        self.etiqueta_tipo = ttk.Label(self.marco, text="", width=14)
        self.entrada_nombre = ttk.Entry(self.marco, textvariable=self.var_nombre, width=24)
        self.boton_quitar = ttk.Button(
            self.marco, text="×", width=3, command=lambda: editor.quitar_calculo(self)
        )
        self.combo_operacion.pack(side="left", padx=(0, 4))
        self.combo_columna.pack(side="left", padx=(0, 4))
        self.etiqueta_tipo.pack(side="left", padx=(0, 4))
        self.entrada_nombre.pack(side="left", padx=(0, 4), fill="x", expand=True)
        self.boton_quitar.pack(side="left")
        self.etiqueta_aviso = ttk.Label(editor.contenedor_calculos, text="", foreground="#8a5a00")
        self.combo_operacion.bind("<<ComboboxSelected>>", lambda _e: self.ajustar())
        self.combo_columna.bind(
            "<<ComboboxSelected>>", lambda _e: self.ajustar(cambio_columna=True)
        )
        self.var_nombre.trace_add("write", self._al_escribir_nombre)
        self.ajustar()

    # ---- estado ----
    @property
    def funcion(self) -> str:
        return OPERACIONES.get(self.var_operacion.get(), "count")

    @property
    def tipo(self) -> str:
        return self.editor.column_types.get(self.var_columna.get(), "texto")

    def operaciones_validas(self) -> list[str]:
        if self.var_columna.get() and self.tipo not in TIPOS_NUMERICOS:
            return [t for t, f in OPERACIONES.items() if f not in NUMERICAS]
        return list(OPERACIONES)

    def ajustar(self, cambio_columna: bool = False) -> None:
        if self.funcion == "count":
            self.var_columna.set("")
            self.combo_columna.configure(state="disabled")
        else:
            self.combo_columna.configure(state="readonly")
        if (
            self.funcion in NUMERICAS
            and self.var_columna.get()
            and self.tipo not in TIPOS_NUMERICOS
        ):
            self.var_operacion.set(ETIQUETA_DE["count_distinct"])
        self.combo_operacion.configure(values=self.operaciones_validas())
        columna = self.var_columna.get()
        self.etiqueta_tipo.configure(text=NOMBRES_TIPO.get(self.tipo, "") if columna else "")
        mixta = bool(columna) and self.funcion in NUMERICAS and self.tipo == "mixto"
        self._mostrar_aviso(TEXTO_MIXTA if mixta else "")
        self.actualizar_nombre()
        self.editor._cambio()

    def _mostrar_aviso(self, texto: str) -> None:
        if texto:
            self.etiqueta_aviso.configure(text=texto)
            self.etiqueta_aviso.pack(after=self.marco, anchor="w", padx=(4, 0))
        else:
            self.etiqueta_aviso.pack_forget()

    # ---- nombre ----
    def nombre_automatico(self) -> str:
        columna = self.var_columna.get()
        if self.funcion != "count" and not columna:
            return ""
        try:
            core = _core()
            return core.default_name(core.Aggregation(columna, self.funcion))  # type: ignore[arg-type]
        except ExcelToolError:
            return ""

    def actualizar_nombre(self) -> None:
        if self.editado:
            return
        self._poniendo = True
        try:
            self.var_nombre.set(self.nombre_automatico())
        finally:
            self._poniendo = False

    def _al_escribir_nombre(self, *_args: Any) -> None:
        if self._poniendo:
            return
        self.editado = bool(self.var_nombre.get().strip())
        if not self.editado:
            self.actualizar_nombre()
        self.editor._cambio()

    def poner(self, columna: str, funcion: str, nombre: str | None) -> None:
        self.var_operacion.set(ETIQUETA_DE.get(funcion, ETIQUETA_DE["count"]))
        self.var_columna.set(columna)
        self.editado = False
        self.ajustar()
        if nombre:
            self._poniendo = True
            try:
                self.var_nombre.set(nombre)
            finally:
                self._poniendo = False
            self.editado = True

    def completa(self) -> bool:
        return self.funcion == "count" or bool(self.var_columna.get())


class EditorResumen(ttk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        df: pd.DataFrame,
        column_types: dict[str, str],
        *,
        fixed_group: list[str] | None = None,
        on_change: Callable[[Any], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self.df = df
        self.column_types = dict(column_types)
        self.nombres = [str(c) for c in df.columns]
        self.fixed_group = [c for c in (fixed_group or []) if c in self.nombres]
        self.on_change = on_change
        self.filas: list[_FilaCalculo] = []
        self._silencio = 0
        self._version = 0
        self.var_check: dict[str, tk.BooleanVar] = {}
        self.var_periodo: dict[str, tk.StringVar] = {}
        self._marcos: dict[str, ttk.Frame] = {}
        self.combos_periodo: dict[str, ttk.Combobox] = {}
        self.checks: dict[str, ttk.Checkbutton] = {}
        self.resultado: Any = None
        self._armar()
        self._cambio()

    # ---- armado ----
    def _armar(self) -> None:
        grupos = ttk.LabelFrame(self, text="Agrupar por", padding=6)
        grupos.pack(fill="x")
        self.var_buscar = tk.StringVar()
        if len(self.nombres) > MUCHAS_COLUMNAS:
            fila = ttk.Frame(grupos)
            fila.pack(fill="x", pady=(0, 4))
            ttk.Label(fila, text="Buscar columna:").pack(side="left")
            self.entrada_buscar = ttk.Entry(fila, textvariable=self.var_buscar, width=24)
            self.entrada_buscar.pack(side="left", padx=4)
            self.var_buscar.trace_add("write", lambda *_a: self._filtrar_columnas())
        cuerpo = ttk.Frame(grupos)
        cuerpo.pack(fill="x")
        lienzo = tk.Canvas(cuerpo, height=ALTO_GRUPOS, highlightthickness=0)
        barra = ttk.Scrollbar(cuerpo, orient="vertical", command=lienzo.yview)
        lienzo.configure(yscrollcommand=barra.set)
        lienzo.pack(side="left", fill="x", expand=True)
        barra.pack(side="right", fill="y")
        self.lista_grupos = ttk.Frame(lienzo)
        ventana = lienzo.create_window((0, 0), window=self.lista_grupos, anchor="nw")
        self.lista_grupos.bind(
            "<Configure>", lambda _e: lienzo.configure(scrollregion=lienzo.bbox("all"))
        )
        lienzo.bind("<Configure>", lambda e: lienzo.itemconfigure(ventana, width=e.width))

        orden = self.fixed_group + [n for n in self.nombres if n not in self.fixed_group]
        for nombre in orden:
            self._armar_columna(nombre)
        self.etiqueta_frase = ttk.Label(grupos, text="", wraplength=620)
        self.etiqueta_frase.pack(fill="x", pady=(6, 0))

        calculos = ttk.LabelFrame(self, text="Calcular", padding=6)
        calculos.pack(fill="x", pady=(8, 0))
        self.contenedor_calculos = ttk.Frame(calculos)
        self.contenedor_calculos.pack(fill="x")
        self.boton_agregar = ttk.Button(
            calculos, text="+ Agregar cálculo", command=lambda: self.agregar_calculo()
        )
        self.boton_agregar.pack(anchor="w", pady=(4, 0))

        previa = ttk.Frame(self)
        previa.pack(fill="x", pady=(8, 0))
        self.boton_previa = ttk.Button(previa, text="Ver cómo queda", command=self.ver_como_queda)
        self.boton_previa.pack(side="left")
        self.etiqueta_estado = ttk.Label(previa, text="", wraplength=520)
        self.etiqueta_estado.pack(side="left", padx=8)
        self.grilla = Grilla(self, None)
        self.grilla.pack(fill="both", expand=True, pady=(6, 0))

    def _armar_columna(self, nombre: str) -> None:
        marco = ttk.Frame(self.lista_grupos)
        marco.pack(fill="x", anchor="w")
        fijo = nombre in self.fixed_group
        var = tk.BooleanVar(value=fijo)
        texto = f"{nombre} {TEXTO_CLAVE}" if fijo else nombre
        check = ttk.Checkbutton(marco, text=texto, variable=var, command=self._cambio)
        check.pack(side="left")
        if fijo:
            check.state(["disabled"])
        self.var_check[nombre] = var
        self.checks[nombre] = check
        self._marcos[nombre] = marco
        if self.column_types.get(nombre) in TIPOS_FECHA:
            periodo = tk.StringVar(value="por día")
            combo = ttk.Combobox(
                marco, textvariable=periodo, values=list(PERIODOS), state="readonly", width=9
            )
            combo.pack(side="left", padx=8)
            combo.bind("<<ComboboxSelected>>", lambda _e: self._cambio())
            self.var_periodo[nombre] = periodo
            self.combos_periodo[nombre] = combo

    def _filtrar_columnas(self) -> None:
        texto = self.var_buscar.get().strip().lower()
        for nombre in self.fixed_group + [n for n in self.nombres if n not in self.fixed_group]:
            marco = self._marcos[nombre]
            marco.pack_forget()
            if not texto or texto in nombre.lower() or self.var_check[nombre].get():
                marco.pack(fill="x", anchor="w")

    # ---- cálculos ----
    def agregar_calculo(self) -> _FilaCalculo:
        fila = _FilaCalculo(self)
        self.filas.append(fila)
        self._cambio()
        return fila

    def quitar_calculo(self, fila: _FilaCalculo) -> None:
        if fila in self.filas:
            self.filas.remove(fila)
        fila.etiqueta_aviso.destroy()
        fila.marco.destroy()
        self._cambio()

    # ---- contrato ----
    def columnas_agrupadas(self) -> list[str]:
        orden = self.fixed_group + [n for n in self.nombres if n not in self.fixed_group]
        return [n for n in orden if self.var_check[n].get()]

    def spec(self) -> Any:
        core = _core()
        grupos = []
        for nombre in self.columnas_agrupadas():
            periodo = None
            if nombre in self.var_periodo:
                periodo = PERIODOS.get(self.var_periodo[nombre].get())
            grupos.append(core.GroupBy(nombre, periodo))
        calculos = [
            core.Aggregation(
                f.var_columna.get() if f.funcion != "count" else "",
                f.funcion,
                f.var_nombre.get().strip() if f.editado else None,
            )
            for f in self.filas
            if f.completa()
        ]
        return core.AggregationSpec(grupos, calculos)

    def set_spec(self, spec: Any) -> None:
        self._silencio += 1
        try:
            pedidas = {g.column: g.period for g in spec.group_by}
            for nombre, var in self.var_check.items():
                var.set(nombre in pedidas or nombre in self.fixed_group)
                if nombre in self.var_periodo:
                    self.var_periodo[nombre].set(PERIODO_TEXTO.get(pedidas.get(nombre), "por día"))
            for fila in list(self.filas):
                self.quitar_calculo(fila)
            for agg in spec.aggregations:
                fila = _FilaCalculo(self)
                self.filas.append(fila)
                fila.poner(agg.column, agg.func, agg.name)
        finally:
            self._silencio -= 1
        self._cambio()

    def _cambio(self) -> None:
        if self._silencio:
            return
        self._actualizar_frase()
        if self.on_change is None:
            return
        try:
            spec = self.spec()
        except ExcelToolError:
            return
        self.on_change(spec)

    def _actualizar_frase(self) -> None:
        partes = []
        for nombre in self.columnas_agrupadas():
            periodo = self.var_periodo[nombre].get() if nombre in self.var_periodo else "por día"
            partes.append(nombre if periodo == "por día" else f"{nombre} ({periodo})")
        if not partes:
            texto = "Una sola fila con el resultado de toda la tabla"
        elif len(partes) == 1:
            texto = f"Una fila por cada {partes[0]}"
        else:
            texto = f"Una fila por cada {', '.join(partes[:-1])} y {partes[-1]}"
        self.etiqueta_frase.configure(text=texto)

    # ---- vista previa ----
    def hay_algo_que_resumir(self) -> bool:
        return bool(self.columnas_agrupadas()) or any(f.completa() for f in self.filas)

    def poner_estado(self, texto: str) -> None:
        self.etiqueta_estado.configure(text=texto)

    def ver_como_queda(self) -> None:
        if not self.hay_algo_que_resumir():
            self.poner_estado(TEXTO_VACIO)
            return
        try:
            core = _core()
            spec = self.spec()
        except ExcelToolError as exc:
            self.poner_estado(exc.mensaje)
            return
        muestra = len(self.df) > MAX_FILAS_VISTA
        tabla = self.df.iloc[:MAX_FILAS_VISTA] if muestra else self.df
        tipos = dict(self.column_types)
        self._version += 1
        version = self._version
        self.boton_previa.state(["disabled"])
        self.poner_estado("Calculando…")

        def trabajo(progress, cancel):
            return core.summarize(tabla, spec, column_types=tipos, progress=progress, cancel=cancel)

        def listo(resultado):
            self.boton_previa.state(["!disabled"])
            if version == self._version and self.winfo_exists():
                self._mostrar(resultado, muestra)

        def fallo(exc):
            self.boton_previa.state(["!disabled"])
            if version == self._version and self.winfo_exists():
                self.poner_estado(_mensaje_de(exc))

        def cancelado():
            self.boton_previa.state(["!disabled"])
            if self.winfo_exists():
                self.poner_estado("")

        run_in_background(
            self,
            trabajo,
            title="Calculando la vista previa",
            on_success=listo,
            on_error=fallo,
            on_cancel=cancelado,
        )

    def _mostrar(self, resultado: Any, muestra: bool) -> None:
        self.resultado = resultado
        tabla = resultado.df
        frases = ([TEXTO_MUESTRA] if muestra else []) + avisos_de(resultado)
        filas = f"{formatear_numero(len(tabla))} filas en el resumen."
        self.poner_estado(" ".join([filas, *frases]))
        self.grilla.set_proveedor(
            lambda pagina: (
                tabla.iloc[pagina * self.grilla.page_size : (pagina + 1) * self.grilla.page_size],
                len(tabla),
            )
        )


def abrir_resumen(
    parent: tk.Misc,
    df: pd.DataFrame,
    column_types: dict[str, str],
    nombre_archivo: str,
) -> VentanaResumen | None:
    """Abre la ventana "Resumir". Devuelve la ventana (o None si el core no está)."""
    try:
        _core()
    except ExcelToolError as exc:
        messagebox.showinfo("Resumir", exc.mensaje, parent=parent)
        return None
    return VentanaResumen(parent, df, column_types, nombre_archivo)


class VentanaResumen:
    def __init__(
        self,
        parent: tk.Misc,
        df: pd.DataFrame,
        column_types: dict[str, str],
        nombre_archivo: str,
    ) -> None:
        self.df = df
        self.nombre_archivo = nombre_archivo or "resultado"
        self.ventana = tk.Toplevel(parent)
        self.ventana.title("Resumir")
        self.ventana.minsize(760, 560)
        self.cuerpo = CuerpoDesplazable(
            self.ventana,
            alto_maximo=ALTO_MAXIMO_CUERPO,
            padding_cuerpo=10,
            padding_pie=(10, 0, 10, 10),
        )
        self.editor = EditorResumen(self.cuerpo.marco, df, column_types)
        self.editor.pack(fill="both", expand=True)
        self.boton_cerrar = ttk.Button(self.cuerpo.pie, text="Cerrar", command=self.cerrar)
        self.boton_cerrar.pack(side="right")
        self.boton_guardar = ttk.Button(
            self.cuerpo.pie, text="Guardar el resumen…", command=self.guardar_resumen
        )
        self.boton_guardar.pack(side="right", padx=6)
        self.ventana.protocol("WM_DELETE_WINDOW", self.cerrar)

    def cerrar(self) -> None:
        self.cuerpo.desactivar_rueda()
        self.ventana.destroy()

    def guardar_resumen(self) -> None:
        editor = self.editor
        if not editor.hay_algo_que_resumir():
            editor.poner_estado(TEXTO_VACIO)
            return
        try:
            core = _core()
            spec = editor.spec()
        except ExcelToolError as exc:
            editor.poner_estado(exc.mensaje)
            return
        df, tipos = self.df, dict(editor.column_types)
        self.boton_guardar.state(["disabled"])

        def trabajo(progress, cancel):
            return core.summarize(df, spec, column_types=tipos, progress=progress, cancel=cancel)

        def listo(resultado):
            self.boton_guardar.state(["!disabled"])
            if not self.ventana.winfo_exists():
                return
            frases = avisos_de(resultado)
            if frases:
                editor.poner_estado(" ".join(frases))
            guardar.abrir_guardar(
                self.ventana,
                resultado.df,
                nombre_sugerido=f"{self.nombre_archivo}_resumen",
                column_types=dict(resultado.column_types),
            )

        def fallo(exc):
            self.boton_guardar.state(["!disabled"])
            if self.ventana.winfo_exists():
                messagebox.showerror("Resumir", _mensaje_de(exc), parent=self.ventana)

        run_in_background(
            self.ventana,
            trabajo,
            title="Calculando el resumen",
            on_success=listo,
            on_error=fallo,
            on_cancel=lambda: self.boton_guardar.state(["!disabled"]),
        )
