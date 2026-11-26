"""Paso "Guardar" reutilizable (PLAN §3.3, docs/contratos.md § Fase 3).

La UI no calcula nada: arma `SplitSpec` y `OutputSpec` del contrato y llama al core. El semáforo,
las partes, los tamaños, las sugerencias y los mensajes vienen de `core/limits.py`, `core/split.py`
y `core/output.py`. El core se importa de forma perezosa (`_core`) para poder reemplazarlo en tests.
"""

from __future__ import annotations

import importlib
import logging
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass, replace
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from types import SimpleNamespace
from typing import TYPE_CHECKING, Any

from excel_tool import platform_utils
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.preferences import load_preferences, save_preferences
from excel_tool.ui.desplazable import CuerpoDesplazable
from excel_tool.ui.grilla import formatear_numero
from excel_tool.ui.worker import MENSAJE_INESPERADO, run_in_background

if TYPE_CHECKING:
    import pandas as pd

log = logging.getLogger(__name__)

TEXTO_NO_DISPONIBLE = "Esta opción todavía no está disponible en esta versión."
AUTO_FILAS = 2_000_000  # por debajo, las columnas para dividir se buscan solas al abrir
MAX_PARTES_VISTA = 200
DEBOUNCE_MS = 300
ALTO_MAXIMO_CUERPO = 470  # píxeles; con el pie entra en una pantalla de 768 de alto

DESTINOS = (
    ("excel", "Microsoft Excel"),
    ("google_sheets", "Google Sheets"),
    ("solo_guardar", "Solo guardarlo (para otro sistema)"),
)
DESTINO_INICIAL = "excel"
FORMATOS = {
    "excel": (("xlsx", "Excel (.xlsx)"), ("csv_excel", "CSV para Excel")),
    "google_sheets": (("xlsx", "Excel (.xlsx)"), ("csv_google", "CSV para Google Sheets")),
    "solo_guardar": (("csv_standard", "CSV (texto con columnas)"),),
}
FORMATO_INICIAL = {"excel": "xlsx", "google_sheets": "csv_google", "solo_guardar": "csv_standard"}
NOMBRE_DESTINO = {
    "excel": "Excel",
    "google_sheets": "Google Sheets",
    "solo_guardar": "otro sistema",
}

NIVELES = {
    "ok": ("TODO BIEN", "#2e7d32", "#ffffff"),
    "recomendado": ("CONVIENE DIVIDIR", "#f9a825", "#000000"),
    "obligatorio": ("NO ENTRA EN UN SOLO ARCHIVO", "#c62828", "#ffffff"),
}
TEXTO_OK = "El resultado entra sin problemas."
TEXTO_OBLIGATORIO = (
    "No se puede guardar en un solo archivo. Elegí cómo dividirlo, o guardalo como CSV en "
    "«Solo guardarlo»."
)

MODOS_SPLIT = {"rows": "rows", "keep": "rows_keep_groups", "value": "by_value", "sheets": "sheets"}


def _core() -> Any:
    """Funciones y dataclasses del core de la Fase 3 (import perezoso, reemplazable en tests)."""
    try:
        limits = importlib.import_module("excel_tool.core.limits")
        split = importlib.import_module("excel_tool.core.split")
        output = importlib.import_module("excel_tool.core.output")
    except ImportError as exc:
        log.warning("El core de la Fase 3 todavía no existe: %s", exc)
        raise ExcelToolError(TEXTO_NO_DISPONIBLE) from exc
    return SimpleNamespace(
        estimate=limits.estimate,
        check=limits.check,
        suggest_split_columns=limits.suggest_split_columns,
        plan_split=split.plan_split,
        SplitSpec=split.SplitSpec,
        OutputSpec=output.OutputSpec,
        save_result=output.save_result,
    )


@dataclass
class _Division:
    modo: str  # "none" | "rows" | "keep" | "value" | "sheets"
    spec: Any | None = None
    max_filas: int | None = None
    falta: str | None = None  # qué le falta elegir al usuario


def _mensaje_de(exc: BaseException) -> str:
    return exc.mensaje if isinstance(exc, ExcelToolError) else MENSAJE_INESPERADO


def _texto_sugerencia(sug: Any) -> str:
    marca = "✅" if sug.fits else "⚠️"
    return (
        f"{sug.column}: {formatear_numero(sug.parts)} archivos, el más grande de "
        f"{formatear_numero(sug.largest_part_rows)} filas {marca}"
    )


def abrir_guardar(
    parent: tk.Misc,
    df: pd.DataFrame,
    *,
    nombre_sugerido: str,
    on_terminado: Callable[[Any], None] | None = None,
    claves_sugeridas: list[str] | None = None,
    column_types: dict | None = None,
    salida_inicial: Any | None = None,
) -> VentanaGuardar | None:
    """Abre la ventana del paso Guardar. Devuelve la ventana (o None si el core no está).

    `salida_inicial` (un `OutputSpec`, por ejemplo el de una receta) deja preelegidos el destino,
    el formato, la división y la carpeta. Después de guardar, `VentanaGuardar.salida_usada` tiene
    el `OutputSpec` con el que se guardó (para ponerlo en una receta).
    """
    try:
        core = _core()
    except ExcelToolError as exc:
        messagebox.showinfo("Guardar", exc.mensaje, parent=parent)
        return None
    return VentanaGuardar(
        parent,
        df,
        core,
        nombre_sugerido=nombre_sugerido,
        on_terminado=on_terminado,
        claves_sugeridas=claves_sugeridas,
        column_types=column_types,
        salida_inicial=salida_inicial,
    )


class VentanaGuardar:
    def __init__(
        self,
        parent: tk.Misc,
        df: pd.DataFrame,
        core: Any,
        *,
        nombre_sugerido: str,
        on_terminado: Callable[[Any], None] | None,
        claves_sugeridas: list[str] | None,
        column_types: dict | None = None,
        salida_inicial: Any | None = None,
    ) -> None:
        self.df = df
        self.salida_inicial = salida_inicial
        self.salida_usada: Any | None = None
        self._spec_en_curso: Any | None = None
        self._valor_pedido: str | None = None
        self.column_types = column_types
        self.core = core
        self.on_terminado = on_terminado
        nombre = Path(nombre_sugerido)
        self.nombre = nombre.stem if nombre.suffix.lower() in (".xlsx", ".csv") else nombre.name
        self.columnas = [str(c) for c in df.columns]
        self.claves = [c for c in (claves_sugeridas or []) if c in self.columnas]

        self._estim: Any | None = None
        self._check: Any | None = None
        self._cache_sug: dict[tuple[str, str], list[Any]] = {}
        self._sug: list[Any] = []
        self._partes: list[Any] | None = None
        self._error_plan: str | None = None
        self._gen_a = 0
        self._gen_b = 0
        self._modo_manual = False
        self._n_editado = False
        self._ocupado = False
        self._after_n: str | None = None
        self._resultado: Any | None = None
        self._sug_por_texto: dict[str, str] = {}

        prefs = load_preferences()
        destino = prefs.last_destination
        if salida_inicial is not None and salida_inicial.destination in dict(DESTINOS):
            destino = salida_inicial.destination
        if destino not in dict(DESTINOS):
            destino = DESTINO_INICIAL

        self.win = tk.Toplevel(parent)
        self.win.title("Guardar el resultado")
        try:
            self.win.transient(parent.winfo_toplevel())
        except tk.TclError:
            pass
        self._fuente_negrita = tkfont.nametofont("TkDefaultFont").copy()
        self._fuente_negrita.configure(weight="bold")

        self.v_destino = tk.StringVar(master=self.win, value=destino)
        formato = FORMATO_INICIAL[destino]
        if salida_inicial is not None and salida_inicial.format in dict(FORMATOS[destino]):
            formato = salida_inicial.format
        self.v_formato = tk.StringVar(master=self.win, value=formato)
        self.v_modo = tk.StringVar(master=self.win, value="none")
        self.v_filas = tk.StringVar(master=self.win, value="")
        self.v_col_grupo = tk.StringVar(
            master=self.win, value=self.claves[0] if self.claves else ""
        )
        self.v_col_valor = tk.StringVar(master=self.win, value="")
        self.v_subdividir = tk.BooleanVar(master=self.win, value=True)

        self._armar()
        self.win.protocol("WM_DELETE_WINDOW", self.cerrar)
        self._ajustar_opciones()
        self._preelegir_division()
        self._etapa_a(buscar=len(df) < AUTO_FILAS)

    # ------------------------------------------------------------------ armado
    def _armar(self) -> None:
        # Abajo, siempre a la vista: resultado y botones. Arriba, lo demás con desplazamiento.
        self.cuerpo_desplazable = CuerpoDesplazable(self.win, alto_maximo=ALTO_MAXIMO_CUERPO)
        self.pie = self.cuerpo_desplazable.pie
        self.lienzo = self.cuerpo_desplazable.lienzo
        self.barra_v = self.cuerpo_desplazable.barra_v
        self.marco = marco = self.cuerpo_desplazable.marco

        # 1. Destino
        f1 = ttk.LabelFrame(marco, text="1. ¿Dónde vas a abrir el archivo?", padding=8)
        f1.pack(fill="x")
        self.rb_destino: dict[str, ttk.Radiobutton] = {}
        for valor, texto in DESTINOS:
            rb = ttk.Radiobutton(
                f1,
                text=texto,
                value=valor,
                variable=self.v_destino,
                command=self._al_cambiar_destino,
            )
            rb.pack(anchor="w", pady=2)
            self.rb_destino[valor] = rb

        # 2. Semáforo
        f2 = ttk.LabelFrame(marco, text="2. Tamaño del resultado", padding=8)
        f2.pack(fill="x", pady=(8, 0))
        self.lbl_nivel = tk.Label(f2, text="Calculando el tamaño…", anchor="w", padx=8, pady=2)
        self.lbl_nivel.pack(fill="x")
        self.lbl_principal = ttk.Label(f2, text="", font=self._fuente_negrita, wraplength=620)
        self.lbl_principal.pack(anchor="w", pady=(6, 0))
        self.lbl_detalle = ttk.Label(f2, text="", wraplength=620, justify="left")
        self.lbl_detalle.pack(anchor="w")

        # 3. Dividir
        f3 = ttk.LabelFrame(marco, text="3. ¿Cómo lo querés guardar?", padding=8)
        f3.pack(fill="x", pady=(8, 0))
        self.f_dividir = f3
        self.rb_uno = self._radio(f3, "none", "Todo en un solo archivo")
        self.rb_uno.grid(row=0, column=0, columnspan=4, sticky="w", pady=2)

        self.rb_filas = self._radio(f3, "rows", "En partes de")
        self.rb_filas.grid(row=1, column=0, sticky="w", pady=2)
        self.ent_filas = ttk.Entry(f3, textvariable=self.v_filas, width=12)
        self.ent_filas.grid(row=1, column=1, sticky="w")
        self.ent_filas.bind("<KeyRelease>", self._al_editar_filas)
        ttk.Label(f3, text="filas").grid(row=1, column=2, sticky="w", padx=4)

        self.rb_grupo = self._radio(f3, "keep", "Sin cortar un mismo valor de la columna")
        self.rb_grupo.grid(row=2, column=0, sticky="w", pady=2)
        self.cmb_grupo = ttk.Combobox(
            f3, textvariable=self.v_col_grupo, values=self.columnas, state="readonly", width=28
        )
        self.cmb_grupo.grid(row=2, column=1, columnspan=3, sticky="w")
        self.cmb_grupo.bind("<<ComboboxSelected>>", lambda _e: self._etapa_b())

        self.rb_valor = self._radio(f3, "value", "Un archivo por cada valor de la columna")
        self.rb_valor.grid(row=3, column=0, sticky="w", pady=2)
        self.cmb_valor = ttk.Combobox(
            f3, textvariable=self.v_col_valor, values=[], state="readonly", width=52
        )
        self.cmb_valor.grid(row=3, column=1, columnspan=3, sticky="w")
        self.cmb_valor.bind("<<ComboboxSelected>>", lambda _e: self._etapa_b())
        self.btn_buscar = ttk.Button(
            f3, text="Buscar columnas para dividir", command=lambda: self._etapa_a(buscar=True)
        )
        self.btn_buscar.grid(row=4, column=1, columnspan=3, sticky="w", pady=2)
        self.chk_subdividir = ttk.Checkbutton(
            f3,
            text="Si alguna parte queda muy grande, partirla también en partes de esa cantidad "
            "de filas",
            variable=self.v_subdividir,
            command=self._etapa_b,
        )
        self.chk_subdividir.grid(row=5, column=1, columnspan=3, sticky="w")

        self.rb_hojas = self._radio(f3, "sheets", "En varias hojas del mismo archivo")
        self.rb_hojas.grid(row=6, column=0, columnspan=4, sticky="w", pady=2)

        self.lbl_vista = ttk.Label(f3, text="", wraplength=620, justify="left")
        self.lbl_vista.grid(row=7, column=0, columnspan=4, sticky="w", pady=(6, 2))
        cuerpo = ttk.Frame(f3)
        cuerpo.grid(row=8, column=0, columnspan=4, sticky="ew")
        self.tabla = ttk.Treeview(
            cuerpo, columns=("nombre", "filas", "desde", "hasta"), show="headings", height=5
        )
        for col, texto, ancho in (
            ("nombre", "Nombre", 300),
            ("filas", "Filas", 90),
            ("desde", "Desde la fila", 100),
            ("hasta", "Hasta la fila", 100),
        ):
            self.tabla.heading(col, text=texto)
            self.tabla.column(col, width=ancho, anchor="w" if col == "nombre" else "e")
        barra = ttk.Scrollbar(cuerpo, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=barra.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")

        # 4. Formato
        f4 = ttk.LabelFrame(marco, text="4. Formato del archivo", padding=8)
        f4.pack(fill="x", pady=(8, 0))
        self.f_formato = f4
        self.rb_formato: dict[str, ttk.Radiobutton] = {}

        # 5. Guardar
        f5 = ttk.Frame(self.pie)
        f5.pack(side="bottom", fill="x", pady=(8, 0))
        self.f5 = f5
        self.lbl_estado = ttk.Label(f5, text="", wraplength=420, justify="left")
        self.lbl_estado.pack(side="left", fill="x", expand=True)
        self.btn_cerrar = ttk.Button(f5, text="Cerrar", command=self.cerrar)
        self.btn_cerrar.pack(side="right", padx=(6, 0))
        self.btn_guardar = ttk.Button(f5, text="Guardar…", command=self._guardar)
        self.btn_guardar.pack(side="right")

        # Resultado
        self.f_resultado = ttk.LabelFrame(self.pie, text="Listo", padding=8)
        self.lbl_resultado = ttk.Label(self.f_resultado, text="", wraplength=620, justify="left")
        self.lbl_resultado.pack(anchor="w")
        self.lbl_avisos = ttk.Label(self.f_resultado, text="", wraplength=620, justify="left")
        self.lbl_avisos.pack(anchor="w", pady=(4, 0))
        self.btn_carpeta = ttk.Button(
            self.f_resultado, text="Abrir carpeta", command=self._abrir_carpeta
        )
        self.btn_carpeta.pack(anchor="e", pady=(6, 0))

    def _radio(self, padre: tk.Misc, valor: str, texto: str) -> ttk.Radiobutton:
        return ttk.Radiobutton(
            padre,
            text=texto,
            value=valor,
            variable=self.v_modo,
            command=self._al_elegir_modo,
        )

    def _preelegir_division(self) -> None:
        """Deja elegida la división de `salida_inicial` (si la hay y se puede)."""
        dividir = getattr(self.salida_inicial, "split", None)
        if dividir is None:
            return
        modo = {v: k for k, v in MODOS_SPLIT.items()}.get(dividir.mode)
        if modo is None or (modo == "sheets" and not self.hojas_disponibles()):
            return
        self.v_modo.set(modo)
        self._modo_manual = True
        if dividir.rows_per_part:
            self.v_filas.set(formatear_numero(int(dividir.rows_per_part)))
            self._n_editado = True
        grupo = getattr(dividir, "group_column", None)
        if grupo in self.columnas:
            self.v_col_grupo.set(grupo)
        valores = list(getattr(dividir, "value_columns", None) or [])
        if valores:
            self._valor_pedido = valores[0]
            self.v_subdividir.set(bool(dividir.rows_per_part))

    # --------------------------------------------------------------- opciones
    def hojas_disponibles(self) -> bool:
        """Solo Excel, y solo con formato xlsx (una hoja por parte)."""
        return self.v_destino.get() == "excel" and self.v_formato.get() == "xlsx"

    def _ajustar_opciones(self) -> None:
        """Formatos según el destino y opción de varias hojas."""
        destino = self.v_destino.get()
        permitidos = [valor for valor, _t in FORMATOS[destino]]
        if self.v_formato.get() not in permitidos:
            self.v_formato.set(FORMATO_INICIAL[destino])
        for rb in self.rb_formato.values():
            rb.destroy()
        self.rb_formato = {}
        for valor, texto in FORMATOS[destino]:
            rb = ttk.Radiobutton(
                self.f_formato,
                text=texto,
                value=valor,
                variable=self.v_formato,
                command=self._al_cambiar_formato,
            )
            rb.pack(anchor="w", pady=2)
            self.rb_formato[valor] = rb

        if destino == "excel":
            self.rb_hojas.grid()
        else:
            self.rb_hojas.grid_remove()
        habilitar = self.hojas_disponibles()
        self.rb_hojas.state(["!disabled"] if habilitar else ["disabled"])
        if not habilitar and self.v_modo.get() == "sheets":
            self.v_modo.set("rows")

    def _al_cambiar_destino(self) -> None:
        prefs = load_preferences()
        save_preferences(replace(prefs, last_destination=self.v_destino.get()))
        self.v_formato.set(FORMATO_INICIAL[self.v_destino.get()])
        # La división vuelve a la sugerida para el nuevo destino. Lo preelegido de una receta
        # vale solo para el destino de esa receta.
        self._modo_manual = False
        self._n_editado = False
        self.v_filas.set("")
        self.v_modo.set("none")
        self._ajustar_opciones()
        if getattr(self.salida_inicial, "destination", None) == self.v_destino.get():
            self._preelegir_division()
        self._etapa_a(buscar=len(self.df) < AUTO_FILAS)

    def _al_cambiar_formato(self) -> None:
        self._ajustar_opciones()
        self._etapa_a(buscar=len(self.df) < AUTO_FILAS)

    def _al_elegir_modo(self) -> None:
        self._modo_manual = True
        if self.v_modo.get() == "keep":
            self._sugerir_filas()
        self._etapa_b()

    def _al_editar_filas(self, _evento: object = None) -> None:
        self._n_editado = True
        if self._after_n is not None:
            try:
                self.win.after_cancel(self._after_n)
            except tk.TclError:
                pass
        try:
            self._after_n = self.win.after(DEBOUNCE_MS, self._etapa_b)
        except tk.TclError:
            self._after_n = None

    # ------------------------------------------------------ etapa A: tamaño
    def _etapa_a(self, *, buscar: bool) -> None:
        """Estima el tamaño y calcula el semáforo (y las columnas para dividir) en el worker."""
        destino, formato = self.v_destino.get(), self.v_formato.get()
        clave = (destino, formato)
        self._gen_a += 1
        gen = self._gen_a
        self._check = None
        self.lbl_nivel.configure(text="Calculando el tamaño…", bg=self._color_neutro(), fg="black")
        self.lbl_principal.configure(text="")
        self.lbl_detalle.configure(text="")
        self._actualizar_estado()
        sugerencias = self._cache_sug.get(clave)
        buscar_ahora = buscar and sugerencias is None
        if buscar and sugerencias is not None:
            buscar_ahora = False
        previa = self._estim
        df, core = self.df, self.core

        def trabajo(progress: Any, cancel: Any) -> tuple[Any, list[Any] | None, Any]:
            est = previa if previa is not None else core.estimate(df)
            sug = sugerencias
            if buscar_ahora:
                sug = core.suggest_split_columns(
                    df, destino, formato, progress=progress, cancel=cancel
                )
            chk = core.check(est, destino, formato, suggestions=sug)
            return est, sug, chk

        def listo(res: object) -> None:
            if gen != self._gen_a:
                return
            est, sug, chk = res  # type: ignore[misc]
            self._estim = est
            if sug is not None:
                self._cache_sug[clave] = sug
            self._sug = list(sug) if sug is not None else []
            self._check = chk
            self._aplicar_check()
            self._etapa_b()

        def fallo(exc: BaseException) -> None:
            if gen != self._gen_a:
                return
            self.lbl_nivel.configure(text="No se pudo calcular el tamaño")
            self.lbl_principal.configure(text=_mensaje_de(exc))
            self._actualizar_estado()

        def cancelado() -> None:
            if gen == self._gen_a and buscar_ahora:
                self._etapa_a(buscar=False)

        run_in_background(
            self.win,
            trabajo,
            title="Buscando columnas para dividir…" if buscar_ahora else "Calculando el tamaño…",
            on_success=listo,
            on_error=fallo,
            cancellable=buscar_ahora,
            modal=False,
            on_cancel=cancelado,
        )

    def _color_neutro(self) -> str:
        return str(ttk.Style(self.win).lookup("TFrame", "background") or "#d9d9d9")

    def _aplicar_check(self) -> None:
        chk = self._check
        if chk is None:
            return
        etiqueta, fondo, frente = NIVELES[chk.level]
        self.lbl_nivel.configure(text=f"●  {etiqueta}", bg=fondo, fg=frente)
        razones = list(chk.reasons)
        if razones:
            principal, detalle = razones[0], "\n".join(razones[1:])
        else:
            principal, detalle = TEXTO_OK if chk.level == "ok" else "", ""
        if chk.level == "obligatorio":
            detalle = f"{detalle}\n{TEXTO_OBLIGATORIO}".strip()
        self.lbl_principal.configure(text=principal)
        self.lbl_detalle.configure(text=detalle)

        obligatorio = chk.level == "obligatorio"
        self.rb_uno.configure(
            text="Todo en un solo archivo" + (" (no entra)" if obligatorio else "")
        )
        self.rb_uno.state(["disabled"] if obligatorio else ["!disabled"])

        self._cargar_sugerencias()
        preferido = "keep" if self.v_col_grupo.get() else "rows"
        if obligatorio and self.v_modo.get() == "none":
            self.v_modo.set(preferido)
        elif not self._modo_manual:
            self.v_modo.set("none" if chk.level == "ok" else preferido)
        self._sugerir_filas()

    def _sugerir_filas(self) -> None:
        """Filas por parte sugeridas (si el usuario no las tocó). "Sin cortar un mismo grupo"
        usa el tope de filas del destino (`max_rows_per_part`): con la cantidad sugerida no
        queda margen y siempre sale una parte de más."""
        chk = self._check
        if chk is None or self._n_editado:
            return
        filas = chk.suggested_rows_per_part
        if self.v_modo.get() == "keep":
            filas = getattr(chk, "max_rows_per_part", None) or filas
        if filas:
            self.v_filas.set(formatear_numero(int(filas)))

    def _cargar_sugerencias(self) -> None:
        self._sug_por_texto = {_texto_sugerencia(s): s.column for s in self._sug}
        textos = list(self._sug_por_texto)
        self.cmb_valor.configure(values=textos)
        if self.v_col_valor.get() not in self._sug_por_texto:
            pedido = next(
                (t for t, c in self._sug_por_texto.items() if c == self._valor_pedido), None
            )
            self.v_col_valor.set(pedido or (textos[0] if textos else ""))

    # -------------------------------------------------------- etapa B: partes
    def _filas_por_parte(self) -> int | None:
        texto = self.v_filas.get().replace(".", "").replace(" ", "").strip()
        if not texto.isdigit() or int(texto) <= 0:
            return None
        return int(texto)

    def _leer_division(self) -> _Division:
        modo = self.v_modo.get()
        if modo == "none":
            return _Division("none")
        filas = self._filas_por_parte()
        falta_filas = "Escribí cuántas filas querés en cada parte (un número mayor que cero)."
        spec_cls = self.core.SplitSpec
        if modo == "rows":
            if filas is None:
                return _Division(modo, falta=falta_filas)
            return _Division(modo, spec_cls(mode="rows", rows_per_part=filas))
        if modo == "sheets":
            if filas is None:
                return _Division(modo, falta=falta_filas)
            return _Division(modo, spec_cls(mode="sheets", rows_per_part=filas))
        if modo == "keep":
            columna = self.v_col_grupo.get()
            if not columna:
                return _Division(modo, falta="Elegí la columna que no se tiene que cortar.")
            if filas is None:
                return _Division(modo, falta=falta_filas)
            return _Division(
                modo, spec_cls(mode="rows_keep_groups", rows_per_part=filas, group_column=columna)
            )
        # "value"
        columna = self._sug_por_texto.get(self.v_col_valor.get(), "")
        if not columna:
            return _Division(
                modo,
                falta="Elegí una columna. Si no hay ninguna, tocá «Buscar columnas para dividir».",
            )
        subdividir = self.v_subdividir.get()
        if subdividir and filas is None:
            return _Division(modo, falta=falta_filas)
        tope = filas if subdividir else None
        return _Division(
            modo,
            spec_cls(mode="by_value", value_columns=[columna], rows_per_part=tope),
            max_filas=tope,
        )

    def _etapa_b(self) -> None:
        """Vista previa de las partes (con `plan_split`, en el worker)."""
        self._after_n = None
        self._gen_b += 1
        gen = self._gen_b
        div = self._leer_division()
        self._partes = None
        self._error_plan = None
        self._vaciar_tabla()
        self._mostrar_opciones_valor(div.modo)
        if div.spec is None:
            self.lbl_vista.configure(
                text=div.falta
                or "Se guarda todo en un solo archivo: no hace falta calcular partes."
            )
            self._actualizar_estado()
            return
        self.lbl_vista.configure(text="Calculando las partes…")
        df, core, nombre = self.df, self.core, self.nombre

        def trabajo(_progress: Any, _cancel: Any) -> list[Any]:
            return list(
                core.plan_split(df, div.spec, base_name=nombre, max_rows_per_part=div.max_filas)
            )

        def listo(partes: object) -> None:
            if gen != self._gen_b:
                return
            self._partes = list(partes)  # type: ignore[call-overload]
            self._mostrar_partes(div.modo)
            self._actualizar_estado()

        def fallo(exc: BaseException) -> None:
            if gen != self._gen_b:
                return
            self._error_plan = _mensaje_de(exc)
            self.lbl_vista.configure(text=self._error_plan)
            self._actualizar_estado()

        run_in_background(
            self.win,
            trabajo,
            title="Calculando las partes…",
            on_success=listo,
            on_error=fallo,
            cancellable=False,
            modal=False,
        )
        self._actualizar_estado()

    def _mostrar_opciones_valor(self, modo: str) -> None:
        if modo == "value":
            self.chk_subdividir.state(["!disabled"])
        else:
            self.chk_subdividir.state(["disabled"])

    def _vaciar_tabla(self) -> None:
        self.tabla.delete(*self.tabla.get_children())

    def _mostrar_partes(self, modo: str) -> None:
        partes = self._partes or []
        self._vaciar_tabla()
        for parte in partes[:MAX_PARTES_VISTA]:
            self.tabla.insert(
                "",
                "end",
                values=(
                    parte.name,
                    formatear_numero(parte.rows),
                    formatear_numero(parte.first_row),
                    formatear_numero(parte.last_row),
                ),
            )
        total = len(partes)
        unidad = "hojas" if modo == "sheets" else "archivos"
        if total == 1:
            unidad = unidad[:-1] if modo != "sheets" else "hoja"
        texto = f"Se van a generar {formatear_numero(total)} {unidad}."
        if total > MAX_PARTES_VISTA:
            texto += f" Se muestran las primeras {MAX_PARTES_VISTA}."
        self.lbl_vista.configure(text=texto)

    # ----------------------------------------------------------------- estado
    def puede_guardar(self) -> bool:
        chk = self._check
        if self._ocupado or chk is None or self._error_plan:
            return False
        div = self._leer_division()
        if div.modo == "none":
            return chk.level != "obligatorio"
        return div.spec is not None and self._partes is not None

    def _actualizar_estado(self) -> None:
        self.btn_guardar.state(["!disabled"] if self.puede_guardar() else ["disabled"])

    def _bloquear(self, bloqueado: bool) -> None:
        self._ocupado = bloqueado
        self._actualizar_estado()

    # ---------------------------------------------------------------- guardar
    def _guardar(self) -> None:
        if not self.puede_guardar():
            return
        destino, formato = self.v_destino.get(), self.v_formato.get()
        ext = ".xlsx" if formato == "xlsx" else ".csv"
        div = self._leer_division()
        opciones: dict[str, Any] = {
            "parent": self.win,
            "title": (
                "Elegí el nombre (se crea una carpeta con las partes)"
                if div.spec is not None and div.modo != "sheets"
                else "Guardar el resultado"
            ),
            "initialfile": f"{self.nombre}{ext}",
            "defaultextension": ext,
            "filetypes": [
                ("Excel" if ext == ".xlsx" else "CSV (texto con columnas)", f"*{ext}"),
            ],
        }
        carpeta = load_preferences().last_folder
        previa = getattr(self.salida_inicial, "path", None)
        if previa is not None and Path(previa).parent.is_dir():
            carpeta = Path(previa).parent
        if carpeta is not None and carpeta.is_dir():
            opciones["initialdir"] = str(carpeta)
        elegido = filedialog.asksaveasfilename(**opciones)
        if not elegido:
            return
        ruta = Path(elegido)
        if ruta.suffix.lower() != ext:
            ruta = ruta.with_name(ruta.name + ext)

        spec = self.core.OutputSpec(
            path=ruta,
            format=formato,
            destination=destino,
            split=div.spec,
            column_types=self.column_types,
        )
        self._spec_en_curso = spec
        prefs = load_preferences()
        save_preferences(replace(prefs, last_folder=ruta.parent, last_destination=destino))

        df, core = self.df, self.core
        self.lbl_estado.configure(text="Guardando…")
        self.f_resultado.pack_forget()
        self._bloquear(True)

        def trabajo(progress: Any, cancel: Any) -> Any:
            return core.save_result(df, spec, progress=progress, cancel=cancel)

        run_in_background(
            self.win,
            trabajo,
            title="Guardando el resultado…",
            on_success=self._al_guardar,
            on_error=self._al_fallar,
            on_cancel=self._al_cancelar,
        )

    def _al_guardar(self, resultado: object) -> None:
        self._bloquear(False)
        self._resultado = resultado
        self.salida_usada = self._spec_en_curso
        archivos = len({p.path for p in resultado.parts}) or len(resultado.files)  # type: ignore[attr-defined]
        filas = formatear_numero(resultado.rows_total)  # type: ignore[attr-defined]
        if archivos == 1:
            texto = f"Se guardó 1 archivo con {filas} filas en total"
            hojas = len(resultado.parts)  # type: ignore[attr-defined]
            if hojas > 1:
                texto += f", repartidas en {formatear_numero(hojas)} hojas"
        else:
            texto = f"Se guardaron {formatear_numero(archivos)} archivos con {filas} filas en total"
        self.lbl_resultado.configure(text=f"{texto} (coincide con el resultado ✅)")
        avisos = list(resultado.warnings)  # type: ignore[attr-defined]
        for columna, cantidad in resultado.unconverted.items():  # type: ignore[attr-defined]
            avisos.append(
                f"En la columna «{columna}», {formatear_numero(cantidad)} celdas no se pudieron "
                "convertir y quedaron como texto."
            )
        self.lbl_avisos.configure(text="\n".join(f"• {a}" for a in avisos))
        self.lbl_estado.configure(text="")
        self.f_resultado.pack(fill="x", pady=(8, 0), before=self.f5)
        if self.on_terminado is not None:
            try:
                self.on_terminado(resultado)
            except Exception:
                log.exception("Falló on_terminado del paso Guardar")

    def _al_fallar(self, exc: BaseException) -> None:
        self._bloquear(False)
        self.lbl_estado.configure(text="No se pudo guardar.")
        try:
            messagebox.showerror("Guardar el resultado", _mensaje_de(exc), parent=self.win)
        except tk.TclError:
            log.exception("No se pudo mostrar el mensaje de error")

    def _al_cancelar(self) -> None:
        self._bloquear(False)
        self.lbl_estado.configure(text="Se canceló. No se guardó nada.")

    def _abrir_carpeta(self) -> None:
        res = self._resultado
        if res is None:
            return
        carpeta = res.folder if res.folder is not None else res.files[0].parent
        try:
            platform_utils.open_folder(Path(carpeta))
        except Exception:
            log.exception("No se pudo abrir la carpeta %s", carpeta)
            messagebox.showwarning(
                "Abrir carpeta", f"No se pudo abrir la carpeta.\n{carpeta}", parent=self.win
            )

    def cerrar(self) -> None:
        self._gen_a += 1
        self._gen_b += 1
        # Soltar la tabla: queda vacía (con sus columnas) para que nada la use por error.
        self.df = self.df.iloc[0:0]
        self._partes = None
        self._resultado = None
        try:
            self.cuerpo_desplazable.desactivar_rueda()
            self.win.destroy()
        except tk.TclError:
            pass
