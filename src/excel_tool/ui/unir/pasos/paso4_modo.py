"""Paso 4 del asistente: ¿Detalle o resumen? (PLAN §3.2, paso 4).

Una tarjeta por archivo sumado. Nunca hay una opción elegida de entrada. La UI no calcula: los
ejemplos, las filas y el aviso de "más de un detalle" salen de `join.diagnose_join`, y el aviso de
tamaño de `limits.check`; ambos corren en el worker.
"""

from __future__ import annotations

import dataclasses
import logging
import re
import tkinter as tk
from tkinter import ttk
from typing import Any

from excel_tool.core import limits
from excel_tool.core.aggregate import Aggregation, AggregationSpec, GroupBy
from excel_tool.core.errors import ExcelToolError
from excel_tool.core.types import detect_types
from excel_tool.ui.desplazable import CuerpoDesplazable
from excel_tool.ui.editor_resumen import EditorResumen
from excel_tool.ui.grilla import formatear_numero
from excel_tool.ui.unir.estado import FuenteSecundaria, Modo
from excel_tool.ui.unir.pasos.base import Paso
from excel_tool.ui.worker import MENSAJE_INESPERADO, run_in_background

log = logging.getLogger(__name__)

ANCHO_TEXTO = 760
ALTO_EDITOR = 520
TIPOS_SUMABLES = {"entero", "decimal"}
TEXTO_DETALLE = "🧾 Ver el detalle (una fila por cada registro)"
FRASE_DETALLE = "Para saber qué compró cada cliente"
TEXTO_RESUMEN = "📊 Ver un resumen"
FRASE_RESUMEN = "Para saber cuánto compró cada cliente"
TEXTO_PRIMERA = "1️⃣ Traer solo la primera fila"
TEXTO_MAS_OPCIONES = "Más opciones"
TEXTO_VARIOS_DETALLES = (
    "Si traés el detalle de dos archivos a la vez, las filas se combinan todas con todas y los "
    "totales quedan inflados. Te recomendamos ver el detalle de uno solo y el resumen de los otros."
)
ICONOS_TAMANO = {"recomendado": "🟡", "obligatorio": "🔴"}
TEXTO_TAMANO = {
    "recomendado": "va a quedar muy pesado para una sola hoja de Excel",
    "obligatorio": "no entra en una sola hoja de Excel",
}


def _sustantivo(columna: str) -> str:
    """'ID cliente' -> 'cliente' (para armar frases como 'El cliente 1045')."""
    resto = re.sub(r"^(id|nro\.?|n[úu]mero|c[óo]digo|cod\.?)\s+", "", columna.strip(), flags=re.I)
    return (resto.strip() or columna.strip()).lower()


def _articulo(sustantivo: str) -> str:
    if " " not in sustantivo and sustantivo.endswith("a") and not sustantivo.endswith("ma"):
        return "La"
    return "El"


def _tipos_de(fuente: Any) -> dict[str, str]:
    """Tipos por columna: del perfil si lo hay; si no, los detecta el core."""
    if fuente.perfil is not None:
        return {c.name: c.detected_type for c in fuente.perfil.columns_profile}
    return {nombre: det.detected for nombre, det in detect_types(fuente.df).items()}


def resumen_por_defecto(sec: FuenteSecundaria) -> AggregationSpec:
    """Cantidad de filas + sumar las columnas numéricas, agrupando por la clave del archivo."""
    claves = [b for _a, b in sec.key_pairs]
    tipos = _tipos_de(sec)
    calculos = [Aggregation("", "count")]
    for nombre, tipo in tipos.items():
        if tipo in TIPOS_SUMABLES and nombre not in claves:
            calculos.append(Aggregation(nombre, "sum"))
    return AggregationSpec([GroupBy(c) for c in claves], calculos)


def _vacio(spec: AggregationSpec | None) -> bool:
    return spec is None or not (spec.group_by or spec.aggregations)


def texto_resumen(spec: AggregationSpec) -> str:
    grupos = [g.column for g in spec.group_by]
    cuenta = len(spec.aggregations)
    calculos = "1 cálculo" if cuenta == 1 else f"{cuenta} cálculos"
    if not grupos:
        return f"Resumen armado: una sola fila con {calculos}."
    return f"Resumen armado: una fila por cada {' y '.join(grupos)}, con {calculos}."


class _Tarjeta:
    """Los widgets de un archivo sumado."""

    def __init__(self, paso: PasoModo, sec: FuenteSecundaria) -> None:
        self.paso = paso
        self.nombre = sec.nombre
        self.marco = ttk.LabelFrame(paso.contenedor, text=sec.nombre, padding=10)
        self.marco.pack(fill="x", pady=(0, 10))
        self.etiqueta_ejemplo = ttk.Label(
            self.marco, text="Mirando los datos…", wraplength=ANCHO_TEXTO, font="TkHeadingFont"
        )
        self.etiqueta_ejemplo.pack(anchor="w", pady=(0, 6))
        self.opciones = ttk.Frame(self.marco)
        self.opciones.pack(fill="x")
        self.var = tk.StringVar(master=paso, value=sec.mode or "")
        self.radio_detalle = self._opcion(TEXTO_DETALLE, FRASE_DETALLE, "detail", sec)
        self.etiqueta_filas = ttk.Label(self.opciones, text="", wraplength=ANCHO_TEXTO - 30)
        self.etiqueta_filas.pack(anchor="w", padx=(26, 0))
        self.radio_resumen = self._opcion(TEXTO_RESUMEN, FRASE_RESUMEN, "summary", sec)
        self.marco_resumen = ttk.Frame(self.opciones)
        self.marco_resumen.pack(fill="x", padx=(26, 0))
        self.etiqueta_resumen = ttk.Label(self.marco_resumen, text="", wraplength=ANCHO_TEXTO - 150)
        self.etiqueta_resumen.pack(side="left")
        self.boton_editar = ttk.Button(
            self.marco_resumen, text="Editar el resumen", command=lambda: paso.editar_resumen(self)
        )
        self.mas = ttk.Frame(self.opciones)
        self.mas.pack(fill="x", pady=(6, 0))
        self.var_mas = tk.BooleanVar(master=paso, value=False)
        self.boton_mas = ttk.Button(self.mas, command=self.alternar_mas)
        self.boton_mas.pack(anchor="w")
        self.cuerpo_mas = ttk.Frame(self.mas)
        self.radio_primera = ttk.Radiobutton(
            self.cuerpo_mas,
            text=TEXTO_PRIMERA,
            value="first",
            variable=self.var,
            command=lambda: paso.elegir(self, "first"),
        )
        self.radio_primera.pack(anchor="w", padx=(12, 0))
        self.etiqueta_sola_vez = ttk.Label(self.marco, text="", wraplength=ANCHO_TEXTO)
        self._mas_abierto = False
        self.alternar_mas(abrir=sec.mode == "first")

    def _opcion(self, texto: str, frase: str, valor: str, sec: FuenteSecundaria) -> ttk.Radiobutton:
        radio = ttk.Radiobutton(
            self.opciones,
            text=texto,
            value=valor,
            variable=self.var,
            command=lambda: self.paso.elegir(self, valor),
        )
        radio.pack(anchor="w", pady=(4, 0))
        ttk.Label(self.opciones, text=frase, foreground="#555555").pack(anchor="w", padx=(26, 0))
        return radio

    def alternar_mas(self, abrir: bool | None = None) -> None:
        self._mas_abierto = (not self._mas_abierto) if abrir is None else abrir
        if self._mas_abierto:
            self.cuerpo_mas.pack(fill="x")
        else:
            self.cuerpo_mas.pack_forget()
        flecha = "▾" if self._mas_abierto else "▸"
        self.boton_mas.configure(text=f"{flecha} {TEXTO_MAS_OPCIONES}")

    def mostrar_opciones(self, mostrar: bool) -> None:
        if mostrar:
            self.etiqueta_sola_vez.pack_forget()
            if not self.opciones.winfo_manager():
                self.opciones.pack(fill="x")
        else:
            self.opciones.pack_forget()
            self.etiqueta_sola_vez.pack(anchor="w")

    def refrescar(self, sec: FuenteSecundaria) -> None:
        self.var.set(sec.mode or "")
        if sec.mode == "first" and not self._mas_abierto:
            self.alternar_mas(abrir=True)
        if sec.mode == "summary" and sec.summary is not None:
            self.etiqueta_resumen.configure(text=texto_resumen(sec.summary))
            self.boton_editar.pack(side="left", padx=(10, 0))
        else:
            self.etiqueta_resumen.configure(text="")
            self.boton_editar.pack_forget()
        if sec.mode != "detail":
            self.etiqueta_filas.configure(text="")


class PasoModo(Paso):
    titulo = "¿Detalle o resumen?"

    def __init__(self, parent, estado, asistente) -> None:
        super().__init__(parent, estado, asistente)
        ttk.Label(self, text=self.titulo, font="TkHeadingFont").pack(anchor="w")
        ttk.Label(
            self,
            text="Para cada archivo que sumaste, elegí qué querés traer.",
            wraplength=ANCHO_TEXTO,
        ).pack(anchor="w", pady=(0, 8))
        self.tarjetas: list[_Tarjeta] = []
        self._firma: list[str] | None = None
        self.automaticos: set[str] = set()  # archivos que quedaron en "primera fila" sin preguntar
        self.diagnostico: Any = None
        self.chequeo: Any = None
        self.version = 0
        self.pendiente = False
        self.dialogo: tk.Toplevel | None = None
        self.editor: EditorResumen | None = None
        self.marco_avisos = ttk.Frame(self)
        self.marco_avisos.pack(fill="x", side="bottom")
        self.etiqueta_total = ttk.Label(self.marco_avisos, text="", wraplength=ANCHO_TEXTO)
        self.etiqueta_total.pack(anchor="w")
        self.etiqueta_tamano = ttk.Label(
            self.marco_avisos, text="", wraplength=ANCHO_TEXTO, foreground="#8a5a00"
        )
        self.etiqueta_tamano.pack(anchor="w", pady=(4, 0))
        self.marco_varios = ttk.Frame(self.marco_avisos)
        self.etiqueta_varios = ttk.Label(
            self.marco_varios, text=f"⚠️ {TEXTO_VARIOS_DETALLES}", wraplength=ANCHO_TEXTO
        )
        self.etiqueta_varios.pack(anchor="w")
        self.etiqueta_varios_filas = ttk.Label(self.marco_varios, text="", wraplength=ANCHO_TEXTO)
        self.etiqueta_varios_filas.pack(anchor="w", pady=(2, 0))
        self.botones_cambiar: list[ttk.Button] = []
        self.marco_botones = ttk.Frame(self.marco_varios)
        self.marco_botones.pack(anchor="w", pady=(4, 0))
        self.contenedor = ttk.Frame(self)
        self.contenedor.pack(fill="both", expand=True)

    # ---- ciclo del paso ----
    def al_entrar(self) -> None:
        for sec in self.estado.secundarios:
            if sec.nombre in self.automaticos and sec.mode == "first":
                sec.mode = None  # la clave pudo cambiar en el paso 3: se vuelve a decidir
        self.automaticos.clear()
        firma = [s.nombre for s in self.estado.secundarios]
        if firma != self._firma:
            for tarjeta in self.tarjetas:
                tarjeta.marco.destroy()
            self.tarjetas = []
            for sec in self.estado.secundarios:
                tarjeta = _Tarjeta(self, sec)
                self.tarjetas.append(tarjeta)
            self._firma = firma
        for tarjeta, sec in zip(self.tarjetas, self.estado.secundarios, strict=True):
            tarjeta.refrescar(sec)
        self.recalcular()

    def validar(self) -> str | None:
        if not self.estado.secundarios:
            return "Primero sumá al menos un archivo."
        for sec in self.estado.secundarios:
            if sec.mode is None:
                return f"Elegí si de «{sec.nombre}» querés ver el detalle o un resumen."
            if sec.mode == "summary" and _vacio(sec.summary):
                return f"Falta armar el resumen de «{sec.nombre}»."
        return None

    def al_salir(self) -> None:
        self._cerrar_dialogo()

    # ---- elecciones ----
    def _tarjeta_y_sec(self, tarjeta: _Tarjeta) -> FuenteSecundaria:
        indice = self.tarjetas.index(tarjeta)
        return self.estado.secundarios[indice]

    def elegir(self, tarjeta: _Tarjeta, modo: Modo) -> None:
        sec = self._tarjeta_y_sec(tarjeta)
        if modo == "summary":
            self._abrir_editor(tarjeta, sec, previo=(sec.mode, sec.summary), inicial=sec.summary)
            return
        sec.mode = modo
        tarjeta.refrescar(sec)
        self.recalcular()

    def editar_resumen(self, tarjeta: _Tarjeta) -> None:
        sec = self._tarjeta_y_sec(tarjeta)
        self._abrir_editor(tarjeta, sec, previo=(sec.mode, sec.summary), inicial=sec.summary)

    def cambiar_a_resumen(self, nombre: str) -> None:
        """Un clic: resumen por defecto y se abre el editor para ajustarlo."""
        for tarjeta, sec in zip(self.tarjetas, self.estado.secundarios, strict=True):
            if sec.nombre == nombre:
                sec.summary = resumen_por_defecto(sec)
                sec.mode = "summary"
                tarjeta.refrescar(sec)
                self.recalcular()
                self._abrir_editor(
                    tarjeta, sec, previo=(sec.mode, sec.summary), inicial=sec.summary
                )
                return

    # ---- editor de resumen ----
    def _abrir_editor(
        self,
        tarjeta: _Tarjeta,
        sec: FuenteSecundaria,
        *,
        previo: tuple[Modo | None, AggregationSpec | None],
        inicial: AggregationSpec | None,
    ) -> None:
        if self.dialogo is not None:
            self.dialogo.lift()
            return
        top = tk.Toplevel(self)
        top.title(f"Resumen de {sec.nombre}")
        top.transient(self.winfo_toplevel())
        self.dialogo = top
        cuerpo = CuerpoDesplazable(top, alto_maximo=ALTO_EDITOR)
        ttk.Label(
            cuerpo.marco,
            text=f"Armá el resumen de «{sec.nombre}».",
            font="TkHeadingFont",
            wraplength=ANCHO_TEXTO,
        ).pack(anchor="w", pady=(0, 8))
        editor = EditorResumen(
            cuerpo.marco,
            sec.df,
            _tipos_de(sec),
            fixed_group=[b for _a, b in sec.key_pairs],
        )
        editor.pack(fill="both", expand=True)
        if inicial is not None:
            editor.set_spec(inicial)
        self.editor = editor
        aviso = ttk.Label(cuerpo.pie, text="", foreground="#c62828", wraplength=420)
        aviso.pack(side="left", fill="x", expand=True)
        self.aviso_editor = aviso

        def listo() -> None:
            if not editor.hay_algo_que_resumir():
                aviso.configure(text="Elegí por qué columna agrupar o agregá al menos un cálculo.")
                return
            sec.summary = editor.spec()
            sec.mode = "summary"
            self._cerrar_dialogo()
            tarjeta.refrescar(sec)
            self.recalcular()

        def cancelar() -> None:
            sec.mode, sec.summary = previo
            self._cerrar_dialogo()
            tarjeta.refrescar(sec)

        ttk.Button(cuerpo.pie, text="Cancelar", command=cancelar).pack(side="right", padx=(6, 0))
        self.boton_listo = ttk.Button(cuerpo.pie, text="Listo", command=listo)
        self.boton_listo.pack(side="right")
        top.protocol("WM_DELETE_WINDOW", cancelar)
        top.update_idletasks()
        top.geometry(f"{max(top.winfo_reqwidth(), 820)}x{min(top.winfo_reqheight(), 700)}")

    def _cerrar_dialogo(self) -> None:
        dialogo, self.dialogo, self.editor = self.dialogo, None, None
        if dialogo is None:
            return
        try:
            dialogo.destroy()
        except tk.TclError:
            pass

    # ---- diagnóstico (en el worker, con contador de versión) ----
    def _plan_para_diagnosticar(self):
        """El plan con lo elegido; a los que todavía no eligieron se les pone 'primera fila'
        solo para poder mirar sus datos (no cambia lo que ve ni lo que guarda el usuario)."""
        secundarios = [
            s
            if s.mode is not None and not (s.mode == "summary" and _vacio(s.summary))
            else dataclasses.replace(s, mode="first", summary=None)
            for s in self.estado.secundarios
        ]
        return dataclasses.replace(self.estado, secundarios=secundarios).plan()

    def recalcular(self) -> None:
        self.version += 1
        version = self.version
        try:
            plan = self._plan_para_diagnosticar()
            tablas = self.estado.tablas()
        except ExcelToolError as exc:
            self.pendiente = False
            self.etiqueta_total.configure(text=exc.mensaje)
            return
        self.pendiente = True
        self.etiqueta_total.configure(text="Calculando cuántas filas va a tener…")

        def trabajo(progress, cancel):
            diag = self.estado.diagnosticar(plan, tablas, progress=progress, cancel=cancel)
            return diag, limits.check(diag.size, "excel", "xlsx")

        def listo(resultado) -> None:
            if version != self.version or not self.winfo_exists():
                return  # llegó tarde: ya cambió algo y hay otro cálculo en camino
            self.pendiente = False
            self._mostrar(*resultado)

        def fallo(exc: BaseException) -> None:
            if version != self.version or not self.winfo_exists():
                return
            self.pendiente = False
            mensaje = exc.mensaje if isinstance(exc, ExcelToolError) else MENSAJE_INESPERADO
            self.etiqueta_total.configure(text=mensaje)

        run_in_background(
            self,
            trabajo,
            title="Mirando tus archivos",
            on_success=listo,
            on_error=fallo,
            cancellable=False,
            modal=False,
        )

    def _mostrar(self, diag: Any, chequeo: Any) -> None:
        self.diagnostico, self.chequeo = diag, chequeo
        claves = self.estado.claves or [a for a, _b in self.estado.secundarios[0].key_pairs[:1]]
        sustantivo = _sustantivo(claves[0]) if len(claves) == 1 else "combinación"
        articulo = _articulo(sustantivo) if len(claves) == 1 else "La"
        for tarjeta, sec, dsec in zip(
            self.tarjetas, self.estado.secundarios, diag.secondaries, strict=True
        ):
            if dsec.example_many is None:
                tarjeta.etiqueta_ejemplo.configure(
                    text=f"Cada {sustantivo} aparece una sola vez: se trae una fila por "
                    f"{sustantivo}"
                )
                tarjeta.etiqueta_sola_vez.configure(text="")
                if sec.mode is None:
                    sec.mode = "first"  # decisión de nivel 1: sin repeticiones no hay qué elegir
                    self.automaticos.add(sec.nombre)
                if sec.nombre in self.automaticos:
                    tarjeta.mostrar_opciones(False)
                    tarjeta.etiqueta_sola_vez.configure(
                        text="Se trae una fila por cada uno, no hace falta elegir."
                    )
                    tarjeta.var.set("first")
            else:
                valor, veces = dsec.example_many
                tarjeta.etiqueta_ejemplo.configure(
                    text=f"{articulo} {sustantivo} {valor} aparece {formatear_numero(veces)} "
                    f"veces en {sec.nombre}"
                )
                tarjeta.mostrar_opciones(True)
            if sec.mode == "detail":
                tarjeta.etiqueta_filas.configure(
                    text=f"El resultado va a tener {formatear_numero(dsec.rows_after)} filas."
                )
        self._mostrar_avisos(diag, chequeo)

    def _mostrar_avisos(self, diag: Any, chequeo: Any) -> None:
        filas = formatear_numero(diag.rows_result)
        hay_detalle = any(s.mode == "detail" for s in self.estado.secundarios)
        if hay_detalle:
            texto = f"Con lo que elegiste, el resultado va a tener {filas} filas."
        else:
            texto = (
                f"Con lo que elegiste, el resultado va a tener {filas} filas "
                "(las mismas que el archivo principal)."
            )
        self.etiqueta_total.configure(text=texto if self.estado.secundarios else "")
        nivel = getattr(chequeo, "level", "ok")
        if nivel in ICONOS_TAMANO:
            self.etiqueta_tamano.configure(
                text=f"{ICONOS_TAMANO[nivel]} El resultado va a tener ~{filas} filas: "
                f"{TEXTO_TAMANO[nivel]}. Al guardar te vamos a proponer dividirlo."
            )
        else:
            self.etiqueta_tamano.configure(text="")
        self._mostrar_varios(diag)

    def _mostrar_varios(self, diag: Any) -> None:
        for boton in self.botones_cambiar:
            boton.destroy()
        self.botones_cambiar = []
        if not diag.many_details:
            self.marco_varios.pack_forget()
            return
        self.etiqueta_varios_filas.configure(
            text=f"Si lo dejás así, el resultado va a tener {formatear_numero(diag.rows_result)} "
            "filas."
        )
        detalles = [s.nombre for s in self.estado.secundarios if s.mode == "detail"]
        for nombre in detalles[1:]:
            boton = ttk.Button(
                self.marco_botones,
                text=f"Cambiar {nombre} a resumen",
                command=lambda n=nombre: self.cambiar_a_resumen(n),
            )
            boton.pack(side="left", padx=(0, 8))
            self.botones_cambiar.append(boton)
        self.marco_varios.pack(fill="x", pady=(6, 0), before=self.etiqueta_tamano)
