"""Paso 6 del asistente: Resumen final (opcional) (PLAN §3.2, paso 7).

Pregunta si se quiere agrupar el resultado. Con "Sí" abre el editor de resumen sobre todas las
columnas del resultado (de cualquier archivo), con sus nombres finales. La UI no calcula: la tabla
sobre la que se arma sale de `join.run_join` (una muestra si el resultado es enorme, con aviso).
"""

from __future__ import annotations

import logging
import tkinter as tk
from tkinter import ttk

from excel_tool.core.aggregate import AggregationSpec
from excel_tool.core.errors import ExcelToolError
from excel_tool.ui.desplazable import CuerpoDesplazable
from excel_tool.ui.editor_resumen import EditorResumen
from excel_tool.ui.grilla import formatear_numero
from excel_tool.ui.unir.pasos.base import Paso
from excel_tool.ui.unir.pasos.paso4_modo import texto_resumen
from excel_tool.ui.unir.tamano import TablaParaResumen, nombres_finales, tabla_para_resumen
from excel_tool.ui.worker import MENSAJE_INESPERADO, run_in_background

log = logging.getLogger(__name__)

ANCHO_TEXTO = 760
ALTO_EDITOR = 520
TEXTO_NO = "No, dejarlo así"
TEXTO_SI = "Sí, agruparlo"
TEXTO_PREGUNTA = "¿Querés agrupar el resultado?"
TEXTO_EJEMPLOS = (
    "Sirve, por ejemplo, para ver una fila por Región con el total del Importe, o una fila por "
    "Vendedor y Artículo con lo que vendió cada uno. Podés usar columnas de cualquiera de los "
    "archivos."
)
TEXTO_FALTA = (
    "Armá el agrupamiento (elegí por qué columnas agrupar o agregá un cálculo) "
    "o elegí «No, dejarlo así»."
)
TEXTO_REARMAR = (
    "Como cambiaste las columnas o los archivos, el agrupamiento que tenías ya no sirve: "
    "hay que armarlo de nuevo."
)


def vacio(spec: AggregationSpec | None) -> bool:
    return spec is None or not (spec.group_by or spec.aggregations)


def columnas_usadas(spec: AggregationSpec) -> set[str]:
    usadas = {g.column for g in spec.group_by}
    usadas.update(a.column for a in spec.aggregations if a.column)
    return usadas


class PasoResumenFinal(Paso):
    titulo = "Resumen final"

    def __init__(self, parent, estado, asistente) -> None:
        super().__init__(parent, estado, asistente)
        self.version = 0
        self.ocupado = False
        self.tabla: TablaParaResumen | None = None
        self.dialogo: tk.Toplevel | None = None
        self.editor: EditorResumen | None = None
        self.boton_listo: ttk.Button | None = None
        self.aviso_editor: ttk.Label | None = None
        self._guardado: AggregationSpec | None = None  # por si el usuario vuelve a elegir "Sí"

        ttk.Label(self, text="Resumen final (opcional)", font="TkHeadingFont").pack(anchor="w")
        ttk.Label(self, text=TEXTO_PREGUNTA, wraplength=ANCHO_TEXTO).pack(anchor="w", pady=(2, 6))
        self.var = tk.StringVar(master=self, value="no")
        self.radio_no = ttk.Radiobutton(
            self, text=TEXTO_NO, value="no", variable=self.var, command=self._eligio_no
        )
        self.radio_no.pack(anchor="w", pady=2)
        self.radio_si = ttk.Radiobutton(
            self, text=TEXTO_SI, value="si", variable=self.var, command=self._eligio_si
        )
        self.radio_si.pack(anchor="w", pady=2)
        ttk.Label(self, text=TEXTO_EJEMPLOS, wraplength=ANCHO_TEXTO, foreground="#555555").pack(
            anchor="w", pady=(8, 0)
        )
        self.marco_si = ttk.Frame(self)
        self.etiqueta_resumen = ttk.Label(self.marco_si, text="", wraplength=ANCHO_TEXTO - 200)
        self.etiqueta_resumen.pack(side="left")
        self.boton_armar = ttk.Button(
            self.marco_si, text="Armar el agrupamiento", command=self.abrir_editor
        )
        self.boton_armar.pack(side="left", padx=(10, 0))
        self.etiqueta_estado = ttk.Label(
            self, text="", wraplength=ANCHO_TEXTO, foreground="#c62828"
        )
        self.etiqueta_estado.pack(anchor="w", pady=(8, 0))
        self.bind("<Destroy>", self._al_destruir, add="+")

    # ---- ciclo del paso ----
    def al_entrar(self) -> None:
        self.tabla = None  # lo anterior pudo cambiar: la tabla se vuelve a armar si hace falta
        self.etiqueta_estado.configure(text="")
        self._descartar_si_no_sirve()
        self.var.set("si" if self.estado.resumen_final is not None else "no")
        self._refrescar()

    def validar(self) -> str | None:
        if self.var.get() == "si" and vacio(self.estado.resumen_final):
            return TEXTO_FALTA
        return None

    def al_salir(self) -> None:
        self._cerrar_dialogo()
        if self.var.get() == "no":
            self.estado.resumen_final = None

    def _descartar_si_no_sirve(self) -> None:
        """Si ya no existen las columnas del agrupamiento (cambió algo antes), se descarta."""
        spec = self.estado.resumen_final
        if spec is None:
            return
        try:
            nombres = set(nombres_finales(self.estado.plan(), self.estado.tablas()).values())
        except ExcelToolError:
            return
        if nombres and not columnas_usadas(spec) <= nombres:
            self.estado.resumen_final = None
            self._guardado = None
            self.etiqueta_estado.configure(text=TEXTO_REARMAR)

    def _refrescar(self) -> None:
        spec = self.estado.resumen_final
        if self.var.get() == "si":
            self.marco_si.pack(anchor="w", pady=(10, 0), before=self.etiqueta_estado)
            if vacio(spec):
                self.etiqueta_resumen.configure(text="Todavía no armaste el agrupamiento.")
            else:
                texto = texto_resumen(spec).replace("Resumen armado", "Agrupamiento armado")
                self.etiqueta_resumen.configure(text=texto)
            self.boton_armar.configure(
                text="Armar el agrupamiento" if vacio(spec) else "Cambiar el agrupamiento"
            )
        else:
            self.marco_si.pack_forget()

    # ---- elecciones ----
    def _eligio_no(self) -> None:
        self._cerrar_dialogo()
        if self.estado.resumen_final is not None:
            self._guardado = self.estado.resumen_final
        self.estado.resumen_final = None
        self.etiqueta_estado.configure(text="")
        self._refrescar()

    def _eligio_si(self) -> None:
        if vacio(self.estado.resumen_final) and self._guardado is not None:
            self.estado.resumen_final = self._guardado
        self._refrescar()
        if vacio(self.estado.resumen_final):
            self.abrir_editor()

    # ---- editor ----
    def abrir_editor(self) -> None:
        if self.dialogo is not None:
            self.dialogo.lift()
            return
        if self.ocupado:
            return
        try:
            plan = self.estado.plan()
            tablas = self.estado.tablas()
        except ExcelToolError as exc:
            self.etiqueta_estado.configure(text=exc.mensaje)
            self._sin_resumen_si_vacio()
            return
        if self.tabla is not None:
            self._mostrar_editor(self.tabla)
            return
        self.version += 1
        version = self.version
        self.ocupado = True
        self.etiqueta_estado.configure(text="")

        def trabajo(progress, cancel):
            return tabla_para_resumen(plan, tablas, progress, cancel, self.estado.diagnosticar)

        def listo(tabla: TablaParaResumen) -> None:
            if version != self.version or not self.winfo_exists():
                return
            self.ocupado = False
            self.tabla = tabla
            self._mostrar_editor(tabla)

        def fallo(exc: BaseException) -> None:
            if version != self.version or not self.winfo_exists():
                return
            self.ocupado = False
            if isinstance(exc, ExcelToolError):
                log.warning("No se pudo armar la tabla para agrupar: %s (%s)", exc, exc.detalle)
                texto = exc.mensaje
            else:
                log.error("Error armando la tabla para agrupar", exc_info=exc)
                texto = MENSAJE_INESPERADO
            self.etiqueta_estado.configure(text=texto)
            self._sin_resumen_si_vacio()

        def cancelado() -> None:
            if version == self.version and self.winfo_exists():
                self.ocupado = False
                self._sin_resumen_si_vacio()

        run_in_background(
            self,
            trabajo,
            title="Preparando el resultado para agruparlo",
            on_success=listo,
            on_error=fallo,
            on_cancel=cancelado,
        )

    def _sin_resumen_si_vacio(self) -> None:
        if vacio(self.estado.resumen_final):
            self.var.set("no")
            self._refrescar()

    def _mostrar_editor(self, tabla: TablaParaResumen) -> None:
        previo = self.estado.resumen_final
        top = tk.Toplevel(self)
        top.title("Agrupar el resultado")
        top.transient(self.winfo_toplevel())
        self.dialogo = top
        cuerpo = CuerpoDesplazable(top, alto_maximo=ALTO_EDITOR)
        ttk.Label(
            cuerpo.marco,
            text="Armá cómo querés agrupar el resultado.",
            font="TkHeadingFont",
            wraplength=ANCHO_TEXTO,
        ).pack(anchor="w", pady=(0, 4))
        if tabla.es_muestra:
            ttk.Label(
                cuerpo.marco,
                text=(
                    f"El resultado es muy grande ({formatear_numero(tabla.filas_total)} filas). "
                    f"Para armar el agrupamiento se usa una muestra de "
                    f"{formatear_numero(len(tabla.df))} filas: los números de la vista previa "
                    "son de esa muestra. Al unir se calcula con todo."
                ),
                wraplength=ANCHO_TEXTO,
                foreground="#8a5a00",
            ).pack(anchor="w", pady=(0, 6))
        editor = EditorResumen(cuerpo.marco, tabla.df, tabla.column_types)
        editor.pack(fill="both", expand=True)
        if not vacio(previo):
            editor.set_spec(previo)
        self.editor = editor
        aviso = ttk.Label(cuerpo.pie, text="", foreground="#c62828", wraplength=420)
        aviso.pack(side="left", fill="x", expand=True)
        self.aviso_editor = aviso

        def listo() -> None:
            if not editor.hay_algo_que_resumir():
                aviso.configure(text="Elegí por qué columna agrupar o agregá al menos un cálculo.")
                return
            self.estado.resumen_final = editor.spec()
            self._guardado = None
            self._cerrar_dialogo()
            self.var.set("si")
            self._refrescar()

        def cancelar() -> None:
            self._cerrar_dialogo()
            self._sin_resumen_si_vacio()

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

    def _al_destruir(self, evento: tk.Event) -> None:
        if evento.widget is self:
            self.version += 1
