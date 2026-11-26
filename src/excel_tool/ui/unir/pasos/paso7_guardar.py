"""Paso 7 del asistente: Revisar y guardar (PLAN §3.2 paso 8, §3.3 puntos 4 y 6).

Muestra el diagnóstico, une, muestra una vista previa y guarda. La UI no calcula: coincidencias,
filas, tamaños y alternativas salen de `join.diagnose_join` y `join.suggest_reductions`; la unión,
de `join.run_join`; el guardado, de `ui/guardar.py` (`output.save_result`) y
`output.save_workbook`. Todo lo largo corre en el worker.
"""

from __future__ import annotations

import dataclasses
import gc
import logging
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from excel_tool.core.errors import Cancelado, ExcelToolError
from excel_tool.core.join import JoinResult, run_join
from excel_tool.core.output import SheetSpec, plan_workbook, save_workbook
from excel_tool.core.preferences import load_preferences, save_preferences
from excel_tool.core.recipe import save_recipe
from excel_tool.ui import guardar
from excel_tool.ui.explorador import abrir_explorador
from excel_tool.ui.grilla import Grilla, formatear_numero
from excel_tool.ui.unir.pasos.base import Paso
from excel_tool.ui.unir.tamano import Revision, revisar
from excel_tool.ui.worker import MENSAJE_INESPERADO, run_in_background

log = logging.getLogger(__name__)

ANCHO_TEXTO = 760
FILAS_VISTA_PREVIA = 50
ICONOS = {"ok": "🟢", "recomendado": "🟡", "obligatorio": "🔴"}
TEXTO_CONTROL_IGUAL = (
    "Control de filas: el resultado va a tener las mismas filas que el archivo principal "
    "({filas}). Lo verificamos al unir y, si no coincide, no se guarda nada."
)
TEXTO_CONTROL_DISTINTO = (
    "Control de filas: se esperan {filas} filas (el archivo principal tiene {base}) "
    "porque {motivo}. Al unir verificamos que el resultado tenga exactamente esa cantidad; "
    "si no, no se guarda nada."
)
TEXTO_TITULO_RESULTADO = "Resultado de la unión"
TEXTO_LIBRO_GRANDE = (
    "Alguna de las hojas no entra en un Excel. Guardá el resultado por separado con "
    "«Guardar el resultado» (ahí podés dividirlo en varios archivos) y, si querés, "
    "el resumen o los que no tuvieron coincidencia en otro archivo."
)


def _motivos_de_filas(estado: Any) -> list[str]:
    """Por qué el resultado puede tener otras filas que el principal (solo para explicarlo)."""
    motivos: list[str] = []
    detalles = [s.nombre for s in estado.secundarios if s.mode == "detail"]
    if detalles:
        motivos.append("traés el detalle de " + " y de ".join(f"«{d}»" for d in detalles))
    if any(s.how == "inner" for s in estado.secundarios):
        motivos.append("quedan solo los que están en ambos archivos")
    for s in estado.secundarios:
        if s.mode == "summary" and s.summary is not None:
            claves = {b for _a, b in s.key_pairs}
            if any(g.column not in claves or g.period for g in s.summary.group_by):
                motivos.append(f"el resumen de «{s.nombre}» agrupa por más que la columna en común")
    if estado.resumen_final is not None:
        motivos.append("agrupaste el resultado")
    return motivos


class PasoGuardar(Paso):
    titulo = "Revisar y guardar"

    def __init__(self, parent, estado, asistente) -> None:
        super().__init__(parent, estado, asistente)
        self.version = 0  # sube con cada cambio: un resultado de antes se descarta
        self.revision: Revision | None = None
        self.resultado: JoinResult | None = None
        self.pendiente = False
        self.uniendo = False
        self.ventana_guardar: guardar.VentanaGuardar | None = None
        self.botones_reduccion: list[ttk.Button] = []
        self.etiquetas_archivo: list[ttk.Label] = []
        self.explorador: Any = None

        ttk.Label(self, text=self.titulo, font="TkHeadingFont").pack(anchor="w")
        # ---- diagnóstico ----
        self.marco_diag = ttk.LabelFrame(self, text="Antes de unir", padding=8)
        self.marco_diag.pack(fill="x", pady=(4, 6))
        self.etiqueta_total = ttk.Label(
            self.marco_diag, text="", wraplength=ANCHO_TEXTO, font="TkHeadingFont"
        )
        self.etiqueta_total.pack(anchor="w")
        self.marco_archivos = ttk.Frame(self.marco_diag)
        self.marco_archivos.pack(fill="x", pady=(4, 0))
        self.etiqueta_control = ttk.Label(self.marco_diag, text="", wraplength=ANCHO_TEXTO)
        self.etiqueta_control.pack(anchor="w", pady=(4, 0))
        self.etiqueta_avisos = ttk.Label(
            self.marco_diag, text="", wraplength=ANCHO_TEXTO, foreground="#8a5a00"
        )
        self.etiqueta_avisos.pack(anchor="w", pady=(4, 0))
        self.etiqueta_semaforo = ttk.Label(self.marco_diag, text="", wraplength=ANCHO_TEXTO)
        self.etiqueta_semaforo.pack(anchor="w", pady=(4, 0))
        # ---- achicar ----
        self.marco_achicar = ttk.LabelFrame(self, text="Achicar el resultado", padding=8)
        self.etiqueta_achicar = ttk.Label(
            self.marco_achicar,
            text="El resultado es muy grande para Excel. Probá con alguna de estas opciones:",
            wraplength=ANCHO_TEXTO,
        )
        self.etiqueta_achicar.pack(anchor="w")
        self.marco_botones_achicar = ttk.Frame(self.marco_achicar)
        self.marco_botones_achicar.pack(fill="x", pady=(4, 0))
        # ---- unir ----
        fila = ttk.Frame(self)
        fila.pack(fill="x", pady=(2, 4))
        self.boton_unir = ttk.Button(fila, text="Unir", command=self.unir)
        self.boton_unir.pack(side="left")
        self.etiqueta_estado = ttk.Label(
            fila, text="", wraplength=ANCHO_TEXTO - 120, foreground="#c62828"
        )
        self.etiqueta_estado.pack(side="left", padx=10)
        # ---- resultado ----
        self.marco_resultado = ttk.LabelFrame(self, text="Resultado", padding=8)
        self.etiqueta_filas = ttk.Label(
            self.marco_resultado, text="", wraplength=ANCHO_TEXTO, font="TkHeadingFont"
        )
        self.etiqueta_filas.pack(anchor="w")
        self.etiqueta_avisos_resultado = ttk.Label(
            self.marco_resultado, text="", wraplength=ANCHO_TEXTO, foreground="#8a5a00"
        )
        self.etiqueta_avisos_resultado.pack(anchor="w")
        self.grilla = Grilla(self.marco_resultado, None, page_size=FILAS_VISTA_PREVIA)
        self.grilla.pack(fill="both", expand=True, pady=(4, 4))
        self.boton_informacion = ttk.Button(
            self.marco_resultado,
            text="🔍 Ver información del resultado",
            command=self.ver_informacion,
        )
        self.boton_informacion.pack(anchor="w")
        self.marco_guardar = ttk.Frame(self.marco_resultado)
        self.marco_guardar.pack(fill="x", pady=(6, 0))
        self.boton_guardar = ttk.Button(
            self.marco_guardar, text="Guardar el resultado", command=self.guardar_resultado
        )
        self.boton_guardar.pack(side="left")
        self.boton_libro = ttk.Button(
            self.marco_guardar,
            text="Guardar en un Excel con varias hojas",
            command=self.guardar_libro,
        )
        # ---- receta y avisos de guardado ----
        self.etiqueta_guardado = ttk.Label(self, text="", wraplength=ANCHO_TEXTO)
        self.etiqueta_guardado.pack(anchor="w", pady=(6, 0))
        self.boton_receta = ttk.Button(
            self, text="Guardar esta receta…", command=self.guardar_receta
        )
        self.boton_receta.pack(anchor="w", pady=(6, 0))
        ttk.Label(
            self,
            text="La receta recuerda lo que elegiste para repetirlo después con otros archivos.",
            foreground="#555555",
            wraplength=ANCHO_TEXTO,
        ).pack(anchor="w")
        self.bind("<Destroy>", self._al_destruir, add="+")

    # ---- ciclo del paso ----
    def al_entrar(self) -> None:
        self.soltar_resultado()
        self.etiqueta_guardado.configure(text="")
        self.etiqueta_estado.configure(text="")
        self.diagnosticar()

    def al_salir(self) -> None:
        self.soltar_resultado()

    def soltar_resultado(self) -> None:
        """Suelta el resultado unido (cientos de MB en archivos grandes) y todo lo que lo
        referencia: la grilla, el explorador y la ventana de guardar. Después recolecta."""
        tenia = self.resultado is not None or self.ventana_guardar is not None
        ventana, self.ventana_guardar = self.ventana_guardar, None
        self.resultado = None
        self.explorador = None
        self._ocultar_resultado()
        if ventana is not None:
            ventana.cerrar()
        if tenia:
            gc.collect()

    def validar(self) -> str | None:
        return None

    def _ocultar_resultado(self) -> None:
        try:
            self.marco_resultado.pack_forget()
            self.grilla.set_proveedor(None)
        except tk.TclError:
            pass  # el asistente se está cerrando

    # ---- diagnóstico ----
    def diagnosticar(self) -> None:
        self.version += 1
        version = self.version
        self.revision = None
        self._mostrar_reducciones([])
        self.etiqueta_semaforo.configure(text="")
        self.etiqueta_avisos.configure(text="")
        self.etiqueta_control.configure(text="")
        for e in self.etiquetas_archivo:
            e.destroy()
        self.etiquetas_archivo = []
        try:
            plan = self.estado.plan()
            tablas = self.estado.tablas()
        except ExcelToolError as exc:
            self.pendiente = False
            self.etiqueta_total.configure(text=exc.mensaje)
            self.boton_unir.state(["disabled"])
            return
        self.pendiente = True
        self.boton_unir.state(["disabled"])
        self.etiqueta_total.configure(text="Calculando cuántas filas va a tener…")

        def trabajo(progress, cancel):
            return revisar(plan, tablas, progress, cancel, self.estado.diagnosticar)

        def listo(revision: Revision) -> None:
            if version != self.version or not self.winfo_exists():
                return
            self.pendiente = False
            self.revision = revision
            self.boton_unir.state(["!disabled"])
            self._mostrar_revision(revision)

        def fallo(exc: BaseException) -> None:
            if version != self.version or not self.winfo_exists() or isinstance(exc, Cancelado):
                return
            self.pendiente = False
            self.etiqueta_total.configure(text=self._mensaje(exc))
            self.boton_unir.state(["!disabled"])  # al unir, el core vuelve a explicar el motivo

        run_in_background(
            self,
            trabajo,
            title="Mirando tus archivos",
            on_success=listo,
            on_error=fallo,
            cancellable=False,
            modal=False,
        )

    def _mensaje(self, exc: BaseException) -> str:
        if isinstance(exc, ExcelToolError):
            log.warning("Paso Revisar y guardar: %s (%s)", exc, exc.detalle)
            return exc.mensaje
        log.error("Error inesperado en el paso Revisar y guardar", exc_info=exc)
        return MENSAJE_INESPERADO

    def _mostrar_revision(self, rev: Revision) -> None:
        diag = rev.diagnostico
        self.etiqueta_total.configure(
            text=(
                f"El resultado va a tener {formatear_numero(diag.rows_result)} filas y "
                f"{formatear_numero(diag.columns_result)} columnas."
            )
        )
        for sec, d in zip(self.estado.secundarios, diag.secondaries, strict=True):
            sin_pareja = d.keys_principal - d.keys_matched
            texto = (
                f"• {sec.nombre}: {formatear_numero(d.keys_matched)} de "
                f"{formatear_numero(d.keys_principal)} con coincidencia, "
                f"{formatear_numero(sin_pareja)} sin pareja"
            )
            if d.secondary_orphans:
                texto += (
                    f"; {formatear_numero(d.secondary_orphans)} del archivo no están en el "
                    "principal"
                )
            texto += f". Filas hasta acá: {formatear_numero(d.rows_after)}."
            etiqueta = ttk.Label(self.marco_archivos, text=texto, wraplength=ANCHO_TEXTO)
            etiqueta.pack(anchor="w")
            self.etiquetas_archivo.append(etiqueta)
        base = len(self.estado.principal.df) if self.estado.principal is not None else 0
        motivos = _motivos_de_filas(self.estado)
        if diag.rows_result == base and not motivos:
            control = TEXTO_CONTROL_IGUAL.format(filas=formatear_numero(diag.rows_result))
        else:
            control = TEXTO_CONTROL_DISTINTO.format(
                filas=formatear_numero(diag.rows_result),
                base=formatear_numero(base),
                motivo=" y ".join(motivos) if motivos else "así lo calculamos con lo que elegiste",
            )
        self.etiqueta_control.configure(text=control)
        avisos = list(diag.warnings)
        if not diag.memory.ok:
            avisos.append(diag.memory.message)
        self.etiqueta_avisos.configure(text="\n".join(f"⚠️ {a}" for a in avisos))
        nivel = rev.excel.level
        motivo = rev.excel.reasons[0] if rev.excel.reasons else "entra sin problemas."
        self.etiqueta_semaforo.configure(text=f"{ICONOS[nivel]} Excel: {motivo}")
        self._mostrar_reducciones(rev.reducciones)

    # ---- achicar el resultado ----
    def _mostrar_reducciones(self, reducciones: list) -> None:
        for b in self.botones_reduccion:
            b.destroy()
        self.botones_reduccion = []
        if not reducciones:
            self.marco_achicar.pack_forget()
            return
        for reduccion, chequeo in reducciones:
            texto = (
                f"{reduccion.description} → {formatear_numero(reduccion.rows)} filas "
                f"{ICONOS[chequeo.level]}"
            )
            boton = ttk.Button(
                self.marco_botones_achicar,
                text=texto,
                command=lambda r=reduccion: self.aplicar_reduccion(r),
            )
            boton.pack(anchor="w", pady=2)
            self.botones_reduccion.append(boton)
        self.marco_achicar.pack(fill="x", pady=(0, 6), before=self.boton_unir.master)

    def aplicar_reduccion(self, reduccion: Any) -> None:
        """Aplica el plan alternativo al estado y vuelve a diagnosticar."""
        self.estado.aplicar_plan(reduccion.plan)
        self.soltar_resultado()
        self.etiqueta_guardado.configure(text=f"Listo: {reduccion.description}.")
        self.diagnosticar()

    # ---- unir ----
    def unir(self) -> None:
        if self.uniendo:
            return
        try:
            plan = self.estado.plan()
            tablas = self.estado.tablas()
        except ExcelToolError as exc:
            self.etiqueta_estado.configure(text=exc.mensaje)
            return
        self.soltar_resultado()  # antes de unir de nuevo no puede haber dos resultados vivos
        self.version += 1
        version = self.version
        self.uniendo = True
        self.etiqueta_estado.configure(text="")
        self.boton_unir.state(["disabled"])

        def trabajo(progress, cancel):
            return run_join(plan, tablas, progress=progress, cancel=cancel)

        def listo(res: JoinResult) -> None:
            self.uniendo = False
            if version != self.version or not self.winfo_exists():
                return
            self.boton_unir.state(["!disabled"])
            self.resultado = res
            self._mostrar_resultado(res)

        def fallo(exc: BaseException) -> None:
            self.uniendo = False
            if not self.winfo_exists():
                return
            self.boton_unir.state(["!disabled"])
            self.etiqueta_estado.configure(text=self._mensaje(exc))

        def cancelado() -> None:
            self.uniendo = False
            if self.winfo_exists():
                self.boton_unir.state(["!disabled"])
                self.etiqueta_estado.configure(text="Se canceló. No se unió nada.")

        run_in_background(
            self,
            trabajo,
            title="Uniendo los archivos",
            on_success=listo,
            on_error=fallo,
            on_cancel=cancelado,
        )

    def _hay_extras(self, res: JoinResult) -> bool:
        return self.estado.resumen_final is not None or any(
            len(df) for df in res.unmatched.values()
        )

    def _mostrar_resultado(self, res: JoinResult) -> None:
        df = res.df
        self.etiqueta_filas.configure(
            text=(
                f"El resultado tiene {formatear_numero(len(df))} filas y "
                f"{formatear_numero(len(df.columns))} columnas"
            )
        )
        self.etiqueta_avisos_resultado.configure(text="\n".join(f"⚠️ {a}" for a in res.warnings))
        tam = self.grilla.page_size
        self.grilla.set_proveedor(lambda p: (df.iloc[p * tam : (p + 1) * tam], len(df)))
        if self._hay_extras(res):
            self.boton_libro.pack(side="left", padx=(8, 0))
        else:
            self.boton_libro.pack_forget()
        self.marco_resultado.pack(
            fill="both", expand=True, pady=(4, 0), before=self.etiqueta_guardado
        )

    def ver_informacion(self) -> None:
        if self.resultado is None:
            return
        self.explorador = abrir_explorador(
            self.winfo_toplevel(), tabla=self.resultado.df, titulo=TEXTO_TITULO_RESULTADO
        )

    # ---- guardar el resultado (una tabla) ----
    def _nombre_sugerido(self) -> str:
        principal = self.estado.principal
        base = Path(principal.nombre).stem if principal is not None else "resultado"
        return f"{base}_unido"

    def guardar_resultado(self) -> None:
        res = self.resultado
        if res is None:
            return
        self.ventana_guardar = guardar.abrir_guardar(
            self.winfo_toplevel(),
            res.df,
            nombre_sugerido=self._nombre_sugerido(),
            on_terminado=self._guardado_ok,
            claves_sugeridas=list(self.estado.claves),
            column_types=dict(res.column_types),
            salida_inicial=self.estado.salida,
        )
        ventana = self.ventana_guardar
        if ventana is not None:
            ventana.win.bind("<Destroy>", lambda e, v=ventana: self._guardar_cerrado(e, v), add="+")

    def _guardar_cerrado(self, evento: tk.Event, ventana: Any) -> None:
        """Se cerró «Guardar»: se suelta la ventana y su tabla. El resultado se conserva para
        poder guardarlo de nuevo (por ejemplo para otro destino) sin volver a unir; se suelta
        al volver a un paso anterior o antes de unir otra vez."""
        if evento.widget is not ventana.win or ventana is not self.ventana_guardar:
            return
        self.ventana_guardar = None
        ventana.cerrar()
        gc.collect()

    def _guardado_ok(self, resultado: Any) -> None:
        ventana = self.ventana_guardar
        spec = getattr(ventana, "salida_usada", None)
        if spec is not None:
            # La receta recuerda cómo se guardó (destino, formato, división), no los tipos.
            self.estado.salida = dataclasses.replace(spec, column_types=None)
        self.etiqueta_guardado.configure(text="✅ El resultado quedó guardado.")
        if self.asistente is not None:
            self.asistente.marcar_guardado(True)

    # ---- guardar en un libro con varias hojas ----
    def _hojas(self, res: JoinResult, progress, cancel) -> list[SheetSpec]:
        hojas: list[SheetSpec] = []
        if self.estado.resumen_final is not None:
            sin_resumen = dataclasses.replace(self.estado.plan(), final_summary=None)
            detalle = run_join(sin_resumen, self.estado.tablas(), progress=progress, cancel=cancel)
            hojas.append(SheetSpec("Detalle", detalle.df, dict(detalle.column_types)))
            hojas.append(SheetSpec("Resumen", res.df, dict(res.column_types)))
        else:
            hojas.append(SheetSpec("Resultado", res.df, dict(res.column_types)))
        for nombre, df in res.unmatched.items():
            if len(df):
                hojas.append(SheetSpec(f"Sin coincidencia en {Path(nombre).stem}", df))
        return hojas

    def guardar_libro(self) -> None:
        res = self.resultado
        if res is None:
            return

        def trabajo(progress, cancel):
            hojas = self._hojas(res, progress, cancel)
            return hojas, plan_workbook(hojas, "excel")

        run_in_background(
            self,
            trabajo,
            title="Preparando las hojas",
            on_success=self._libro_listo,
            on_error=lambda exc: messagebox.showerror(
                "Guardar en un Excel con varias hojas",
                self._mensaje(exc),
                parent=self.winfo_toplevel(),
            ),
        )

    def _libro_listo(self, listo: tuple[list[SheetSpec], Any]) -> None:
        if not self.winfo_exists():
            return
        hojas, chequeo = listo
        padre = self.winfo_toplevel()
        titulo = "Guardar en un Excel con varias hojas"
        if chequeo.level == "obligatorio":
            motivos = "\n".join(f"• {m}" for m in chequeo.reasons)
            messagebox.showwarning(titulo, f"{TEXTO_LIBRO_GRANDE}\n\n{motivos}", parent=padre)
            return
        if chequeo.level == "recomendado" and not messagebox.askyesno(
            titulo,
            "El libro va a quedar pesado para Excel. ¿Querés guardarlo igual?",
            parent=padre,
        ):
            return
        opciones: dict[str, Any] = {
            "parent": padre,
            "title": titulo,
            "initialfile": f"{self._nombre_sugerido()}.xlsx",
            "defaultextension": ".xlsx",
            "filetypes": [("Excel", "*.xlsx")],
        }
        carpeta = load_preferences().last_folder
        if carpeta is not None and Path(carpeta).is_dir():
            opciones["initialdir"] = str(carpeta)
        elegido = filedialog.asksaveasfilename(**opciones)
        if not elegido:
            return
        ruta = Path(elegido)
        if ruta.suffix.lower() != ".xlsx":
            ruta = ruta.with_name(ruta.name + ".xlsx")
        save_preferences(dataclasses.replace(load_preferences(), last_folder=ruta.parent))

        def trabajo(progress, cancel):
            return save_workbook(hojas, ruta, destination="excel", progress=progress, cancel=cancel)

        def guardado(res: Any) -> None:
            nombres = ", ".join(f"«{p.sheet}»" for p in res.parts)
            self.etiqueta_guardado.configure(
                text=f"✅ Se guardó {ruta.name} con las hojas {nombres}."
            )
            if self.asistente is not None:
                self.asistente.marcar_guardado(True)

        run_in_background(
            self,
            trabajo,
            title="Guardando el libro",
            on_success=guardado,
            on_error=lambda exc: messagebox.showerror(titulo, self._mensaje(exc), parent=padre),
        )

    # ---- receta ----
    def guardar_receta(self) -> None:
        try:
            plan = self.estado.plan()
        except ExcelToolError as exc:
            messagebox.showerror("Guardar la receta", exc.mensaje, parent=self.winfo_toplevel())
            return
        opciones: dict[str, Any] = {
            "parent": self.winfo_toplevel(),
            "title": "Guardar esta receta",
            "initialfile": f"receta_{self._nombre_sugerido()}.json",
            "defaultextension": ".json",
            "filetypes": [("Recetas", "*.json")],
        }
        carpeta = load_preferences().last_folder
        if carpeta is not None and Path(carpeta).is_dir():
            opciones["initialdir"] = str(carpeta)
        elegido = filedialog.asksaveasfilename(**opciones)
        if not elegido:
            return
        ruta = Path(elegido)
        try:
            save_recipe(plan, ruta)
        except Exception as exc:  # noqa: BLE001 - se explica al usuario; el detalle va al log
            messagebox.showerror(
                "Guardar la receta", self._mensaje(exc), parent=self.winfo_toplevel()
            )
            return
        self.etiqueta_guardado.configure(text=f"✅ La receta quedó guardada en {ruta.name}.")

    def _al_destruir(self, evento: tk.Event) -> None:
        if evento.widget is self:
            self.version += 1
