"""Conductor del asistente para unir archivos (no es un test): maneja la UI como un usuario.

Parte desde la pantalla de inicio ("Unir archivos"), hace clic en los botones reales y espera al
worker. Solo se reemplazan el selector de archivos (devuelve archivos del escenario), los diálogos
del sistema (en `tests/integracion/conftest.py`) y `platform_utils.open_folder`.
"""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from pathlib import Path

from excel_tool import platform_utils
from excel_tool.core.io import CsvOptions
from excel_tool.ui import guardar
from excel_tool.ui.app import PantallaInicio
from excel_tool.ui.selector import Seleccion
from excel_tool.ui.unir import widgets_archivos

CSV_AR = CsvOptions(";", "cp1252")


def textos(widget: tk.Misc) -> list[str]:
    """Todos los textos visibles (etiquetas, botones, valores) de un árbol de widgets."""
    hallados: list[str] = []
    for hijo in widget.winfo_children():
        try:
            hallados.append(str(hijo.cget("text")))
        except tk.TclError:
            pass
        try:
            hallados.append(" ".join(str(x) for x in hijo.cget("values")))
        except tk.TclError:
            pass
        hallados.extend(textos(hijo))
    return hallados


class Conductor:
    def __init__(self, entorno, monkeypatch, csv: CsvOptions | None = CSV_AR) -> None:
        self.e = entorno
        self.csv = csv
        self.cola: list[Path] = []
        self.carpetas: list[Path] = []
        self.guardados: list = []
        self.selecciones: list[Path] = []
        original = guardar.abrir_guardar

        def espia_guardar(*args, **kwargs):
            v = original(*args, **kwargs)
            self.guardados.append(v)
            return v

        def selector_falso(parent, on_accept, path=None):
            ruta = self.cola.pop(0)
            self.selecciones.append(ruta)
            on_accept(Seleccion(ruta, None, self.csv if ruta.suffix.lower() == ".csv" else None, 0))

        monkeypatch.setattr(guardar, "abrir_guardar", espia_guardar)
        monkeypatch.setattr(widgets_archivos, "abrir_selector", selector_falso)
        monkeypatch.setattr(platform_utils, "open_folder", lambda p: self.carpetas.append(Path(p)))
        self.top = tk.Toplevel(entorno.raiz)
        self.top.title("Inicio (test)")
        self.inicio = PantallaInicio(self.top)
        self.asistente = None

    # ---- inicio ----
    def abrir(self):
        boton = next(b for b in self.inicio.botones if b.cget("text") == "Unir archivos")
        boton.invoke()
        self.asistente = self.inicio.asistente
        assert self.asistente is not None
        return self.asistente

    @property
    def estado(self):
        return self.asistente.estado

    def paso(self, n: int):
        return self.asistente.pasos[n - 1]

    # ---- navegación ----
    def siguiente(self) -> int:
        antes = self.asistente.indice
        self.asistente.boton_siguiente.invoke()
        self.e.raiz.update()
        assert self.asistente.indice == antes + 1, (
            f"no avanzó del paso {antes + 1}: {self.asistente.estado_texto.cget('text')!r}"
        )
        self.verificar_entra_en_768()
        return self.asistente.indice + 1

    def verificar_entra_en_768(self) -> None:
        """La ventana del asistente y su pie (Anterior / Siguiente / Cancelar) entran en 768 px."""
        a = self.asistente
        self.e.raiz.update()
        assert a.winfo_height() <= 768, a.winfo_height()
        assert a.winfo_reqheight() <= 768 or a.winfo_height() <= 768
        for boton in (a.boton_cancelar, a.boton_anterior):
            fondo = boton.winfo_rooty() - a.winfo_rooty() + boton.winfo_height()
            assert 0 < fondo <= a.winfo_height(), (boton.cget("text"), fondo, a.winfo_height())
        if a.boton_siguiente.winfo_manager():
            b = a.boton_siguiente
            assert b.winfo_rooty() - a.winfo_rooty() + b.winfo_height() <= a.winfo_height()

    def anterior(self) -> None:
        antes = self.asistente.indice
        self.asistente.boton_anterior.invoke()
        self.e.raiz.update()
        assert self.asistente.indice == antes - 1

    def sin_avance(self) -> str:
        """Intenta avanzar y devuelve el mensaje del pie (no debe avanzar)."""
        antes = self.asistente.indice
        self.asistente.boton_siguiente.invoke()
        self.e.raiz.update()
        assert self.asistente.indice == antes
        return str(self.asistente.estado_texto.cget("text"))

    # ---- paso 1 ----
    def elegir_principal(self, ruta: Path) -> None:
        self.cola.append(ruta)
        p1 = self.paso(1)
        p1.boton_elegir.invoke()
        self.e.esperar(
            lambda: self.estado.principal is not None and not self.estado.principal.falta, 120
        )

    # ---- paso 2 ----
    def marcar_claves(self, *columnas: str) -> None:
        p2 = self.paso(2)
        for c in columnas:
            if not p2._variables[c].get():
                p2.casillas[c].invoke()
        self.e.raiz.update()

    # ---- paso 3 ----
    def esperar_comparacion(self, segundos: float = 120.0) -> None:
        p3 = self.paso(3)
        self.e.esperar(lambda: not p3.calculando, segundos)

    def agregar(self, ruta: Path, esperar: bool = True) -> None:
        p3 = self.paso(3)
        antes = len(self.estado.secundarios)
        self.cola.append(ruta)
        p3.boton_agregar.invoke()
        self.e.esperar(lambda: len(self.estado.secundarios) == antes + 1, 300)
        if esperar:
            self.esperar_comparacion()

    def indicador(self, nombre: str) -> str:
        return str(self.paso(3).filas[nombre].indicador.cget("text"))

    def elegir_columna_comun(self, nombre: str, clave: str, columna: str) -> None:
        combo = self.paso(3).filas[nombre].combos_clave[clave]
        combo.set(columna)
        combo.event_generate("<<ComboboxSelected>>")
        self.e.raiz.update()
        self.esperar_comparacion()

    def limpieza(self, atributo: str, valor: bool = True) -> None:
        p3 = self.paso(3)
        if bool(p3._variables[atributo].get()) != valor:
            # la casilla es un Checkbutton: se simula el clic cambiando la variable y llamando
            # al comando, que es lo que hace el clic
            p3._variables[atributo].set(valor)
            p3._cambio_limpieza()
        self.esperar_comparacion()

    def solo_en_ambos(self, nombre: str) -> None:
        fila = self.paso(3).filas[nombre]
        fila.combo_union.set("Solo los que están en los dos")
        fila.combo_union.event_generate("<<ComboboxSelected>>")
        self.e.raiz.update()

    # ---- paso 4 ----
    def esperar_paso4(self, segundos: float = 300.0) -> None:
        p4 = self.paso(4)
        self.e.esperar(lambda: not p4.pendiente, segundos)

    def tarjeta(self, nombre: str):
        return next(t for t in self.paso(4).tarjetas if t.nombre == nombre)

    def modo_detalle(self, nombre: str) -> None:
        self.tarjeta(nombre).radio_detalle.invoke()
        self.e.raiz.update()
        self.esperar_paso4()

    def modo_primera(self, nombre: str) -> None:
        self.tarjeta(nombre).radio_primera.invoke()
        self.e.raiz.update()
        self.esperar_paso4()

    def modo_resumen(
        self,
        nombre: str,
        *,
        grupos: list[tuple[str, str | None]] = (),
        calculos: list[tuple[str, str] | tuple[str, str, str]] = (),
        quitar_calculos_iniciales: bool = False,
    ):
        """Elige "Ver un resumen" y arma el resumen en el editor (como el usuario): `grupos` son
        columnas extra para agrupar (con el período),
        `calculos` son (operación, columna[, nombre])."""
        p4 = self.paso(4)
        self.tarjeta(nombre).radio_resumen.invoke()
        self.e.esperar(lambda: p4.editor is not None, 30)
        editor = p4.editor
        for col, periodo in grupos:
            editor.checks[col].invoke()
            if periodo:
                editor.var_periodo[col].set(periodo)
                editor.combos_periodo[col].event_generate("<<ComboboxSelected>>")
        for c in calculos:
            fila = editor.agregar_calculo()
            fila.var_operacion.set(c[0])
            fila.combo_operacion.event_generate("<<ComboboxSelected>>")
            if c[1]:
                fila.var_columna.set(c[1])
                fila.combo_columna.event_generate("<<ComboboxSelected>>")
            if len(c) > 2:
                fila.var_nombre.set(c[2])
        self.e.raiz.update()
        p4.boton_listo.invoke()
        self.e.raiz.update()
        self.esperar_paso4()
        return editor

    # ---- paso 5 ----
    def esperar_tamano(self, segundos: float = 300.0) -> None:
        p5 = self.paso(5)
        self.e.esperar(
            lambda: p5.version_mostrada == p5.version and p5._id_espera is None, segundos
        )

    def seccion(self, nombre: str):
        return next(s for s in self.paso(5).secciones if s.nombre == nombre)

    # ---- paso 7 ----
    def esperar_revision(self, segundos: float = 600.0) -> None:
        p7 = self.paso(7)
        self.e.esperar(lambda: not p7.pendiente, segundos)

    def unir(self, segundos: float = 900.0):
        p7 = self.paso(7)
        self.esperar_revision(segundos)
        p7.boton_unir.invoke()
        self.e.esperar(lambda: p7.resultado is not None or bool(self.e.carteles), segundos)
        assert self.e.carteles == [], self.e.carteles
        return p7.resultado

    def guardar_una_tabla(
        self,
        destino: Path,
        *,
        configurar: Callable | None = None,
        segundos: float = 900.0,
    ):
        """ "Guardar el resultado" -> ventana Guardar -> Guardar… -> espera."""
        p7 = self.paso(7)
        previos = len(self.guardados)
        p7.boton_guardar.invoke()
        self.e.esperar(lambda: len(self.guardados) == previos + 1, 30)
        v = self.guardados[-1]
        self.e.destino = destino
        self.e.esperar(lambda: v.puede_guardar() or v._error_plan is not None, 300)
        if configurar is not None:
            configurar(v)
            self.e.esperar(lambda: v.puede_guardar(), 300)
        v._guardar()
        self.e.esperar(lambda: v._resultado is not None or bool(self.e.carteles), segundos)
        assert self.e.carteles == [], self.e.carteles
        return v
