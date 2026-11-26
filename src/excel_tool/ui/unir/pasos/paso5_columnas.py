"""Paso 5: Columnas. Qué columnas de cada archivo van al resultado, en qué orden, y cuánto pesa.

No calcula nada: los nombres finales y el tamaño salen del core (`ui/unir/tamano.py`, en el
worker, con una espera corta para no recalcular en cada clic).
"""

from __future__ import annotations

import logging
import tkinter as tk
from dataclasses import dataclass
from tkinter import ttk

from excel_tool.core.errors import Cancelado, ExcelToolError
from excel_tool.ui.unir.estado import EstadoUnion, Fuente, FuenteSecundaria
from excel_tool.ui.unir.lista_casillas import ListaCasillas
from excel_tool.ui.unir.lista_ordenable import ListaOrdenable
from excel_tool.ui.unir.pasos.base import Paso
from excel_tool.ui.unir.tamano import Calculo, calcular, columnas_de_resumen, nombres_finales
from excel_tool.ui.worker import MENSAJE_INESPERADO, run_in_background

log = logging.getLogger(__name__)

ESPERA_MS = 300
VERDE, AMBAR, ROJO = "#2e7d32", "#b26a00", "#c62828"
COLORES = {"ok": VERDE, "recomendado": AMBAR, "obligatorio": ROJO}
TEXTOS_EXCEL = {
    "ok": "Entra sin problemas en Excel.",
    "recomendado": "En Excel va a andar pesado: conviene dividirlo en varios archivos.",
    "obligatorio": "No entra en una hoja de Excel: hay que dividirlo en varios archivos.",
}
TEXTOS_SHEETS = {
    "ok": "Entra sin problemas en Google Sheets.",
    "recomendado": "En Google Sheets va a andar pesado: conviene dividirlo en varios archivos.",
    "obligatorio": "No entra en Google Sheets: hay que dividirlo en varios archivos.",
}
ROTULOS_MODO = {"detail": "todas las filas", "first": "primera fila", "summary": "resumen"}
SIN_COLUMNAS = "Elegí al menos una columna además de la clave."

Clave = tuple[str, str]  # (nombre del archivo, columna de origen)


def formatear_entero(valor: int) -> str:
    return f"{valor:,}".replace(",", ".")


def texto_celdas(celdas: int) -> str:
    if celdas >= 1_000_000:
        millones = f"{celdas / 1_000_000:.1f}".replace(".", ",").removesuffix(",0")
        return f"{millones} {'millón' if millones == '1' else 'millones'} de celdas"
    return f"{formatear_entero(celdas)} {'celda' if celdas == 1 else 'celdas'}"


def texto_tamano(filas: int, columnas: int, celdas: int) -> str:
    return (
        f"{formatear_entero(filas)} {'fila' if filas == 1 else 'filas'} × "
        f"{formatear_entero(columnas)} {'columna' if columnas == 1 else 'columnas'} = "
        f"{texto_celdas(celdas)}"
    )


@dataclass
class _Seccion:
    nombre: str
    rotulo: str
    disponibles: list[str]
    fijas: frozenset[str]
    lista: ListaCasillas | None  # None = columnas que no se pueden quitar (resumen)

    def marcadas(self) -> list[str]:
        return self.lista.marcadas() if self.lista else list(self.disponibles)


class PasoColumnas(Paso):
    titulo = "Columnas"

    def __init__(self, parent, estado: EstadoUnion, asistente) -> None:
        super().__init__(parent, estado, asistente)
        self.espera_ms = ESPERA_MS
        self.secciones: list[_Seccion] = []
        self.items: list[Clave] = []
        self.nombres: dict[Clave, str] = {}
        self.version = 0  # sube con cada cambio; un resultado viejo se descarta
        self.version_mostrada = 0
        self.calculo: Calculo | None = None
        self._mapeo_listo = False
        self._reordenar_con_mapeo = False
        self._orden_tocado = False
        self._id_espera: str | None = None
        self._token = None

        ttk.Label(self, text=self.titulo, font="TkHeadingFont").pack(anchor="w")
        ttk.Label(
            self,
            text=(
                "Elegí qué columnas de cada archivo querés en el resultado y en qué orden. "
                "La columna en común aparece una sola vez."
            ),
            wraplength=820,
        ).pack(anchor="w", pady=(2, 8))
        cuerpo = ttk.Frame(self)
        cuerpo.pack(fill="both", expand=True)
        cuerpo.columnconfigure(0, weight=3, uniform="col")
        cuerpo.columnconfigure(1, weight=2, uniform="col")
        self.panel = ttk.Frame(cuerpo)
        self.panel.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        derecha = ttk.Frame(cuerpo)
        derecha.grid(row=0, column=1, sticky="nsew")
        self.rotulo_lista = ttk.Label(derecha, text="Orden de las columnas en el resultado")
        self.rotulo_lista.pack(anchor="w")
        self.lista_final = ListaOrdenable(derecha, on_change=self._al_reordenar)
        self.lista_final.pack(fill="both", expand=True, pady=(2, 8))
        marco = ttk.LabelFrame(derecha, text="Tamaño del resultado", padding=8)
        marco.pack(fill="x")
        self.etiqueta_tamano = ttk.Label(marco, text="", wraplength=330, justify="left")
        self.etiqueta_tamano.pack(anchor="w")
        self.etiqueta_excel = ttk.Label(marco, text="", wraplength=330, justify="left")
        self.etiqueta_excel.pack(anchor="w", pady=(4, 0))
        self.etiqueta_sheets = ttk.Label(marco, text="", wraplength=330, justify="left")
        self.etiqueta_sheets.pack(anchor="w")
        self.bind("<Destroy>", self._al_destruir, add="+")

    # ---- ciclo del paso ----
    def al_entrar(self) -> None:
        self._armar_secciones()
        self.items = self._items_marcados()
        self._mapeo_listo = False
        self._orden_tocado = False
        self._reordenar_con_mapeo = bool(self.estado.orden)
        if self.nombres:
            self._aplicar_orden_guardado()
        self._mostrar_lista()
        self._guardar()  # el plan para calcular el tamaño parte de lo que se ve
        self._programar(inmediato=True)

    def validar(self) -> str | None:
        return None if self._hay_columnas() else SIN_COLUMNAS

    def al_salir(self) -> None:
        self._guardar()
        if self.version_mostrada != self.version:
            # Los nombres que se ven pueden ser de antes del último clic: se renuevan ya
            # (con una muestra chica) para que el orden guardado coincida con el resultado.
            try:
                plan = self.estado.plan()
                self.nombres.update(nombres_finales(plan, self.estado.tablas()))
                self._mapeo_listo = True
            except ExcelToolError:
                return
            self._mostrar_lista(self.lista_final.indice)
            self._guardar()

    # ---- secciones ----
    def _armar_secciones(self) -> None:
        for hijo in self.panel.winfo_children():
            hijo.destroy()
        self.secciones = []
        estado = self.estado
        if estado.principal is not None:
            self._agregar_seccion(estado.principal, None)
        for sec in estado.secundarios:
            self._agregar_seccion(sec, sec)
        if not self.secciones:
            ttk.Label(self.panel, text="Todavía no hay archivos elegidos.").pack(anchor="w")

    def _marcadas_iniciales(self, fuente: Fuente, disponibles: list[str]) -> set[str]:
        if fuente.nombre in self.estado.columnas:
            elegidas = self.estado.columnas[fuente.nombre]
        elif fuente.spec.columns:
            elegidas = fuente.spec.columns
        else:
            return set(disponibles)
        return set(elegidas) & set(disponibles)

    def _agregar_seccion(self, fuente: Fuente, sec: FuenteSecundaria | None) -> None:
        if fuente.df is None:
            return
        if sec is None:
            rotulo = "archivo principal"
            disponibles = [str(c) for c in fuente.df.columns]
            fijas = frozenset(c for c in self.estado.claves if c in disponibles)
        else:
            rotulo = ROTULOS_MODO.get(sec.mode or "detail", "")
            claves = {q for _, q in sec.key_pairs}
            disponibles = [str(c) for c in fuente.df.columns if str(c) not in claves]
            fijas = frozenset()
        marco = ttk.LabelFrame(self.panel, text=f"{fuente.nombre} ({rotulo})", padding=6)
        marco.pack(fill="x", pady=(0, 8))
        if sec is not None and sec.mode == "summary" and sec.summary is not None:
            columnas = columnas_de_resumen(sec.summary, {q for _, q in sec.key_pairs})
            ttk.Label(
                marco,
                text="Columnas del resumen (se cambian en el paso «¿Detalle o resumen?»):",
            ).pack(anchor="w")
            for c in columnas:
                ttk.Label(marco, text=f"•  {c}").pack(anchor="w", padx=(12, 0))
            self.secciones.append(_Seccion(fuente.nombre, rotulo, columnas, frozenset(), None))
            return
        lista = ListaCasillas(
            marco,
            disponibles,
            self._marcadas_iniciales(fuente, disponibles),
            fijas=fijas,
            on_change=self._al_cambiar_casillas,
        )
        lista.pack(fill="x")
        self.secciones.append(_Seccion(fuente.nombre, rotulo, disponibles, fijas, lista))

    def _items_marcados(self) -> list[Clave]:
        return [(s.nombre, c) for s in self.secciones for c in s.marcadas()]

    def _hay_columnas(self) -> bool:
        if not self.secciones:
            return False
        claves = set(self.estado.claves)
        for i, s in enumerate(self.secciones):
            for c in s.marcadas():
                if not (i == 0 and c in claves):
                    return True
        return False

    # ---- lista final ----
    def _texto(self, clave: Clave) -> str:
        return self.nombres.get(clave, clave[1])

    def _mostrar_lista(self, seleccionado: int | None = None) -> None:
        self.lista_final.poner([(k, self._texto(k)) for k in self.items], seleccionado)

    def _aplicar_orden_guardado(self) -> None:
        orden = {n: i for i, n in enumerate(self.estado.orden)}
        pos = {k: i for i, k in enumerate(self.items)}
        self.items.sort(key=lambda k: (orden.get(self._texto(k), len(orden)), pos[k]))

    def _al_cambiar_casillas(self) -> None:
        marcadas = set(self._items_marcados())
        conservadas = [k for k in self.items if k in marcadas]
        nuevas = [k for k in self._items_marcados() if k not in set(conservadas)]
        self.items = conservadas + nuevas
        self._mostrar_lista()
        self._guardar()
        self._programar()

    def _al_reordenar(self) -> None:
        self.items = list(self.lista_final.claves())  # type: ignore[arg-type]
        self._orden_tocado = True
        self._reordenar_con_mapeo = False
        self._guardar()

    def _guardar(self) -> None:
        """Vuelca columnas y orden al estado."""
        for s in self.secciones:
            if s.lista is not None:
                self.estado.columnas[s.nombre] = s.marcadas()
        if self._mapeo_listo or self._orden_tocado:
            self.estado.orden = [self._texto(k) for k in self.items]

    # ---- tamaño ----
    def _programar(self, inmediato: bool = False) -> None:
        self.version += 1
        if self._id_espera is not None:
            try:
                self.after_cancel(self._id_espera)
            except tk.TclError:
                pass
        if self._token is not None:
            self._token.cancel()
            self._token = None
        self._poner_mensaje("Calculando el tamaño del resultado…")
        version = self.version
        self._id_espera = self.after(0 if inmediato else self.espera_ms, self._lanzar, version)

    def _poner_mensaje(self, texto: str) -> None:
        self.etiqueta_tamano.configure(text=texto)
        self.etiqueta_excel.configure(text="")
        self.etiqueta_sheets.configure(text="")

    def _lanzar(self, version: int) -> None:
        self._id_espera = None
        if version != self.version or not self.winfo_exists():
            return
        try:
            plan = self.estado.plan()
            tablas = self.estado.tablas()
        except ExcelToolError as exc:
            self._poner_mensaje(f"Todavía no se puede calcular el tamaño: {exc}")
            self.version_mostrada = version
            return

        def tarea(progress, cancel):
            return calcular(plan, tablas, progress, cancel, self.estado.diagnosticar)

        self._token = run_in_background(
            self,
            tarea,
            title="Calculando el tamaño del resultado",
            on_success=lambda r: self._llego(version, r),
            on_error=lambda e: self._fallo(version, e),
            modal=False,
            on_cancel=lambda: self._poner_mensaje("Se canceló el cálculo del tamaño."),
        )

    def _llego(self, version: int, calculo: Calculo) -> None:
        if version != self.version or not self.winfo_exists():
            return
        self.version_mostrada = version
        self.calculo = calculo
        self.nombres.update(calculo.nombres)
        self._mapeo_listo = True
        if self._reordenar_con_mapeo:
            self._reordenar_con_mapeo = False
            self._aplicar_orden_guardado()
        self._mostrar_lista(self.lista_final.indice)
        self._guardar()
        diag = calculo.diagnostico
        self.etiqueta_tamano.configure(
            text=texto_tamano(diag.rows_result, diag.columns_result, diag.size.cells)
        )
        self._semaforo(self.etiqueta_excel, calculo.excel.level, TEXTOS_EXCEL)
        self._semaforo(self.etiqueta_sheets, calculo.sheets.level, TEXTOS_SHEETS)

    def _semaforo(self, etiqueta: ttk.Label, nivel: str, textos: dict[str, str]) -> None:
        etiqueta.configure(text=f"●  {textos[nivel]}", foreground=COLORES[nivel])

    def _fallo(self, version: int, exc: BaseException) -> None:
        if version != self.version or not self.winfo_exists() or isinstance(exc, Cancelado):
            return
        self.version_mostrada = version
        if isinstance(exc, ExcelToolError):
            log.warning("No se pudo calcular el tamaño: %s (%s)", exc, exc.detalle)
            self._poner_mensaje(str(exc))
        else:
            log.error("Error calculando el tamaño", exc_info=exc)
            self._poner_mensaje(MENSAJE_INESPERADO)

    def _al_destruir(self, evento: tk.Event) -> None:
        if evento.widget is self:
            self.version += 1
            if self._token is not None:
                self._token.cancel()
