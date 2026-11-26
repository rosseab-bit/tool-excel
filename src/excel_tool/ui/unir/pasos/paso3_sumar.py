"""Paso 3: archivos a sumar, con la limpieza de la clave (ADR-038).

Por cada archivo: columna en común (emparejada sola o elegida a mano), qué filas quedan y un
indicador ✅/⚠️ con las coincidencias. Los números salen de `diagnose_join` (core) y se calculan
en segundo plano cada vez que cambia algo.
"""

from __future__ import annotations

import dataclasses
import logging
import tkinter as tk
from dataclasses import dataclass, field
from tkinter import ttk

from excel_tool.core.join import JoinDiagnosis, JoinPlan
from excel_tool.core.normalize import KeyOptions, match_key_columns
from excel_tool.ui.unir.estado import Fuente, FuenteSecundaria
from excel_tool.ui.unir.pasos.base import Paso
from excel_tool.ui.unir.widgets_archivos import (
    pedir_fuente,
    resumen_fuente,
    texto_coincidencia,
    ver_informacion,
)
from excel_tool.ui.worker import run_in_background

log = logging.getLogger(__name__)

TEXTO_AYUDA = (
    "Sumá los archivos de los que querés traer datos. Cada uno se une al principal por la "
    "columna en común."
)
UNIONES = {
    "left": "Completar (todos los del principal)",
    "inner": "Solo los que están en los dos",
}
TEXTO_COMPARANDO = "Comparando los archivos…"
TEXTO_SIN_COMPARAR = "No se pudo comparar este archivo con el principal."
TEXTO_SIN_CLAVE = "Primero marcá la columna clave en el paso anterior."
TEXTO_ELEGI_COLUMNA = "Elegí la columna en común para poder comparar."
TEXTO_FALTA = "Falta el archivo"
MENSAJE_SIN_ARCHIVOS = "Sumá al menos un archivo para seguir."
LIMPIEZA = (
    ("ignore_case", "No distinguir mayúsculas de minúsculas"),
    ("trim", "Recortar espacios al principio y al final"),
    ("strip_leading_zeros", "Sacar ceros a la izquierda (00012 = 12)"),
)


@dataclass
class FilaArchivo:
    """Los widgets de un archivo de la lista (para actualizarlos y para los tests)."""

    marco: ttk.Frame
    indicador: ttk.Label | None = None
    combos_clave: dict[str, ttk.Combobox] = field(default_factory=dict)
    combo_union: ttk.Combobox | None = None
    boton_subir: ttk.Button | None = None
    boton_bajar: ttk.Button | None = None
    boton_quitar: ttk.Button | None = None
    boton_info: ttk.Button | None = None
    boton_elegir: ttk.Button | None = None


class PasoSumar(Paso):
    titulo = "Archivos a sumar"

    def __init__(self, parent, estado, asistente) -> None:
        super().__init__(parent, estado, asistente)
        self.filas: dict[str, FilaArchivo] = {}
        self.fila_principal: FilaArchivo | None = None
        self.diagnostico: dict[str, object] = {}  # nombre -> SecondaryDiagnosis
        self.calculando = False
        self.error_comparacion = False
        self._version = 0
        self._token = None
        self._variables: dict[str, tk.BooleanVar] = {}

        ttk.Label(self, text="¿Qué archivos querés sumar?", font="TkHeadingFont").pack(
            anchor="w", pady=(0, 6)
        )
        ttk.Label(self, text=TEXTO_AYUDA, wraplength=760, justify="left").pack(
            anchor="w", pady=(0, 8)
        )
        self.boton_agregar = ttk.Button(self, text="+ Agregar archivo", command=self._agregar)
        self.boton_agregar.pack(anchor="w")
        self.zona_faltante = ttk.Frame(self)
        self.zona_faltante.pack(fill="x")
        self.lista = ttk.Frame(self)
        self.lista.pack(fill="x", pady=(8, 0))
        self._armar_limpieza()

    # ---- armado ----
    def _armar_limpieza(self) -> None:
        marco = ttk.LabelFrame(self, text="Cómo se comparan las claves (para todos los archivos)")
        marco.pack(fill="x", pady=(12, 0), ipadx=6, ipady=4)
        for atributo, texto in LIMPIEZA:
            var = tk.BooleanVar(master=marco, value=getattr(self.estado.key_options, atributo))
            ttk.Checkbutton(marco, text=texto, variable=var, command=self._cambio_limpieza).pack(
                anchor="w", padx=6
            )
            self._variables[atributo] = var

    def _opciones_desde_pantalla(self) -> KeyOptions:
        return KeyOptions(**{a: bool(v.get()) for a, v in self._variables.items()})

    def _cambio_limpieza(self) -> None:
        self.estado.key_options = self._opciones_desde_pantalla()
        self._recalcular()

    # ---- ciclo del paso ----
    def al_entrar(self) -> None:
        for atributo, var in self._variables.items():
            var.set(getattr(self.estado.key_options, atributo))
        for sec in self.estado.secundarios:
            self._parear(sec)
        self._dibujar()
        self._recalcular()

    def validar(self) -> str | None:
        faltantes = self.estado.faltantes
        if faltantes:
            return f"Falta elegir el archivo «{faltantes[0]}»."
        if not self.estado.secundarios:
            return MENSAJE_SIN_ARCHIVOS
        if not self.estado.claves:
            return TEXTO_SIN_CLAVE
        for sec in self.estado.secundarios:
            if not self._completo(sec):
                return f"Elegí la columna en común de «{sec.nombre}»."
        return None

    # ---- emparejar las claves ----
    def _completo(self, sec: FuenteSecundaria) -> bool:
        return bool(self.estado.claves) and len(sec.key_pairs) == len(self.estado.claves)

    @staticmethod
    def _columnas(sec: Fuente) -> list[str]:
        return [str(c) for c in sec.df.columns] if sec.df is not None else []

    def _automaticas(self, sec: FuenteSecundaria) -> dict[str, str]:
        return dict(match_key_columns(self.estado.claves, self._columnas(sec)))

    def _parear(self, sec: FuenteSecundaria) -> None:
        """Deja en `sec.key_pairs` lo elegido que sigue valiendo y completa con lo emparejado
        solo. Las claves sin pareja quedan sin par hasta que el usuario las elija."""
        if sec.falta:
            return
        columnas = self._columnas(sec)
        auto = self._automaticas(sec)
        actual = dict(sec.key_pairs)
        pares: list[tuple[str, str]] = []
        for clave in self.estado.claves:
            if actual.get(clave) in columnas:
                pares.append((clave, actual[clave]))
            elif clave in auto:
                pares.append((clave, auto[clave]))
        sec.key_pairs = pares

    # ---- dibujo ----
    def _dibujar(self) -> None:
        for zona in (self.zona_faltante, self.lista):
            for hijo in zona.winfo_children():
                hijo.destroy()
        self.filas = {}
        self.fila_principal = None
        principal = self.estado.principal
        if principal is not None and principal.falta:
            self.fila_principal = self._fila_faltante(self.zona_faltante, principal, "Principal: ")
        total = len(self.estado.secundarios)
        for i, sec in enumerate(self.estado.secundarios):
            self.filas[sec.nombre] = self._fila(sec, i, total)
        self._pintar()

    def _marco(self, zona: ttk.Frame) -> ttk.Frame:
        marco = ttk.Frame(zona, padding=8, relief="groove", borderwidth=1)
        marco.pack(fill="x", pady=(0, 6))
        return marco

    def _fila_faltante(self, zona: ttk.Frame, fuente: Fuente, prefijo: str = "") -> FilaArchivo:
        fila = FilaArchivo(self._marco(zona))
        ttk.Label(fila.marco, text=f"{prefijo}{fuente.nombre}", font="TkHeadingFont").pack(
            side="left"
        )
        ttk.Label(fila.marco, text=f" — {TEXTO_FALTA} — ").pack(side="left")
        fila.boton_elegir = ttk.Button(
            fila.marco, text="Elegir…", command=lambda n=fuente.nombre: self._elegir_faltante(n)
        )
        fila.boton_elegir.pack(side="left")
        return fila

    def _fila(self, sec: FuenteSecundaria, i: int, total: int) -> FilaArchivo:
        if sec.falta:
            fila = self._fila_faltante(self.lista, sec)
            self._botones_orden(fila, sec, i, total)
            return fila
        fila = FilaArchivo(self._marco(self.lista))
        cabecera = ttk.Frame(fila.marco)
        cabecera.pack(fill="x")
        ttk.Label(cabecera, text=sec.nombre, font="TkHeadingFont").pack(side="left")
        ttk.Label(cabecera, text=f"  {resumen_fuente(sec)}").pack(side="left")
        self._botones_orden(fila, sec, i, total, cabecera)
        fila.boton_info = ttk.Button(
            cabecera, text="🔍 Ver información", command=lambda s=sec: ver_informacion(self, s)
        )
        fila.boton_info.pack(side="right", padx=(0, 6))

        clave = ttk.Frame(fila.marco)
        clave.pack(fill="x", pady=(6, 0))
        ttk.Label(clave, text="Columna en común:").pack(side="left")
        auto = self._automaticas(sec)
        elegidas = dict(sec.key_pairs)
        for col_principal in self.estado.claves:
            ttk.Label(clave, text=f"  {col_principal} ↔ ").pack(side="left")
            if col_principal in auto and elegidas.get(col_principal) == auto[col_principal]:
                ttk.Label(clave, text=auto[col_principal]).pack(side="left")
                continue
            combo = ttk.Combobox(clave, state="readonly", width=24, values=self._columnas(sec))
            if col_principal in elegidas:
                combo.set(elegidas[col_principal])
            combo.bind(
                "<<ComboboxSelected>>",
                lambda _e, s=sec, c=col_principal, w=combo: self._eligio_columna(s, c, w),
            )
            combo.pack(side="left")
            fila.combos_clave[col_principal] = combo

        union = ttk.Frame(fila.marco)
        union.pack(fill="x", pady=(6, 0))
        ttk.Label(union, text="Qué filas quedan:").pack(side="left")
        fila.combo_union = ttk.Combobox(
            union, state="readonly", width=36, values=list(UNIONES.values())
        )
        fila.combo_union.set(UNIONES[sec.how])
        fila.combo_union.bind(
            "<<ComboboxSelected>>", lambda _e, s=sec, w=fila.combo_union: self._eligio_union(s, w)
        )
        fila.combo_union.pack(side="left", padx=(6, 0))

        fila.indicador = ttk.Label(fila.marco, text="", wraplength=700, justify="left")
        fila.indicador.pack(anchor="w", pady=(6, 0))
        return fila

    def _botones_orden(
        self,
        fila: FilaArchivo,
        sec: FuenteSecundaria,
        i: int,
        total: int,
        padre: ttk.Frame | None = None,
    ) -> None:
        padre = padre or fila.marco
        fila.boton_quitar = ttk.Button(
            padre, text="Quitar", command=lambda n=sec.nombre: self._quitar(n)
        )
        fila.boton_quitar.pack(side="right")
        fila.boton_bajar = ttk.Button(
            padre, text="↓", width=3, command=lambda n=sec.nombre: self._mover(n, 1)
        )
        fila.boton_bajar.pack(side="right", padx=(0, 4))
        fila.boton_subir = ttk.Button(
            padre, text="↑", width=3, command=lambda n=sec.nombre: self._mover(n, -1)
        )
        fila.boton_subir.pack(side="right", padx=(0, 4))
        if i == 0:
            fila.boton_subir.state(["disabled"])
        if i == total - 1:
            fila.boton_bajar.state(["disabled"])

    # ---- acciones sobre la lista ----
    def _nombres_usados(self) -> list[str]:
        return [f.nombre for f in self.estado.fuentes()]

    def _agregar(self) -> None:
        pedir_fuente(self, self._nombres_usados, self._agregado)

    def _agregado(self, fuente: Fuente) -> None:
        if not self.winfo_exists():
            return
        sec = FuenteSecundaria(spec=fuente.spec, df=fuente.df, perfil=fuente.perfil)
        self._parear(sec)
        self.estado.secundarios.append(sec)
        self.asistente.avisar("")
        self._dibujar()
        self._recalcular()

    def _elegir_faltante(self, nombre: str) -> None:
        pedir_fuente(self, self._nombres_usados, lambda f, n=nombre: self._completado(n, f))

    def _completado(self, nombre: str, fuente: Fuente) -> None:
        if not self.winfo_exists():
            return
        self.estado.completar(nombre, fuente)
        for sec in self.estado.secundarios:
            self._parear(sec)
        self.asistente.avisar("")
        self._dibujar()
        self._recalcular()

    def _indice(self, nombre: str) -> int:
        for i, sec in enumerate(self.estado.secundarios):
            if sec.nombre == nombre:
                return i
        raise KeyError(nombre)

    def _mover(self, nombre: str, paso: int) -> None:
        lista = self.estado.secundarios
        i = self._indice(nombre)
        j = i + paso
        if not 0 <= j < len(lista):
            return
        lista[i], lista[j] = lista[j], lista[i]
        self._dibujar()

    def _quitar(self, nombre: str) -> None:
        del self.estado.secundarios[self._indice(nombre)]
        self.estado.columnas.pop(nombre, None)
        self.estado.orden = []
        self.diagnostico.pop(nombre, None)
        self._dibujar()
        self._recalcular()

    def _eligio_columna(self, sec: FuenteSecundaria, clave: str, combo: ttk.Combobox) -> None:
        elegida = combo.get()
        actual = dict(sec.key_pairs)
        actual[clave] = elegida
        sec.key_pairs = [(c, actual[c]) for c in self.estado.claves if c in actual]
        self.asistente.avisar("")
        self._recalcular()

    def _eligio_union(self, sec: FuenteSecundaria, combo: ttk.Combobox) -> None:
        for how, texto in UNIONES.items():
            if texto == combo.get():
                sec.how = how  # type: ignore[assignment]

    # ---- indicadores ----
    def _plan_provisorio(self) -> tuple[JoinPlan, dict] | None:
        """Plan solo para diagnosticar: lo que todavía no se eligió (modo) va como 'primera
        fila' y no se guarda en el estado."""
        estado = self.estado
        principal = estado.principal
        if principal is None or principal.falta or not estado.claves:
            return None
        usables: list[FuenteSecundaria] = []
        for sec in estado.secundarios:
            if sec.falta or not self._completo(sec):
                continue
            modo = sec.mode
            if modo is None or (modo == "summary" and sec.summary is None):
                modo = "first"
            usables.append(dataclasses.replace(sec, mode=modo))
        if not usables:
            return None
        copia = dataclasses.replace(estado, secundarios=usables)
        plan = copia.plan()
        tablas = {f.nombre: f.df for f in [principal, *usables]}
        return plan, tablas

    def _recalcular(self) -> None:
        self._version += 1
        version = self._version
        if self._token is not None:
            self._token.cancel()
            self._token = None
        self.diagnostico = {}
        self.error_comparacion = False
        try:
            armado = self._plan_provisorio()
        except Exception:
            log.exception("No se pudo armar el plan provisorio para comparar")
            armado = None
            self.error_comparacion = True
        if armado is None:
            self.calculando = False
            self._pintar()
            return
        plan, tablas = armado
        self.calculando = True
        self._pintar()
        self._token = run_in_background(
            self,
            lambda progress, cancel: self.estado.diagnosticar(
                plan, tablas, progress=progress, cancel=cancel
            ),
            title="Comparando los archivos",
            on_success=lambda diag: self._llego(version, diag),
            on_error=lambda exc: self._fallo(version, exc),
            on_cancel=lambda: None,
            modal=False,
        )

    def _vivo(self) -> bool:
        try:
            return bool(self.winfo_exists())
        except tk.TclError:
            return False

    def _llego(self, version: int, diag: JoinDiagnosis) -> None:
        if version != self._version or not self._vivo():
            return
        self.calculando = False
        self.diagnostico = {d.name: d for d in diag.secondaries}
        self._pintar()

    def _fallo(self, version: int, exc: BaseException) -> None:
        if version != self._version or not self._vivo():
            return
        log.warning("Falló la comparación de los archivos: %r", exc)
        self.calculando = False
        self.error_comparacion = True
        self._pintar()

    def texto_indicador(self, sec: FuenteSecundaria) -> str:
        if not self.estado.claves:
            return TEXTO_SIN_CLAVE
        if not self._completo(sec):
            return TEXTO_ELEGI_COLUMNA
        diag = self.diagnostico.get(sec.nombre)
        if diag is not None:
            return texto_coincidencia(diag, self.estado.claves)  # type: ignore[arg-type]
        if self.error_comparacion:
            return TEXTO_SIN_COMPARAR
        return TEXTO_COMPARANDO

    def _pintar(self) -> None:
        for sec in self.estado.secundarios:
            fila = self.filas.get(sec.nombre)
            if fila is not None and fila.indicador is not None:
                fila.indicador.configure(text=self.texto_indicador(sec))
