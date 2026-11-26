"""Piezas compartidas por los pasos 1 a 3 del asistente: elegir y leer un archivo, abrir su
información y armar los textos del indicador de coincidencias.

No calcula nada: los números de los textos vienen de `diagnose_join` (core).
"""

from __future__ import annotations

import dataclasses
import logging
import tkinter as tk
from collections.abc import Callable, Iterable

from excel_tool.core.join import SecondaryDiagnosis, SourceSpec
from excel_tool.ui.explorador import abrir_explorador
from excel_tool.ui.formato import formatear_porcentaje
from excel_tool.ui.grilla import formatear_numero
from excel_tool.ui.selector import Seleccion, abrir_selector
from excel_tool.ui.unir.estado import Fuente, leer_fuente
from excel_tool.ui.worker import run_in_background

log = logging.getLogger(__name__)

UMBRAL_BIEN = 0.8  # desde este porcentaje de coincidencia se muestra ✅
TEXTO_LEYENDO = "Leyendo el archivo"


def nombre_unico(nombre: str, existentes: Iterable[str]) -> str:
    """'ventas.csv' -> 'ventas.csv (2)' si ya hay uno con ese nombre."""
    usados = set(existentes)
    if nombre not in usados:
        return nombre
    n = 2
    while f"{nombre} ({n})" in usados:
        n += 1
    return f"{nombre} ({n})"


def seleccion_de(spec: SourceSpec) -> Seleccion:
    """La `Seleccion` equivalente a un archivo ya cargado (para abrir el explorador)."""
    opciones = spec.read_options
    return Seleccion(spec.path, opciones.sheet, opciones.csv, opciones.header_row or 0)


def ver_informacion(parent: tk.Misc, fuente: Fuente) -> None:
    """Abre el explorador sobre el archivo. Si ya está leído, usa esa tabla y su perfil (no se
    vuelve a leer el archivo)."""
    seleccion = seleccion_de(fuente.spec)
    if fuente.df is None:
        abrir_explorador(parent.winfo_toplevel(), seleccion)
        return
    abrir_explorador(
        parent.winfo_toplevel(),
        seleccion,
        tabla=fuente.df,
        titulo=fuente.nombre,
        perfil=fuente.perfil,
    )


def pedir_fuente(
    parent: tk.Misc,
    nombres_usados: Callable[[], Iterable[str]],
    al_cargar: Callable[[Fuente], None],
) -> None:
    """Abre el selector y lee el archivo elegido en segundo plano (con progreso y cancelar).

    `al_cargar` recibe la `Fuente` con nombre único respecto de `nombres_usados()` y con todas
    sus columnas como columnas a conservar.
    """

    def aceptar(seleccion: Seleccion) -> None:
        nombre = nombre_unico(seleccion.path.name, nombres_usados())
        spec = SourceSpec(nombre, seleccion.path, seleccion.read_options(), [])
        run_in_background(
            parent,
            lambda progress, cancel: leer_fuente(spec, progress=progress, cancel=cancel),
            title=TEXTO_LEYENDO,
            on_success=lambda fuente: al_cargar(_con_todas_las_columnas(fuente)),
        )

    abrir_selector(parent.winfo_toplevel(), aceptar)


def _con_todas_las_columnas(fuente: Fuente) -> Fuente:
    if fuente.df is not None:
        fuente.spec = dataclasses.replace(fuente.spec, columns=[str(c) for c in fuente.df.columns])
    return fuente


def resumen_fuente(fuente: Fuente) -> str:
    """'30 filas · 4 columnas'."""
    if fuente.df is None:
        return "Falta el archivo"
    filas, columnas = fuente.df.shape
    return (
        f"{formatear_numero(filas)} {'fila' if filas == 1 else 'filas'} · "
        f"{formatear_numero(columnas)} {'columna' if columnas == 1 else 'columnas'}"
    )


def texto_coincidencia(diag: SecondaryDiagnosis, claves: list[str]) -> str:
    """Indicador ✅/⚠️ de un archivo a sumar, con los números del diagnóstico."""
    total = diag.keys_principal
    hallados = diag.keys_matched
    que = f"valores de «{claves[0]}»" if len(claves) == 1 else "combinaciones de las columnas clave"
    if total == 0:
        return "⚠️ El archivo principal no tiene valores en la columna clave."
    proporcion = hallados / total
    porcentaje = formatear_porcentaje(100 * proporcion)
    base = f"{formatear_numero(hallados)} de {formatear_numero(total)} {que}"
    if proporcion >= UMBRAL_BIEN:
        texto = f"✅ {base} encontrados ({porcentaje})"
    elif hallados == 0:
        texto = f"⚠️ No se encontró ninguno de los {formatear_numero(total)} {que}"
    else:
        texto = f"⚠️ Solo {base} encontrados ({porcentaje})"
    con_ceros = diag.keys_matched_without_zeros
    if proporcion < UMBRAL_BIEN and con_ceros is not None and con_ceros > hallados:
        texto += (
            f": probá «Sacar ceros a la izquierda» "
            f"(así se encuentran {formatear_numero(con_ceros)})"
        )
    elif proporcion < UMBRAL_BIEN:
        texto += ": revisá que sea la columna correcta"
    if diag.keys_with_many_rows:
        texto += (
            f". En este archivo, {formatear_numero(diag.keys_with_many_rows)} "
            f"{'se repite' if diag.keys_with_many_rows == 1 else 'se repiten'} en varias filas"
        )
        if diag.example_many is not None:
            valor, veces = diag.example_many
            texto += f" (por ejemplo, «{valor}» aparece {formatear_numero(veces)} veces)"
    return texto
