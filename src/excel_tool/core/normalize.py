"""Normalización de claves para unir archivos. Contrato: docs/contratos.md § Fase 4 (ADR-034).

Todo es vectorizado con pyarrow (sin bucles por fila).

Reglas que tiene que saber quien use esto:
- Una clave vacía queda vacía (`""`). Que las claves vacías NO emparejen entre sí lo decide el
  join (F4-03), no esta función.
- Los ceros a la izquierda NO se sacan por defecto (`strip_leading_zeros=False`).
- "123.0" NO se convierte en "123": se compara como texto. Es una decisión consciente: tocarlo
  podría juntar claves que son distintas ("1.0" vs "1" en un código de producto). La lectura de
  Excel ya entrega los números enteros sin ".0".
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass

import pandas as pd
import pyarrow as pa
import pyarrow.compute as pc

_STR = pd.StringDtype("pyarrow")

# Espacios, tabs, saltos, espacio duro (\xa0), espacios Unicode finos/angostos y BOM.
_ESPACIOS = r"[ \t\r\n\f\v\x{00A0}\x{1680}\x{2000}-\x{200B}\x{202F}\x{205F}\x{3000}\x{FEFF}]+"


@dataclass(frozen=True)
class KeyOptions:
    trim: bool = True  # recortar espacios (también dobles espacios internos → uno)
    ignore_case: bool = True
    strip_leading_zeros: bool = False  # "00012" ↔ "12"; solo afecta claves que son todo dígitos


def _a_arrow(s: pd.Series) -> pa.ChunkedArray:
    arr = s.array
    pa_arr = getattr(arr, "_pa_array", None)
    if isinstance(pa_arr, pa.ChunkedArray) and pa.types.is_string(pa_arr.type):
        return pa_arr
    if isinstance(pa_arr, pa.ChunkedArray) and pa.types.is_large_string(pa_arr.type):
        return pa_arr.cast(pa.string())
    return pa.chunked_array([pa.array(s.tolist(), type=pa.string(), from_pandas=True)])


def normalize_keys(s: pd.Series, options: KeyOptions) -> pd.Series:
    """Devuelve la clave normalizada, con el mismo índice y almacenamiento pyarrow.

    Los nulos pasan a `""`. Con `strip_leading_zeros`, solo se tocan los valores que son todo
    dígitos (después del recorte): "00012" → "12", "000" → "0", "00A1" queda igual.
    """
    datos = _a_arrow(s)
    datos = pc.fill_null(datos, "")
    if options.trim:
        datos = pc.replace_substring_regex(datos, _ESPACIOS, " ")
        datos = pc.utf8_trim(datos, " ")
    if options.ignore_case:
        datos = pc.utf8_lower(datos)
    if options.strip_leading_zeros:
        datos = pc.replace_substring_regex(datos, r"^0+([0-9]+)$", r"\1")
    return pd.Series(pd.arrays.ArrowStringArray(datos), index=s.index, name=s.name, dtype=_STR)


def _nombre_clave(nombre: str) -> str:
    sin_tildes = "".join(
        c for c in unicodedata.normalize("NFKD", str(nombre)) if not unicodedata.combining(c)
    )
    return " ".join(sin_tildes.casefold().split())


def match_key_columns(principal: list[str], other: list[str]) -> list[tuple[str, str]]:
    """Empareja columnas por nombre igual sin distinguir mayúsculas, espacios ni tildes.

    Devuelve pares (columna del principal, columna del otro) en el orden del principal. Si un
    nombre normalizado aparece más de una vez de cualquiera de los dos lados, no adivina: no
    devuelve ningún par para ese nombre.
    """
    cuenta_p: dict[str, list[str]] = {}
    for c in principal:
        cuenta_p.setdefault(_nombre_clave(c), []).append(c)
    cuenta_o: dict[str, list[str]] = {}
    for c in other:
        cuenta_o.setdefault(_nombre_clave(c), []).append(c)
    pares = []
    for clave, cols in cuenta_p.items():
        otros = cuenta_o.get(clave, [])
        if len(cols) == 1 and len(otros) == 1:
            pares.append((cols[0], otros[0]))
    return pares
