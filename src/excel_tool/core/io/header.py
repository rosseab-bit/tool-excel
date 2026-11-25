"""Detección de la fila de encabezado y normalización de nombres de columna.

Lo comparten CSV y Excel. Trabaja con listas de texto, sin saber de dónde vienen.
"""

from __future__ import annotations

MAX_FILAS_ENCABEZADO = 30
_FILAS_DE_DATOS_A_MIRAR = 5


def _llenas(fila: list[str]) -> int:
    return sum(1 for celda in fila if celda.strip() != "")


def detect_header_row(rows: list[list[str]]) -> int:
    """Índice (desde 0) de la fila de encabezado, mirando las primeras 30 filas.

    Es la primera fila "completa" (con al menos la mitad de las celdas llenas respecto del
    ancho típico de los datos) seguida de filas de datos consistentes. Si no hay nada que
    parezca un encabezado, devuelve 0.
    """
    muestra = rows[:MAX_FILAS_ENCABEZADO]
    cuentas = [_llenas(f) for f in muestra]
    con_datos = [c for c in cuentas if c > 0]
    if not con_datos:
        return 0
    # Ancho típico: el más frecuente (si empatan, el más ancho).
    tipico = max(set(con_datos), key=lambda c: (con_datos.count(c), c))
    if tipico < 2:
        return 0
    for i, cuenta in enumerate(cuentas):
        if cuenta * 2 < tipico:
            continue
        siguientes = cuentas[i + 1 : i + 1 + _FILAS_DE_DATOS_A_MIRAR]
        if not siguientes:
            continue
        consistentes = sum(1 for c in siguientes if c * 2 >= tipico)
        if consistentes * 2 >= len(siguientes):
            return i
    return 0


def column_letter(index: int) -> str:
    """Letra de columna estilo Excel: 0 -> "A", 25 -> "Z", 26 -> "AA"."""
    letras = ""
    n = index + 1
    while n > 0:
        n, resto = divmod(n - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def _lista(nombres: list[str]) -> str:
    if len(nombres) <= 1:
        return "".join(nombres)
    return ", ".join(nombres[:-1]) + " y " + nombres[-1]


def normalize_column_names(raw: list[str]) -> tuple[list[str], dict[str, str], list[str]]:
    """Devuelve (nombres finales, renamed_columns, advertencias).

    Recorta espacios; vacío -> "Columna <letra>"; repetido -> "<nombre> (2)", "(3)"...
    `renamed_columns` mapea nombre final -> original (para las renombradas).
    """
    recortados = [(c or "").strip() for c in raw]
    usados: set[str] = {n for n in recortados if n}
    finales: list[str] = []
    renamed: dict[str, str] = {}
    sin_nombre: list[str] = []
    repetidos: list[str] = []
    vistos: dict[str, int] = {}

    for pos, nombre in enumerate(recortados):
        if nombre == "":
            nuevo = f"Columna {column_letter(pos)}"
            while nuevo in usados:
                nuevo += " (2)"
            usados.add(nuevo)
            finales.append(nuevo)
            renamed[nuevo] = raw[pos] if raw[pos] is not None else ""
            sin_nombre.append(nuevo)
        elif nombre in vistos:
            n = vistos[nombre]
            while True:
                n += 1
                nuevo = f"{nombre} ({n})"
                if nuevo not in usados:
                    break
            vistos[nombre] = n
            usados.add(nuevo)
            finales.append(nuevo)
            renamed[nuevo] = raw[pos]
            repetidos.append(nombre)
        else:
            vistos[nombre] = 1
            finales.append(nombre)

    avisos: list[str] = []
    if sin_nombre:
        if len(sin_nombre) == 1:
            avisos.append(f"Hay 1 columna sin nombre: se llamó {sin_nombre[0]}.")
        else:
            avisos.append(
                f"Hay {len(sin_nombre)} columnas sin nombre: se llamaron {_lista(sin_nombre)}."
            )
    if repetidos:
        unicos = list(dict.fromkeys(repetidos))
        avisos.append(
            f"Hay {len(repetidos) + len(unicos)} columnas con el mismo nombre "
            f"({_lista(unicos)}): a las repetidas se les agregó un número."
        )
    return finales, renamed, avisos
