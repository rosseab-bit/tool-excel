"""Formato para mostrar el perfil de un archivo (números argentinos, tipos en palabras).

No calcula nada: solo da formato a lo que trae el perfil del core.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from excel_tool.ui.grilla import formatear_numero
from excel_tool.ui.selector import CODIFICACIONES, SEPARADORES

NOMBRES_TIPO = {
    "texto": "Texto",
    "entero": "Número entero",
    "decimal": "Número decimal",
    "fecha": "Fecha",
    "fecha_hora": "Fecha y hora",
    "si_no": "Sí/No",
    "mixto": "Mixto",
    "vacia": "Vacía",
}

NOMBRES_FORMATO = {
    "xlsx": "Excel (.xlsx)",
    "xlsm": "Excel con macros (.xlsm)",
    "xlsb": "Excel binario (.xlsb)",
    "xls": "Excel antiguo (.xls)",
    "ods": "Hoja de cálculo (.ods)",
    "csv": "Texto con columnas (.csv)",
}

MARCA_CLAVE = "Posible clave"
MARCA_POCOS = "Pocos valores"
MARCA_SUMAR = "Se puede sumar"
MAX_AVISOS = 200


def formatear_decimal(valor: float, decimales: int = 2) -> str:
    """1234.5 -> '1.234,50'."""
    texto = f"{valor:,.{decimales}f}"
    return texto.replace(",", "\0").replace(".", ",").replace("\0", ".")


def formatear_valor(valor: float) -> str:
    """Entero sin decimales; con decimales, dos."""
    if valor == int(valor) and abs(valor) < 1e15:
        return formatear_numero(int(valor))
    return formatear_decimal(valor)


def formatear_porcentaje(porcentaje: float) -> str:
    """99.5 -> '99,5 %'."""
    return f"{formatear_decimal(porcentaje, 1)} %"


def formatear_tamano(tamano: int) -> str:
    """12900000 -> '12,3 MB'."""
    if tamano < 1024:
        return f"{formatear_numero(tamano)} bytes"
    valor = float(tamano)
    for unidad in ("KB", "MB", "GB", "TB"):
        valor /= 1024
        if valor < 1024 or unidad == "TB":
            return f"{formatear_decimal(valor, 1)} {unidad}"
    return ""  # pragma: no cover


def formatear_fecha(fecha: datetime) -> str:
    return fecha.strftime("%d/%m/%Y %H:%M")


def nombre_tipo(tipo: str) -> str:
    return NOMBRES_TIPO.get(tipo, tipo)


def nombre_formato(formato: str) -> str:
    return NOMBRES_FORMATO.get(formato, formato)


def marcas_columna(col: Any) -> list[str]:
    marcas = []
    if col.is_unique:
        marcas.append(MARCA_CLAVE)
    if col.few_values:
        marcas.append(MARCA_POCOS)
    if col.summable:
        marcas.append(MARCA_SUMAR)
    return marcas


def fila_columna(col: Any) -> tuple[str, ...]:
    """Valores de una fila de la tabla Columnas."""
    return (
        f"{col.letter} ({col.position + 1})",
        col.name,
        nombre_tipo(col.detected_type),
        formatear_numero(col.non_empty),
        formatear_numero(col.empty),
        formatear_numero(col.distinct),
        " · ".join(marcas_columna(col)),
    )


def datos_generales(perfil: Any) -> list[tuple[str, list[tuple[str, str]]]]:
    """Secciones de la pestaña General: [(título, [(etiqueta, valor)])]."""
    archivo = perfil.file
    secciones = [
        (
            "El archivo",
            [
                ("Nombre", archivo.name),
                ("Carpeta", str(archivo.path.parent)),
                ("Tamaño", formatear_tamano(archivo.size)),
                ("Última modificación", formatear_fecha(archivo.modified)),
                ("Formato", nombre_formato(archivo.format)),
            ],
        )
    ]
    hoja: list[tuple[str, str]] = []
    if perfil.sheet:
        hoja.append(("Hoja", perfil.sheet))
    hoja += [
        ("Filas", formatear_numero(perfil.rows)),
        ("Columnas", formatear_numero(perfil.columns)),
        ("Fila de los títulos", formatear_numero(perfil.header_row + 1)),
        ("Filas vacías", formatear_numero(perfil.empty_rows)),
        ("Filas repetidas", formatear_numero(perfil.duplicate_rows)),
    ]
    secciones.append(("La hoja" if perfil.sheet else "Los datos", hoja))
    if perfil.csv_options is not None:
        csv = perfil.csv_options
        secciones.append(
            (
                "El texto",
                [
                    ("Separador de columnas", SEPARADORES.get(csv.separator, csv.separator)),
                    ("Tipo de texto", CODIFICACIONES.get(csv.encoding, csv.encoding)),
                ],
            )
        )
    return secciones


def texto_hoja(hoja: Any) -> str:
    filas = formatear_numero(hoja.rows)
    columnas = formatear_numero(hoja.columns)
    return f"{hoja.name} ({filas} filas, {columnas} columnas)"


def detalle_columna(col: Any) -> list[tuple[str, str]]:
    """Líneas (etiqueta, valor) del panel de detalle. Solo se listan las que aplican."""
    lineas: list[tuple[str, str]] = []
    total = sum(col.type_breakdown.values())
    if col.detected_type == "mixto" and total:
        partes = [
            f"{nombre_tipo(tipo)}: {formatear_porcentaje(cant * 100 / total)} "
            f"({formatear_numero(cant)})"
            for tipo, cant in sorted(col.type_breakdown.items(), key=lambda p: -p[1])
        ]
        lineas.append(("Mezcla de tipos", " · ".join(partes)))
    if col.min is not None:
        lineas.append(("Mínimo", str(col.min)))
    if col.max is not None:
        lineas.append(("Máximo", str(col.max)))
    if col.sum is not None:
        lineas.append(("Suma", formatear_valor(col.sum)))
    if col.mean is not None:
        lineas.append(("Promedio", formatear_valor(col.mean)))
    if col.min_len is not None and col.max_len is not None:
        lineas.append(
            (
                "Largo del texto",
                f"de {formatear_numero(col.min_len)} a {formatear_numero(col.max_len)} caracteres",
            )
        )
    if col.leading_zeros:
        lineas.append(
            (
                "Ceros a la izquierda",
                f"{formatear_numero(col.leading_zeros)} valores (se conservan)",
            )
        )
    if col.untrimmed:
        lineas.append(
            ("Espacios de más", f"{formatear_numero(col.untrimmed)} valores con espacios al borde")
        )
    if col.case_variants:
        lineas.append(
            (
                "Escritos distinto",
                f"{formatear_numero(col.case_variants)} valores que solo cambian en "
                "mayúsculas o espacios (por ejemplo Norte y NORTE)",
            )
        )
    if col.not_converted:
        lineas.append(
            (
                "No son del tipo detectado",
                f"{formatear_numero(col.not_converted)} celdas",
            )
        )
    if col.sample_values:
        lineas.append(("Ejemplos", ", ".join(col.sample_values)))
    return lineas
