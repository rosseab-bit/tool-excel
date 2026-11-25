"""Perfil de una tabla ya leída: filas, columnas, tipos, vacíos, distintos y marcas de calidad.

Contrato: docs/contratos.md § Fase 2 (`core/profile.py`) y ADR-027.

Trabaja columna por columna y sobre los valores DISTINTOS de cada una (con su cantidad de
repeticiones): así los cálculos de texto, números y fechas se hacen una vez por valor distinto
y no una vez por celda. No copia la tabla ni pasa columnas enteras a objetos de Python.
Los duplicados se detectan con un hash de fila armado con los códigos de cada columna.
La detección de tipos y las conversiones son de `core/types.py`.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

from excel_tool.core.explore import COL_CANTIDAD, COL_PORCENTAJE, DistinctResult
from excel_tool.core.io import (
    CsvOptions,
    FileFormat,
    ReadResult,
    SheetSummary,
    WriteFormat,
    WriteOptions,
    WriteResult,
    write_sheets,
    write_table,
)
from excel_tool.core.io.writer import nombre_de_hoja
from excel_tool.core.progress import CancelToken, ProgressCallback, ProgressEvent, ThrottledProgress
from excel_tool.core.types import (
    ColumnType,
    ValueKind,
    detect_column_type,
    to_date,
    to_number,
)

MAX_POCOS_VALORES = 50
MAX_TOP = 10
MAX_MUESTRA = 5
MIN_NUMERICO_SUMABLE = 0.9
BLOQUE = 200_000  # valores distintos por bloque al convertir a número o fecha
_MULTIPLICADOR = np.uint64(0x9E3779B97F4A7C15)
_MAX_EJEMPLOS = 3


@dataclass
class FileInfo:
    name: str
    path: Path
    size: int
    modified: datetime
    format: FileFormat


@dataclass
class ColumnProfile:
    name: str
    position: int  # 0 = primera
    letter: str  # "A", "B", ...
    detected_type: ColumnType
    type_breakdown: dict[ValueKind, int]
    non_empty: int
    empty: int
    distinct: int  # valores tal cual (sin normalizar), sin contar los vacíos
    is_unique: bool  # sin vacíos y distinct == filas: "posible clave"
    few_values: bool  # distinct <= 50 y filas >= 2 x distinct
    summable: bool  # entero, decimal o mixto con >= 90 % numérico
    top_values: list[tuple[str, int]]  # los 10 más frecuentes (sin vacíos)
    sample_values: list[str]  # hasta 5 distintos, no vacíos
    min: str | None  # números y fechas, como texto para mostrar
    max: str | None
    sum: float | None  # solo lo convertible
    mean: float | None
    min_len: int | None  # texto
    max_len: int | None
    leading_zeros: int  # celdas que empiezan con 0 seguido de otro dígito ("00123", "03/07/2024")
    untrimmed: int  # celdas con espacios al principio o al final
    case_variants: int  # valores distintos que solo difieren de otro en mayúsculas o espacios
    not_converted: int  # celdas no vacías que no son del tipo detectado (mixto)
    notes: list[str]


@dataclass
class TableProfile:
    file: FileInfo
    sheets: list[SheetSummary]  # [] en CSV
    sheet: str | None
    csv_options: CsvOptions | None
    header_row: int
    rows: int
    columns: int
    empty_rows: int  # filas totalmente vacías
    duplicate_rows: int  # copias de una fila anterior no vacía
    columns_profile: list[ColumnProfile]
    warnings: list[str]


# ---------------------------------------------------------------- formato argentino


def column_letter(position: int) -> str:
    """0 -> "A", 25 -> "Z", 26 -> "AA"."""
    letra = ""
    n = position + 1
    while n > 0:
        n, resto = divmod(n - 1, 26)
        letra = chr(65 + resto) + letra
    return letra


def _numero_a_texto(valor: float) -> str:
    """Coma decimal y sin separador de miles: 1234.5 -> "1234,5"."""
    if float(valor).is_integer() and abs(valor) < 1e15:
        return str(int(valor))
    return np.format_float_positional(valor, trim="-").replace(".", ",")


def _fecha_a_texto(valor: pd.Timestamp, con_hora: bool) -> str:
    # Sin strftime: en Windows no acepta años anteriores a 1900 ni siempre los ceros.
    texto = f"{valor.day:02d}/{valor.month:02d}/{valor.year:04d}"
    if not con_hora:
        return texto
    texto += f" {valor.hour:02d}:{valor.minute:02d}"
    if valor.second:
        texto += f":{valor.second:02d}"
    return texto


# ---------------------------------------------------------------- una columna


def _como_arrow(serie: pd.Series) -> pd.Series:
    if not isinstance(serie.dtype, pd.StringDtype):
        serie = serie.astype(object).where(serie.notna(), "").astype(str)
        return serie.astype(pd.StringDtype("pyarrow"))
    return serie.fillna("")


def _bool(serie: pd.Series) -> np.ndarray:
    return serie.fillna(False).to_numpy(dtype=bool)


def _top(cuentas: np.ndarray, indices: np.ndarray) -> np.ndarray:
    """Índices (de `indices`) de los MAX_TOP más frecuentes; empates por orden de aparición."""
    if len(indices) == 0:
        return indices
    c = cuentas[indices]
    if len(c) <= MAX_TOP:
        orden = np.argsort(-c, kind="stable")
        return indices[orden]
    umbral = np.partition(c, len(c) - MAX_TOP)[len(c) - MAX_TOP]
    mayores = np.flatnonzero(c > umbral)
    iguales = np.flatnonzero(c == umbral)[: MAX_TOP - len(mayores)]
    elegidos = np.concatenate([mayores, iguales])
    orden = np.argsort(-c[elegidos], kind="stable")
    return indices[elegidos[orden]]


def _variantes(originales: pd.Series, limpios: pd.Series) -> tuple[int, list[str]]:
    """(valores que solo difieren de otro en mayúsculas/espacios, ejemplos).

    `originales`: distintos no vacíos tal cual; `limpios`: los mismos sin espacios al borde.
    """
    clave = limpios.str.lower()
    codigos, _ = pd.factorize(clave.array)
    tam = np.bincount(codigos)
    en_grupo = tam[codigos] > 1
    total = int(en_grupo.sum())
    if total == 0:
        return 0, []
    ejemplos: list[str] = []
    for g in np.unique(codigos[en_grupo])[:_MAX_EJEMPLOS]:
        miembros = originales.iloc[np.flatnonzero(codigos == g)[:2]].tolist()
        ejemplos.append(" / ".join(f'"{m}"' for m in miembros))
    return total, ejemplos


def _estadisticas_numericas(
    unicos: pd.Series,
    cuentas: np.ndarray,
    idx_no_vacios: np.ndarray,
    estilo,
    cancel: CancelToken | None,
) -> tuple[float | None, float | None, float | None, int]:
    """(mínimo, máximo, suma, cantidad de celdas convertidas) pesando por repeticiones."""
    minimo = maximo = None
    suma = 0.0
    cant = 0
    for ini in range(0, len(idx_no_vacios), BLOQUE):
        if cancel:
            cancel.check()
        idx = idx_no_vacios[ini : ini + BLOQUE]
        valores = to_number(unicos.iloc[idx].reset_index(drop=True), estilo).values.to_numpy(
            dtype=float
        )
        ok = ~np.isnan(valores)
        if not ok.any():
            continue
        v = valores[ok]
        w = cuentas[idx][ok]
        suma += float(np.dot(v, w))
        cant += int(w.sum())
        minimo = float(v.min()) if minimo is None else min(minimo, float(v.min()))
        maximo = float(v.max()) if maximo is None else max(maximo, float(v.max()))
    return minimo, maximo, suma, cant


def _extremos_fechas(
    unicos: pd.Series, idx_no_vacios: np.ndarray, cancel: CancelToken | None
) -> tuple[pd.Timestamp | None, pd.Timestamp | None]:
    minimo = maximo = None
    for ini in range(0, len(idx_no_vacios), BLOQUE):
        if cancel:
            cancel.check()
        idx = idx_no_vacios[ini : ini + BLOQUE]
        fechas = to_date(unicos.iloc[idx].reset_index(drop=True)).values.dropna()
        if fechas.empty:
            continue
        lo, hi = fechas.min(), fechas.max()
        minimo = lo if minimo is None or lo < minimo else minimo
        maximo = hi if maximo is None or hi > maximo else maximo
    return minimo, maximo


def _no_convertidas(tipo: ColumnType, desglose: dict[str, int], no_vacias: int) -> int:
    """Celdas no vacías que no son del tipo (o grupo de tipos) predominante de una columna mixta."""
    if tipo != "mixto":
        return 0
    grupos = dict(desglose)
    numeros = grupos.pop("entero", 0) + grupos.pop("decimal", 0)
    fechas = grupos.pop("fecha", 0) + grupos.pop("fecha_hora", 0)
    mayor = max([numeros, fechas, *grupos.values()])
    return no_vacias - mayor


def _perfilar_columna(
    nombre: str,
    posicion: int,
    serie: pd.Series,
    filas: int,
    cancel: CancelToken | None,
    hash_filas: np.ndarray,
    fila_vacia: np.ndarray,
) -> ColumnProfile:
    """Perfil de una columna; además acumula el hash de filas y la marca de fila vacía."""
    arreglo = _como_arrow(serie).array
    codigos, unicos_arr = pd.factorize(arreglo)
    unicos = pd.Series(unicos_arr, dtype=pd.StringDtype("pyarrow"))
    cuentas = np.bincount(codigos, minlength=len(unicos))
    limpios = unicos.str.strip()
    vacio_u = _bool(limpios == "")
    if cancel:
        cancel.check()

    # acumulados de tabla (fila vacía y hash para detectar repetidas)
    fila_vacia &= vacio_u[codigos]
    hash_filas *= _MULTIPLICADOR
    hash_filas ^= codigos.astype(np.uint64) + np.uint64(1)
    hash_filas ^= hash_filas >> np.uint64(29)
    del codigos

    idx_no_vacios = np.flatnonzero(~vacio_u)
    vacios = int(cuentas[vacio_u].sum())
    no_vacias = filas - vacios
    distintos = len(idx_no_vacios)

    # marcas de calidad (pesadas por repeticiones)
    sin_recortar = _bool(unicos != limpios) & ~vacio_u
    con_ceros = _bool(limpios.str.match(r"0\d"))
    sin_recortar_n = int(cuentas[sin_recortar].sum())
    ceros_n = int(cuentas[con_ceros].sum())
    variantes, ejemplos = (0, [])
    if distintos > 1:
        variantes, ejemplos = _variantes(unicos.iloc[idx_no_vacios], limpios.iloc[idx_no_vacios])

    top_idx = _top(cuentas, idx_no_vacios)
    top = [(str(unicos.iloc[i]), int(cuentas[i])) for i in top_idx]
    muestra = [str(v) for v in unicos.iloc[idx_no_vacios[:MAX_MUESTRA]]]

    if cancel:
        cancel.check()
    det = detect_column_type(serie, cancel=cancel)
    if cancel:
        cancel.check()
    tipo = det.detected
    desglose = det.breakdown

    # números y fechas
    minimo: str | None = None
    maximo: str | None = None
    suma = promedio = None
    con_numeros = "entero" in desglose or "decimal" in desglose
    numericas = desglose.get("entero", 0) + desglose.get("decimal", 0)
    sumable = tipo in ("entero", "decimal") or (
        tipo == "mixto" and no_vacias > 0 and numericas / no_vacias >= MIN_NUMERICO_SUMABLE
    )
    if con_numeros:
        lo, hi, total, cant = _estadisticas_numericas(
            unicos, cuentas, idx_no_vacios, det.decimal_style, cancel
        )
        if lo is not None and hi is not None:
            minimo, maximo = _numero_a_texto(lo), _numero_a_texto(hi)
        if sumable and cant:
            suma, promedio = total, total / cant
    elif tipo in ("fecha", "fecha_hora"):
        lo_f, hi_f = _extremos_fechas(unicos, idx_no_vacios, cancel)
        if lo_f is not None and hi_f is not None:
            hora = tipo == "fecha_hora"
            minimo, maximo = _fecha_a_texto(lo_f, hora), _fecha_a_texto(hi_f, hora)

    min_len = max_len = None
    if tipo in ("texto", "mixto") and distintos:
        largos = unicos.str.len().to_numpy(dtype=np.int64, na_value=0)[idx_no_vacios]
        min_len, max_len = int(largos.min()), int(largos.max())

    no_convertidas = _no_convertidas(tipo, desglose, no_vacias)

    notas: list[str] = []
    if tipo == "vacia":
        notas.append("La columna está vacía.")
    if not nombre.strip():
        notas.append("La columna no tiene nombre.")
    if no_convertidas:
        notas.append(
            f"Hay {no_convertidas} valores que no son números."
            if numericas >= no_vacias - numericas
            else f"Hay {no_convertidas} valores de otra clase que el resto de la columna."
        )
    if sin_recortar_n:
        notas.append(f"Hay {sin_recortar_n} valores con espacios al principio o al final.")
    if variantes:
        notas.append(
            "Hay valores que solo cambian en mayúsculas o espacios: " + "; ".join(ejemplos) + "."
        )

    return ColumnProfile(
        name=nombre,
        position=posicion,
        letter=column_letter(posicion),
        detected_type=tipo,
        type_breakdown=desglose,
        non_empty=no_vacias,
        empty=vacios,
        distinct=distintos,
        is_unique=vacios == 0 and filas > 0 and distintos == filas,
        few_values=0 < distintos <= MAX_POCOS_VALORES and filas >= 2 * distintos,
        summable=sumable,
        top_values=top,
        sample_values=muestra,
        min=minimo,
        max=maximo,
        sum=suma,
        mean=promedio,
        min_len=min_len,
        max_len=max_len,
        leading_zeros=ceros_n,
        untrimmed=sin_recortar_n,
        case_variants=variantes,
        not_converted=no_convertidas,
        notes=notas,
    )


# ---------------------------------------------------------------- la tabla


def _avisos(data: ReadResult, columnas: list[ColumnProfile]) -> list[str]:
    avisos = list(data.warnings)
    for c in columnas:
        etiqueta = f"La columna {c.letter}" if not c.name.strip() else f'La columna "{c.name}"'
        if c.not_converted:
            avisos.append(f"{etiqueta} tiene {c.not_converted} valores que no son números.")
        if c.untrimmed:
            avisos.append(f"{etiqueta} tiene {c.untrimmed} valores con espacios de más.")
        if c.case_variants:
            avisos.append(f"{etiqueta} tiene valores que solo cambian en mayúsculas o espacios.")
    return avisos


def _info_archivo(data: ReadResult) -> FileInfo:
    estado = data.path.stat()
    return FileInfo(
        name=data.path.name,
        path=data.path,
        size=estado.st_size,
        modified=datetime.fromtimestamp(estado.st_mtime),
        format=data.format,
    )


def profile_table(
    data: ReadResult,
    *,
    sheets: list[SheetSummary] | None = None,
    on_column: Callable[[ColumnProfile], None] | None = None,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> TableProfile:
    """Perfil de lo ya leído (no vuelve a leer el archivo)."""
    df = data.df
    filas, n_cols = df.shape
    reporte = ThrottledProgress(progress)
    hash_filas = np.zeros(filas, dtype=np.uint64)
    fila_vacia = np.ones(filas, dtype=bool)
    columnas: list[ColumnProfile] = []
    for i in range(n_cols):
        if cancel:
            cancel.check()
        nombre = str(df.columns[i])
        etiqueta = nombre if nombre.strip() else column_letter(i)
        reporte.force(ProgressEvent(f"Analizando la columna {etiqueta}…", i, n_cols))
        col = _perfilar_columna(nombre, i, df.iloc[:, i], filas, cancel, hash_filas, fila_vacia)
        columnas.append(col)
        if on_column:
            on_column(col)
    if cancel:
        cancel.check()
    reporte.force(ProgressEvent("Buscando filas repetidas…", n_cols, n_cols))
    vacias = int(fila_vacia.sum()) if n_cols else 0
    if n_cols:
        usados = hash_filas[~fila_vacia]
        repetidas = len(usados) - len(np.unique(usados))
    else:
        repetidas = 0
    return TableProfile(
        file=_info_archivo(data),
        sheets=list(sheets) if sheets else [],
        sheet=data.sheet,
        csv_options=data.csv_options,
        header_row=data.header_row,
        rows=filas,
        columns=n_cols,
        empty_rows=vacias,
        duplicate_rows=repetidas,
        columns_profile=columnas,
        warnings=_avisos(data, columnas),
    )


# ---------------------------------------------------------------- exportar

HOJA_GENERAL = "General"
HOJA_COLUMNAS = "Columnas"
HOJA_AVISOS = "Avisos"
HOJA_DISTINTOS = "Valores distintos"

_NOMBRES_TIPO = {
    "texto": "Texto",
    "entero": "Número entero",
    "decimal": "Número decimal",
    "fecha": "Fecha",
    "fecha_hora": "Fecha y hora",
    "si_no": "Sí/No",
    "mixto": "Mixto",
    "vacia": "Vacía",
}
_NOMBRES_FORMATO = {
    "xlsx": "Excel (.xlsx)",
    "xlsm": "Excel con macros (.xlsm)",
    "xlsb": "Excel binario (.xlsb)",
    "xls": "Excel antiguo (.xls)",
    "ods": "Hoja de cálculo (.ods)",
    "csv": "Texto con columnas (.csv)",
}
_SEPARADORES = {";": "Punto y coma", ",": "Coma", "\t": "Tabulación", "|": "Barra vertical"}


def _miles(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def _decimal_ar(valor: float, decimales: int = 2) -> str:
    """1234.5 -> '1.234,50' (para mostrar como texto)."""
    return f"{valor:,.{decimales}f}".replace(",", "\0").replace(".", ",").replace("\0", ".")


def _tamano(tamano: int) -> str:
    if tamano < 1024:
        return f"{_miles(tamano)} bytes"
    valor = float(tamano)
    for unidad in ("KB", "MB", "GB", "TB"):
        valor /= 1024
        if valor < 1024 or unidad == "TB":
            return f"{_decimal_ar(valor, 1)} {unidad}"
    return ""  # pragma: no cover


def _numero_para_celda(valor: float | None, decimales: int = 4) -> str:
    """Texto con coma decimal y sin separador de miles: lo convierte a número `write_*`."""
    if valor is None or valor != valor:
        return ""
    if float(valor).is_integer() and abs(valor) < 1e15:
        return str(int(valor))
    return np.format_float_positional(round(float(valor), decimales), trim="-").replace(".", ",")


def _datos_generales(perfil: TableProfile) -> pd.DataFrame:
    a = perfil.file
    filas: list[tuple[str, str]] = [
        ("Nombre del archivo", a.name),
        ("Carpeta", str(a.path.parent)),
        ("Tamaño", _tamano(a.size)),
        (
            "Última modificación",
            f"{a.modified.day:02d}/{a.modified.month:02d}/{a.modified.year:04d} "
            f"{a.modified.hour:02d}:{a.modified.minute:02d}",
        ),
        ("Formato", _NOMBRES_FORMATO.get(a.format, str(a.format))),
    ]
    if perfil.sheet:
        filas.append(("Hoja", perfil.sheet))
    if perfil.sheets:
        filas.append(("Hojas del archivo", ", ".join(h.name for h in perfil.sheets)))
    filas += [
        ("Filas", _miles(perfil.rows)),
        ("Columnas", _miles(perfil.columns)),
        ("Fila de los títulos", _miles(perfil.header_row + 1)),
        ("Filas vacías", _miles(perfil.empty_rows)),
        ("Filas repetidas", _miles(perfil.duplicate_rows)),
    ]
    if perfil.csv_options is not None:
        csv = perfil.csv_options
        filas.append(("Separador de columnas", _SEPARADORES.get(csv.separator, csv.separator)))
        filas.append(("Tipo de texto", csv.encoding))
    return pd.DataFrame(filas, columns=["Dato", "Valor"])


def _marcas(c: ColumnProfile) -> str:
    marcas = []
    if c.is_unique:
        marcas.append("Posible clave")
    if c.few_values:
        marcas.append("Pocos valores")
    if c.summable:
        marcas.append("Se puede sumar")
    return " · ".join(marcas)


_TIPOS_COLUMNAS: dict[str, ColumnType] = {
    "Con datos": "entero",
    "Vacías": "entero",
    "Valores distintos": "entero",
    "Suma": "decimal",
    "Promedio": "decimal",
}


def _datos_columnas(perfil: TableProfile) -> pd.DataFrame:
    filas = []
    for c in perfil.columns_profile:
        filas.append(
            {
                "Columna": f"{c.letter} ({c.position + 1})",
                "Nombre": c.name,
                "Tipo": _NOMBRES_TIPO.get(c.detected_type, c.detected_type),
                "Con datos": str(c.non_empty),
                "Vacías": str(c.empty),
                "Valores distintos": str(c.distinct),
                "Marcas": _marcas(c),
                "Mínimo": c.min or "",
                "Máximo": c.max or "",
                "Suma": _numero_para_celda(c.sum),
                "Promedio": _numero_para_celda(c.mean),
                "Ejemplos": ", ".join(c.sample_values),
                "Observaciones": " ".join(c.notes),
            }
        )
    columnas = ["Columna", "Nombre", "Tipo", "Con datos", "Vacías", "Valores distintos", "Marcas"]
    columnas += ["Mínimo", "Máximo", "Suma", "Promedio", "Ejemplos", "Observaciones"]
    return pd.DataFrame(filas, columns=columnas, dtype=object).fillna("")


def _tabla_distintos(resultado: DistinctResult) -> pd.DataFrame:
    """Tabla lista para escribir: cantidades y porcentajes como texto numérico argentino
    (con `column_types` quedan como número en el archivo)."""
    t = resultado.table.copy()
    salida = pd.DataFrame(index=range(len(t)))
    for i, nombre in enumerate(t.columns):
        col = t.iloc[:, i]
        if nombre == COL_CANTIDAD and i >= len(resultado.columns):
            salida[nombre] = [_numero_para_celda(float(v), 0) for v in col]
        elif nombre == COL_PORCENTAJE and i >= len(resultado.columns):
            salida[nombre] = [_numero_para_celda(float(v)) for v in col]
        else:
            salida[nombre] = col.astype(object).where(col.notna(), "").astype(str).to_numpy()
    return salida


_TIPOS_DISTINTOS: dict[str, ColumnType] = {COL_CANTIDAD: "entero", COL_PORCENTAJE: "decimal"}


def _aviso_truncado(resultado: DistinctResult) -> str | None:
    if not resultado.truncated:
        return None
    return (
        f"La lista de valores de {', '.join(resultado.columns)} quedó incompleta: hay "
        f"{_miles(resultado.distinct_total)} valores distintos y solo se guardaron "
        f"{_miles(len(resultado.table))}."
    )


def export_profile(
    profile: TableProfile, distincts: list[DistinctResult], path: Path
) -> list[WriteResult]:
    """Guarda el informe en un .xlsx: hojas General, Columnas, una por cada lista de valores
    distintos y Avisos (solo si hay avisos). Un resultado por hoja, en ese orden."""
    usados: set[str] = set()
    hojas: dict[str, pd.DataFrame] = {}
    tipos: dict[str, dict[str, ColumnType]] = {}

    def agregar(nombre: str, tabla: pd.DataFrame, tipos_hoja: dict[str, ColumnType] | None = None):
        nombre_ok = nombre_de_hoja(nombre, usados)
        hojas[nombre_ok] = tabla
        if tipos_hoja:
            tipos[nombre_ok] = {k: v for k, v in tipos_hoja.items() if k in tabla.columns}

    agregar(HOJA_GENERAL, _datos_generales(profile))
    agregar(HOJA_COLUMNAS, _datos_columnas(profile), _TIPOS_COLUMNAS)
    avisos = list(profile.warnings)
    for r in distincts:
        agregar(" + ".join(r.columns) or HOJA_DISTINTOS, _tabla_distintos(r), _TIPOS_DISTINTOS)
        if (aviso := _aviso_truncado(r)) is not None:
            avisos.append(aviso)
    if avisos:
        agregar(HOJA_AVISOS, pd.DataFrame({"Aviso": avisos}))
    return write_sheets(hojas, Path(path), tipos)


def export_distincts(result: DistinctResult, path: Path, format: WriteFormat) -> WriteResult:
    """Guarda una lista de valores distintos (Cantidad y Porcentaje como número)."""
    opciones = WriteOptions(
        format=format, sheet_name=HOJA_DISTINTOS, column_types=dict(_TIPOS_DISTINTOS)
    )
    return write_table(_tabla_distintos(result), Path(path), opciones)
