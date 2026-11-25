"""Lectura de CSV/TXT: detección de separador y encoding, y lectura como texto."""

from __future__ import annotations

import codecs
import csv
import logging
from collections.abc import Iterator
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

from excel_tool.core.errors import (
    ArchivoBloqueado,
    ArchivoIlegible,
    ArchivoNoEncontrado,
    ExcelToolError,
)
from excel_tool.core.io import CsvOptions, ReadOptions, ReadResult
from excel_tool.core.io.header import detect_header_row, normalize_column_names
from excel_tool.core.progress import CancelToken, ProgressCallback, ProgressEvent, ThrottledProgress

log = logging.getLogger(__name__)

SEPARADORES = ";,\t|"
ENCODINGS = ("utf-8-sig", "utf-8", "cp1252")
BYTES_MUESTRA = 256 * 1024
FILAS_MUESTRA_ENCABEZADO = 30
FILAS_POR_BLOQUE = 100_000


def _leer_bytes_iniciales(path: Path) -> bytes:
    try:
        with path.open("rb") as f:
            return f.read(BYTES_MUESTRA)
    except FileNotFoundError as e:
        raise ArchivoNoEncontrado(
            "No se encontró el archivo. Fijate si se movió o se borró.", str(e)
        ) from e
    except PermissionError as e:
        raise ArchivoBloqueado("Cerrá el archivo en Excel y volvé a intentar.", str(e)) from e
    except OSError as e:
        raise ArchivoIlegible("No se pudo abrir el archivo.", str(e)) from e


def _detectar_encoding(muestra: bytes) -> str:
    if muestra.startswith(codecs.BOM_UTF8):
        return "utf-8-sig"
    # Decodificador incremental: tolera que la muestra corte un carácter a la mitad.
    try:
        codecs.getincrementaldecoder("utf-8")().decode(muestra, final=False)
        return "utf-8"
    except UnicodeDecodeError:
        return "cp1252"


def _puntaje_separador(lineas: list[str], sep: str) -> float:
    """Fracción de filas con la cantidad de campos más común, si esa cantidad es >= 2."""
    if not lineas:
        return 0.0
    cuentas = [len(f) for f in csv.reader(lineas, delimiter=sep)]
    cuentas = [c for c in cuentas if c > 0]
    if not cuentas:
        return 0.0
    moda = max(set(cuentas), key=cuentas.count)
    if moda < 2:
        return 0.0
    return cuentas.count(moda) / len(cuentas)


def _detectar_separador(texto: str) -> str:
    lineas = texto.splitlines()
    if len(texto) >= BYTES_MUESTRA // 2 and lineas:
        lineas = lineas[:-1]  # la última puede estar cortada
    lineas = [x for x in lineas if x.strip()][:200]
    if not lineas:
        return ";"
    puntajes = {s: _puntaje_separador(lineas, s) for s in SEPARADORES}
    try:
        elegido = csv.Sniffer().sniff("\n".join(lineas), delimiters=SEPARADORES).delimiter
    except csv.Error:
        elegido = None
    # El Sniffer a veces confunde la coma decimal con separador: se acepta solo si es consistente.
    if elegido and puntajes.get(elegido, 0.0) >= 0.8:
        return elegido
    mejor = max(SEPARADORES, key=lambda s: puntajes[s])
    return mejor if puntajes[mejor] > 0 else ";"


def detect_csv_options(path: Path) -> CsvOptions:
    """Separador y encoding de un CSV, leyendo solo el comienzo del archivo."""
    path = Path(path)
    muestra = _leer_bytes_iniciales(path)
    if not muestra.strip():
        raise ArchivoIlegible("El archivo está vacío.", f"{path}: sin contenido")
    encoding = _detectar_encoding(muestra)
    texto = muestra.decode(encoding, errors="replace")
    return CsvOptions(separator=_detectar_separador(texto), encoding=encoding)


def _filas_muestra(path: Path, opciones: CsvOptions, cuantas: int) -> list[list[str]]:
    """Primeras filas como listas de texto (para detectar el encabezado)."""
    filas: list[list[str]] = []
    with path.open("r", encoding=opciones.encoding, newline="") as f:
        for fila in csv.reader(f, delimiter=opciones.separator):
            filas.append(fila)
            if len(filas) >= cuantas:
                break
    return filas


def _ancho_maximo(path: Path, opciones: CsvOptions, cuantas: int = 2000) -> int:
    ancho = 1
    with path.open("r", encoding=opciones.encoding, newline="") as f:
        for i, fila in enumerate(csv.reader(f, delimiter=opciones.separator)):
            ancho = max(ancho, len(fila))
            if i >= cuantas:
                break
    return ancho


def _leer_crudo(
    path: Path,
    opciones: CsvOptions,
    nfilas: int | None,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> pd.DataFrame:
    """Todo el archivo (o las primeras `nfilas` filas) como texto, sin encabezado."""
    if cancel is not None:
        cancel.check()
    ancho = _ancho_maximo(path, opciones)
    try:
        if nfilas is not None:
            return _leer_pandas(path, opciones, ancho, nfilas)
        # Por bloques: cada bloque pasa a textos compactos (pyarrow) antes de leer el
        # siguiente, y el pico de memoria baja respecto de leer todo de una vez.
        reporte = ThrottledProgress(progress)
        total = max(path.stat().st_size, 1)
        mensaje = "Leyendo el archivo…"
        reporte(ProgressEvent(mensaje, 0, total))
        bloques: list[pd.DataFrame] = []
        # Handle binario propio: `tell()` da los bytes leídos y se cierra siempre.
        with path.open("rb") as f:
            lector = _leer_pandas(f, opciones, ancho, None)
            try:
                for bloque in lector:
                    bloques.append(bloque)
                    reporte(ProgressEvent(mensaje, min(f.tell(), total), total))
                    if cancel is not None:
                        cancel.check()
            finally:
                lector.close()
        reporte.force(ProgressEvent(mensaje, total, total))
        if not bloques:
            raise pd.errors.EmptyDataError("No columns to parse from file")
        return pd.concat(bloques, ignore_index=True)
    except UnicodeDecodeError:
        raise
    except pd.errors.ParserError:
        # Filas con más campos que las de la muestra: lectura lenta pero tolerante.
        return _leer_con_csv(path, opciones, nfilas, cancel)


def _leer_pandas(fuente: Any, opciones: CsvOptions, ancho: int, nfilas: int | None) -> Any:
    """`read_csv` de pandas: un DataFrame (con `nfilas`) o un lector por bloques."""
    return pd.read_csv(
        fuente,
        sep=opciones.separator,
        encoding=opciones.encoding,
        header=None,
        names=list(range(ancho)),
        dtype=str,
        keep_default_na=False,
        na_filter=False,
        skip_blank_lines=False,
        nrows=nfilas,
        quotechar='"',
        engine="c",
        chunksize=None if nfilas is not None else FILAS_POR_BLOQUE,
    )


def _leer_con_csv(
    path: Path, opciones: CsvOptions, nfilas: int | None, cancel: CancelToken | None = None
) -> pd.DataFrame:
    filas: list[list[str]] = []
    with path.open("r", encoding=opciones.encoding, newline="") as f:
        for fila in csv.reader(f, delimiter=opciones.separator):
            filas.append(fila)
            if cancel is not None and len(filas) % 50_000 == 0:
                cancel.check()
            if nfilas is not None and len(filas) >= nfilas:
                break
    ancho = max((len(f) for f in filas), default=1)
    filas = [f + [""] * (ancho - len(f)) for f in filas]
    return pd.DataFrame(filas, columns=list(range(ancho)), dtype=str)


def _recortar_vacias_al_final(df: pd.DataFrame) -> pd.DataFrame:
    """Saca las filas totalmente vacías del final (no las del medio)."""
    n = len(df)
    bloque = 5000
    fin = n
    while fin > 0:
        ini = max(0, fin - bloque)
        con_datos = (df.iloc[ini:fin] != "").any(axis=1).to_numpy()
        if con_datos.any():
            ultima = ini + int(con_datos.nonzero()[0].max())
            return df.iloc[: ultima + 1]
        fin = ini
    return df.iloc[:0]


def _leer_con(
    path: Path,
    opciones: CsvOptions,
    nfilas: int | None,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> pd.DataFrame:
    try:
        return _leer_crudo(path, opciones, nfilas, progress, cancel)
    except PermissionError as e:
        raise ArchivoBloqueado("Cerrá el archivo en Excel y volvé a intentar.", str(e)) from e
    except FileNotFoundError as e:
        raise ArchivoNoEncontrado(
            "No se encontró el archivo. Fijate si se movió o se borró.", str(e)
        ) from e
    except pd.errors.EmptyDataError as e:
        raise ArchivoIlegible("El archivo está vacío.", str(e)) from e


def read_csv(
    path: Path,
    options: ReadOptions | None = None,
    *,
    progress: ProgressCallback | None = None,
    cancel: CancelToken | None = None,
) -> ReadResult:
    """Lee un CSV/TXT como texto. Lo usa `read_table`."""
    path = Path(path)
    options = options or ReadOptions()
    detectado = options.csv is None
    opciones = options.csv or detect_csv_options(path)
    if not detectado:
        _leer_bytes_iniciales(path)  # errores de archivo inexistente/bloqueado

    try:
        try:
            fila_enc = options.header_row
            if fila_enc is None:
                fila_enc = detect_header_row(
                    _filas_muestra(path, opciones, FILAS_MUESTRA_ENCABEZADO)
                )
            nfilas = None if options.nrows is None else fila_enc + 1 + options.nrows
            crudo = _leer_con(path, opciones, nfilas, progress, cancel)
        except UnicodeDecodeError as e:
            if not detectado or opciones.encoding == "cp1252":
                raise ArchivoIlegible(
                    "No se pudo leer el texto del archivo (caracteres no reconocidos).", str(e)
                ) from e
            # La muestra parecía utf-8 pero más adelante no: se reintenta con cp1252.
            log.info("Reintentando %s con cp1252: %s", path, e)
            opciones = CsvOptions(separator=opciones.separator, encoding="cp1252")
            fila_enc = options.header_row
            if fila_enc is None:
                fila_enc = detect_header_row(
                    _filas_muestra(path, opciones, FILAS_MUESTRA_ENCABEZADO)
                )
            nfilas = None if options.nrows is None else fila_enc + 1 + options.nrows
            crudo = _leer_con(path, opciones, nfilas, progress, cancel)
    except csv.Error as e:
        raise ArchivoIlegible("El archivo está dañado o no se pudo leer.", str(e)) from e
    except OSError as e:
        raise ArchivoIlegible("No se pudo leer el archivo.", str(e)) from e

    crudo = _recortar_vacias_al_final(crudo)
    if len(crudo) == 0 or fila_enc >= len(crudo):
        raise ArchivoIlegible(
            "El archivo no tiene datos en la fila de encabezado indicada."
            if len(crudo)
            else "El archivo está vacío.",
            f"{path}: {len(crudo)} filas",
        )

    nombres, renamed, avisos = normalize_column_names(
        [str(x) for x in crudo.iloc[fila_enc].tolist()]
    )
    df = crudo.iloc[fila_enc + 1 :].reset_index(drop=True)
    df.columns = nombres
    if options.nrows is not None:
        df = df.iloc[: options.nrows]
    return ReadResult(
        df=df,
        path=path,
        format="csv",
        sheet=None,
        sheets=[],
        csv_options=opciones,
        header_row=fila_enc,
        renamed_columns=renamed,
        warnings=avisos,
    )


# ------------------------------------------------------------------ lectura por bloques (F3-05)


@dataclass
class CsvStreamInfo:
    columns: list[str]
    csv_options: CsvOptions
    header_row: int
    renamed_columns: dict[str, str]
    warnings: list[str]
    total_bytes: int = 0


def _envolver_lectura(exc: OSError) -> Exception:
    if isinstance(exc, PermissionError):
        return ArchivoBloqueado("Cerrá el archivo en Excel y volvé a intentar.", str(exc))
    if isinstance(exc, FileNotFoundError):
        return ArchivoNoEncontrado(
            "No se encontró el archivo. Fijate si se movió o se borró.", str(exc)
        )
    return ArchivoIlegible("No se pudo leer el archivo.", str(exc))


class CsvBlockReader:
    """Lee un CSV por bloques de `FILAS_POR_BLOQUE` filas, sin cargarlo entero.

    Mismas reglas que `read_csv` (texto, encabezado detectado, nombres normalizados, filas vacías
    del final descartadas), pero entrega los datos de a bloques (DataFrames con `columns` y
    índice desde 0). Se usa como `with CsvBlockReader(path, options) as lector:`.

    Un `UnicodeDecodeError` (texto que no es del encoding usado) se deja pasar tal cual, para que
    quien llama pueda reintentar con otro encoding. Las filas con más campos que las primeras
    2000 dan `ArchivoIlegible` (no hay forma de seguir leyendo por bloques sin perder columnas).
    """

    def __init__(self, path: Path, options: ReadOptions | None = None) -> None:
        self.path = Path(path)
        options = options or ReadOptions()
        opciones = options.csv or detect_csv_options(self.path)
        self._f: Any = None
        self._lector: Any = None
        try:
            if options.csv is not None:
                _leer_bytes_iniciales(self.path)
            fila_enc = options.header_row
            if fila_enc is None:
                fila_enc = detect_header_row(
                    _filas_muestra(self.path, opciones, FILAS_MUESTRA_ENCABEZADO)
                )
            ancho = _ancho_maximo(self.path, opciones)
            self._f = self.path.open("rb")
            self._lector = _leer_pandas(self._f, opciones, ancho, None)
            try:
                primero = next(self._lector)
            except StopIteration:
                primero = None
            except pd.errors.EmptyDataError:
                primero = None
            if primero is None or fila_enc >= len(primero):
                raise ArchivoIlegible(
                    "El archivo no tiene datos en la fila de encabezado indicada."
                    if primero is not None
                    else "El archivo está vacío.",
                    f"{self.path}: sin filas suficientes",
                )
        except csv.Error as e:
            self.close()
            raise ArchivoIlegible("El archivo está dañado o no se pudo leer.", str(e)) from e
        except pd.errors.ParserError as e:
            self.close()
            raise self._error_de_columnas(e) from e
        except OSError as e:
            self.close()
            raise _envolver_lectura(e) from e
        except BaseException:
            self.close()
            raise
        nombres, renamed, avisos = normalize_column_names(
            [str(x) for x in primero.iloc[fila_enc].tolist()]
        )
        self.info = CsvStreamInfo(
            columns=nombres,
            csv_options=opciones,
            header_row=fila_enc,
            renamed_columns=renamed,
            warnings=avisos,
            total_bytes=max(self.path.stat().st_size, 1),
        )
        self._primero = primero.iloc[fila_enc + 1 :]
        self.rows = 0  # filas de datos entregadas hasta ahora

    @staticmethod
    def _error_de_columnas(e: Exception) -> ArchivoIlegible:
        return ArchivoIlegible(
            "Hay filas con más columnas que las del comienzo del archivo; no se puede "
            "procesar sin abrirlo entero.",
            str(e),
        )

    @property
    def bytes_read(self) -> int:
        try:
            return min(self._f.tell(), self.info.total_bytes) if self._f else 0
        except (OSError, ValueError):
            return 0

    def close(self) -> None:
        if self._lector is not None:
            try:
                self._lector.close()
            except Exception:  # noqa: BLE001 - cerrar nunca debe tapar el error original
                log.debug("No se pudo cerrar el lector", exc_info=True)
            self._lector = None
        if self._f is not None:
            self._f.close()
            self._f = None

    def __enter__(self) -> CsvBlockReader:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def _bloques_crudos(self) -> Iterator[pd.DataFrame]:
        yield self._primero
        if self._lector is None:
            return
        try:
            yield from self._lector
        except pd.errors.ParserError as e:
            raise self._error_de_columnas(e) from e
        except csv.Error as e:
            raise ArchivoIlegible("El archivo está dañado o no se pudo leer.", str(e)) from e
        except OSError as e:
            raise _envolver_lectura(e) from e

    def __iter__(self) -> Iterator[pd.DataFrame]:
        # Las filas vacías se retienen: solo se descartan si son las últimas del archivo.
        pendientes: list[pd.DataFrame] = []
        for crudo in self._bloques_crudos():
            bloque = crudo.reset_index(drop=True)
            bloque.columns = self.info.columns
            recortado = _recortar_vacias_al_final(bloque)
            if len(recortado) == 0:
                pendientes.append(bloque)
                continue
            for p in pendientes:
                self.rows += len(p)
                yield p
            pendientes = []
            if len(recortado) < len(bloque):
                pendientes.append(bloque.iloc[len(recortado) :])
            self.rows += len(recortado)
            yield recortado


def read_csv_exact(path: Path, options: CsvOptions) -> pd.DataFrame:
    """CSV ya escrito por esta app (encabezado en la primera fila) como texto, sin tocar nada:
    ni se descartan filas vacías ni se renombran columnas. Para convertir partes a xlsx."""
    try:
        return pd.read_csv(
            Path(path),
            sep=options.separator,
            encoding=options.encoding,
            header=0,
            dtype=str,
            keep_default_na=False,
            na_filter=False,
            skip_blank_lines=False,
            quotechar='"',
            engine="c",
        )
    except OSError as e:
        raise _envolver_lectura(e) from e
    except (UnicodeDecodeError, pd.errors.ParserError) as e:
        raise ArchivoIlegible("No se pudo leer una parte recién escrita.", str(e)) from e


# ------------------------------------------------------------------ escritura incremental CSV

# Igual que `_CSV_FORMATOS` de writer.py (un test verifica que coincidan):
# formato -> (separador, encoding)
CSV_ESTILOS = {
    "csv_excel": (";", "utf-8-sig"),
    "csv_google": (",", "utf-8"),
    "csv_standard": (",", "utf-8-sig"),
}


@dataclass
class CsvPartWriter:
    """Escribe un CSV de a bloques: abre, escribe el encabezado, agrega bloques y cierra.

    Los bloques se escriben como texto, tal cual (sin convertir números ni fechas), con el
    mismo formato que `write_table` (separador, encoding, `\r\n`, `newline=""`).
    """

    path: Path
    formato: str
    columns: list[str]
    rows: int = 0
    _f: Any = field(default=None, repr=False)
    _sep: str = field(default=";", repr=False)

    def __post_init__(self) -> None:
        if self.formato not in CSV_ESTILOS:
            raise ExcelToolError("Formato de salida no reconocido.", str(self.formato))
        self._sep, encoding = CSV_ESTILOS[self.formato]
        self.path = Path(self.path)
        try:
            self._f = self.path.open("w", encoding=encoding, newline="")
            pd.DataFrame(columns=self.columns).to_csv(
                self._f, sep=self._sep, index=False, header=True, lineterminator="\r\n"
            )
        except OSError as e:
            self.close()
            raise _error_de_escritura(e) from e

    def write(self, bloque: pd.DataFrame) -> None:
        if self._f is None:
            raise ExcelToolError("Se intentó escribir en un archivo ya cerrado.")
        try:
            bloque.to_csv(self._f, sep=self._sep, index=False, header=False, lineterminator="\r\n")
        except OSError as e:
            raise _error_de_escritura(e) from e
        self.rows += len(bloque)

    def close(self) -> None:
        if self._f is not None:
            f, self._f = self._f, None
            try:
                f.close()
            except OSError as e:
                raise _error_de_escritura(e) from e

    def __enter__(self) -> CsvPartWriter:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()


def _error_de_escritura(exc: OSError) -> ExcelToolError:
    if isinstance(exc, PermissionError):
        return ArchivoBloqueado(
            "No se pudo guardar el archivo. Si está abierto en Excel, cerralo y probá de nuevo.",
            str(exc),
        )
    if isinstance(exc, FileNotFoundError):
        return ExcelToolError("No se pudo guardar el archivo: la carpeta no existe.", str(exc))
    return ExcelToolError(
        "No se pudo guardar el archivo. Revisá que haya espacio en el disco "
        "y que el nombre sea válido.",
        str(exc),
    )


__all__ = [
    "CSV_ESTILOS",
    "CsvBlockReader",
    "CsvPartWriter",
    "CsvStreamInfo",
    "detect_csv_options",
    "read_csv",
    "read_csv_exact",
]
