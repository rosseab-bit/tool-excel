# Contratos core ↔ UI

Dataclasses y firmas públicas que cruzan de `src/excel_tool/core/` a `src/excel_tool/ui/`.

- **Lo mantiene el líder técnico.** Ningún agente lo edita: si hace falta un cambio, termina su tarea con estado BLOQUEADO y lo pide.
- Cambiar un contrato existente es nivel 2 si no cambia lo que ve el usuario, y nivel 3 si lo cambia.
- Punto de partida: los modelos de `PLAN.md` §2. El líder los completa y los cierra al planificar cada fase, antes de despachar al backend y al frontend.

## Estado

| Contrato | Fase | Estado |
|---|---|---|
| `app_entry.main` (`--version`, `--self-test`) | 0 | Cerrado |
| `ui.app.run` | 0 | Cerrado |
| `platform_utils` (carpetas de usuario, abrir carpeta, nombres seguros) | 0 | Cerrado |
| `core/errors` (base), `ReadOptions`, `CsvOptions`, `SheetSummary`, `ReadResult`, `WriteOptions`, `read_table`, `list_sheets`, `detect_csv_options`, `write_table`, `write_sheets` | 1 | Cerrado |
| `ColumnType`, `TypeDetection`, `ConversionResult`, funciones de `core/types.py` | 1 | Cerrado |
| `core/progress`, `TableProfile`, `ColumnProfile`, `profile_table`, `value_counts`, `rows_with_value`, filtros, orden, páginas, `export_profile`, preferencias | 2 | Cerrado (OK PM, ADR-023); implementado |
| `ui/worker` (hilo + cola + progreso + cancelar) | 2 | Cerrado (OK PM, ADR-023); implementado |
| Errores del core hacia la UI (cómo se muestran) | 2 | Cerrado (OK PM, ADR-023); implementado |
| `SizeEstimate`, `SizeCheck`, `SplitSpec`, `OutputSpec`, `save_result` | 3 | Borrador en PLAN.md §2 |
| `AggregationSpec`, `JoinPlan`, `SourceSpec`, `SecondarySpec`, `KeyOptions`, `diagnose_join`, `run_join` | 4 | Borrador en PLAN.md §2 |

## Contratos

### Fase 0

#### `excel_tool/__init__.py`

```python
__version__: str  # "0.1.0" en la Fase 0. Única fuente de la versión (pyproject la lee de acá).
APP_NAME = "Excel Join Tool"
APP_DIR_NAME = "ExcelJoinTool"  # nombre de carpeta para logs/preferencias
```

#### `excel_tool/app_entry.py` (backend)

```python
def main(argv: list[str] | None = None) -> int: ...
```

- `argv=None` → usa `sys.argv[1:]`. Lo llaman `python -m excel_tool`, el comando `excel-tool` y `packaging/launcher.py` (`sys.exit(main())`).
- Lo primero que hace, en todos los modos: configurar el log a archivo en `platform_utils.log_dir() / "excel_tool.log"` (rotativo, UTF-8). Si no puede crear la carpeta de logs, sigue sin log: nunca falla por eso.
- **Sin argumentos:** importa `excel_tool.ui.app` recién ahí (import diferido) y llama `run()`. Devuelve 0. Si hay una excepción no manejada: la loguea completa, muestra un `messagebox` amigable si se puede y devuelve 1.
- **`--version`:** devuelve 0. Escribe `"Excel Join Tool <versión>"` a stdout **solo si `sys.stdout` no es `None`** (en el `.exe` sin consola es `None`). También lo loguea.
- **`--self-test`:** devuelve 0 si todo pasa, 1 si algo falla. Pasos, cada uno logueado como `SELF-TEST OK|FALLA <paso>: <detalle>`:
  1. Importar todas las librerías empaquetadas (pandas, numpy, openpyxl, python_calamine, tkinter) y `excel_tool.ui.app`.
  2. Leer `limites_default.json` con `importlib.resources`.
  3. En una carpeta de `tempfile`: escribir y releer un .xlsx y un .csv (`;`, `utf-8-sig`) con `"Ñandú"`, `"acción"` y `"00123"`; verificar que vuelven iguales leyendo como texto. (En la Fase 0 se escribe y lee directo con pandas/openpyxl/calamine dentro de `app_entry`; en la Fase 1 pasa a `core/io`.)
  4. Crear una ventana `tk.Tk()`, `update()`, `destroy()`.
  - Si stdout existe, imprime un resumen de una línea al final. El resultado que mira el script de compilación es el **código de salida**.
- Argumento desconocido → devuelve 2 (y mensaje de uso a stderr solo si existe).

#### `excel_tool/ui/app.py` (frontend)

```python
def run() -> None: ...
```

- Crea la ventana principal, entra en `mainloop()` y vuelve cuando el usuario la cierra.
- Fase 0: título `"Excel Join Tool"`, versión visible (`excel_tool.__version__`), ícono desde `importlib.resources.files("excel_tool") / "assets"` (si el ícono falla, sigue sin ícono y lo loguea).
- No configura el logging (eso lo hace `app_entry`): usa `logging.getLogger(__name__)`.
- Para tests: la ventana se construye en una función o clase separada de `mainloop()` (nivel 1 del frontend: cómo), así un test la crea y la destruye sin bloquear.

#### `excel_tool/platform_utils.py` (backend)

```python
def user_data_dir() -> Path: ...
def log_dir() -> Path: ...
def open_folder(path: Path) -> None: ...
def safe_filename(name: str, replacement: str = "_") -> str: ...
```

- `user_data_dir()`: Windows `%LOCALAPPDATA%\ExcelJoinTool` (si falta la variable, `~\AppData\Local\ExcelJoinTool`); Linux `$XDG_DATA_HOME/ExcelJoinTool` o `~/.local/share/ExcelJoinTool`. La crea si no existe.
- `log_dir()`: `user_data_dir() / "logs"`, creada si no existe.
- `open_folder(path)`: Windows `os.startfile`; Linux `xdg-open` con `subprocess.Popen` (sin shell, sin esperar; ADR-012). Si falla, lanza `OSError` con mensaje en castellano (la UI lo muestra amigable).
- `safe_filename(name, replacement)`: reemplaza `<>:"/\|?*` y caracteres de control por `replacement`; recorta espacios y puntos al final; si el nombre (sin extensión, sin importar mayúsculas) es reservado de Windows (`CON`, `PRN`, `AUX`, `NUL`, `COM1`–`COM9`, `LPT1`–`LPT9`) le agrega `replacement` al final; si queda vacío devuelve `"sin_nombre"`; limita a 150 caracteres preservando la extensión. Se comporta **igual en Linux y Windows** (los archivos se pueden copiar de un sistema al otro).
- Es el **único** módulo que consulta `sys.platform` / `os.name`.

#### `excel_tool/limites_default.json` (backend)

Valores de PLAN §1.2. Lo lee y valida `core/limits.py` en la Fase 3; en la Fase 0 solo se crea y el self-test verifica que se pueda leer.

```json
{
  "version": 1,
  "excel":         {"max_rows": 1048576, "max_columns": 16384, "comfortable_rows": 300000, "comfortable_mb": 50},
  "google_sheets": {"max_cells": 20000000, "max_import_mb": 100, "comfortable_cells": 5000000, "comfortable_mb": 50}
}
```

### Fase 1

Todo vive en `src/excel_tool/core/`. Los valores de las tablas son **siempre texto** (`str`); la celda vacía es `""` (nunca `NaN`/`None`). Las filas y columnas se numeran desde 0 en el core; la UI suma 1 al mostrarlas.

#### `core/errors.py`

```python
class ExcelToolError(Exception):
    """Base. `mensaje` está en castellano, sin jerga, listo para mostrar al usuario."""
    def __init__(self, mensaje: str, detalle: str | None = None): ...
    mensaje: str
    detalle: str | None      # técnico, solo para el log

class ArchivoNoEncontrado(ExcelToolError): ...
class ArchivoBloqueado(ExcelToolError): ...    # PermissionError: "Cerrá el archivo en Excel y volvé a intentar."
class FormatoNoSoportado(ExcelToolError): ...  # extensión fuera de PLAN §1.1
class ArchivoIlegible(ExcelToolError): ...     # dañado, vacío, encoding imposible, hoja inexistente
```

- El core **nunca** deja escapar `PermissionError`, `FileNotFoundError`, errores de calamine/openpyxl/pandas al leer o escribir: los envuelve en una de estas, con la excepción original en `detalle` y encadenada (`raise ... from e`).

#### `core/io/__init__.py`

```python
FileFormat = Literal["xlsx", "xlsm", "xlsb", "xls", "ods", "csv"]   # .txt se trata como "csv"
WriteFormat = Literal["xlsx", "csv_excel", "csv_google", "csv_standard"]

@dataclass(frozen=True)
class CsvOptions:
    separator: str                 # ";", ",", "\t" o "|"
    encoding: str                  # "utf-8-sig", "utf-8" o "cp1252"

@dataclass(frozen=True)
class ReadOptions:
    sheet: str | None = None       # Excel: None = primera hoja. CSV: se ignora
    header_row: int | None = None  # None = detectar. Las filas de arriba se descartan
    csv: CsvOptions | None = None  # None = detectar
    nrows: int | None = None       # None = todas (para preview: 10 o 50)

@dataclass(frozen=True)
class SheetSummary:
    name: str
    rows: int                      # filas de datos sin contar el encabezado detectado
    columns: int

@dataclass
class ReadResult:
    df: pd.DataFrame               # todas las columnas dtype object con str; "" = vacío
    path: Path
    format: FileFormat
    sheet: str | None              # hoja leída (Excel)
    sheets: list[str]              # todas las hojas (Excel); [] en CSV
    csv_options: CsvOptions | None # las usadas (detectadas o pedidas)
    header_row: int                # fila del encabezado usada (0 = primera fila del archivo/hoja)
    renamed_columns: dict[str, str]  # nombre final -> original, solo para las que se renombraron
    warnings: list[str]            # en castellano, para mostrar tal cual

def read_table(path: Path, options: ReadOptions | None = None) -> ReadResult: ...
def list_sheets(path: Path) -> list[SheetSummary]: ...     # CSV: una sola entrada con name = nombre del archivo
def detect_csv_options(path: Path) -> CsvOptions: ...

@dataclass(frozen=True)
class WriteOptions:
    format: WriteFormat
    sheet_name: str = "Datos"                       # xlsx
    column_types: dict[str, ColumnType] | None = None
    # None  -> se escribe todo como texto, exactamente como está (ida y vuelta sin cambios).
    # dict  -> las columnas "entero"/"decimal" se escriben como número, "fecha"/"fecha_hora" como fecha
    #          (convertidas con core/types); las celdas que no se pueden convertir quedan como texto
    #          y se informan en el resultado.

@dataclass
class WriteResult:
    path: Path
    rows: int
    unconverted: dict[str, list[tuple[int, str]]]  # columna -> [(fila, valor)] que quedaron como texto

def write_table(df: pd.DataFrame, path: Path, options: WriteOptions) -> WriteResult: ...
def write_sheets(sheets: dict[str, pd.DataFrame], path: Path,
                 column_types: dict[str, dict[str, ColumnType]] | None = None) -> list[WriteResult]: ...  # solo xlsx
```

**Reglas de lectura**
- Formato por extensión (sin importar mayúsculas). Otra extensión → `FormatoNoSoportado`.
- CSV: separador con `csv.Sniffer` entre `; , \t |`, mostrado en `csv_options`; encoding probando `utf-8-sig` → `utf-8` → `cp1252`.
- Excel (calamine): un número entero sale como `"100"` (nunca `"100.0"`); un decimal en notación con punto y sin exponente (`"1234.56"`); una fecha como `"dd/mm/aaaa"` y fecha con hora como `"dd/mm/aaaa hh:mm:ss"` (así `types.py` las trata igual que en un CSV argentino); booleano como `"VERDADERO"`/`"FALSO"`; error de fórmula (`#N/A`, `=NA()`) sale **vacío** (`""`): calamine no lo distingue (ADR-022). Una fecha a medianoche sin hora sale como fecha (`dd/mm/aaaa`), no como fecha y hora.
- Fórmulas y celdas combinadas se leen como valores: una advertencia por archivo si se detectan (si detectarlo es caro, se deja para la Fase 2 y se anota).
- Encabezado: `header_row=None` → detectarlo en las primeras 30 filas (la primera fila "completa" seguida de filas de datos consistentes). Si no hay nada que parezca encabezado, se usa la fila 0.
- Nombres de columna: se recortan espacios. Vacío → `"Columna <letra>"` (ej. `"Columna F"`). Repetido → `"<nombre> (2)"`, `"(3)"`… Cada renombre va en `renamed_columns` y genera una advertencia ("Hay 2 columnas sin nombre…", "Hay 2 columnas con el mismo nombre…").
- Se descartan las filas totalmente vacías **al final** del archivo/hoja (no las del medio: esas las cuenta el explorador).
- También se descartan las columnas del final que no tienen ni encabezado ni datos (hojas con formato aplicado de más). Celdas de solo hora o duración salen como `hh:mm:ss`.
- `list_sheets` cuenta las filas reales (no confiar en `total_height` de calamine: puede dar una de menos).

**Reglas de escritura**
- `csv_excel`: `;` + `utf-8-sig` + coma decimal. `csv_google`: `,` + `utf-8` + punto decimal. `csv_standard`: `,` + `utf-8-sig` + punto decimal. Siempre `newline=""`. El estilo decimal solo afecta a columnas convertidas a número (`column_types`).
- `xlsx`: encabezado congelado y filtro activado en cada hoja. Nombres de hoja válidos para Excel (máx. 31 caracteres, sin `[]:*?/\`; se ajustan y se avisa).
- Ningún límite de tamaño acá: eso es de `output.save_result` (Fase 3).
- `PermissionError` al escribir → `ArchivoBloqueado`.

#### `core/types.py`

```python
ColumnType = Literal["texto", "entero", "decimal", "fecha", "fecha_hora", "si_no", "mixto", "vacia"]
ValueKind = Literal["texto", "entero", "decimal", "fecha", "fecha_hora", "si_no", "vacia"]

@dataclass
class TypeDetection:
    detected: ColumnType
    breakdown: dict[ValueKind, int]      # cuenta por clase de valor, sin "vacia"
    empty: int                           # vacíos ("" o solo espacios)
    decimal_style: Literal[",", "."] | None   # para entero/decimal/mixto con números

@dataclass
class ConversionResult:
    values: pd.Series                    # float (número), datetime64 (fecha) o boolean; NaN/NaT/NA donde falló o estaba vacío
    failed: list[tuple[int, str]]        # (posición en la serie, valor original) de las celdas NO vacías que no se pudieron convertir

def classify_value(value: str, decimal_style: Literal[",", "."] = ",") -> ValueKind: ...
def detect_column_type(values: pd.Series) -> TypeDetection: ...
def detect_types(df: pd.DataFrame) -> dict[str, TypeDetection]: ...
def to_number(values: pd.Series, decimal_style: Literal[",", "."] | None = None) -> ConversionResult: ...
def to_date(values: pd.Series) -> ConversionResult: ...
def to_bool(values: pd.Series) -> ConversionResult: ...
```

**Reglas**
- Mira **todos** los valores de la columna, vectorizado (sin `apply` fila por fila): una columna de 2 M de filas tiene que resolverse en pocos segundos.
- Entero: `-?\d+` y con separador de miles del estilo de la columna (`1.234` en estilo `,`). **Con ceros a la izquierda (`"00123"`) es texto** (`"0"` sí es entero).
- Decimal: `1.234,56` / `1234,56` (estilo `,`) o `1,234.56` / `1234.56` (estilo `.`). El estilo se decide por columna; si es ambiguo, gana el argentino (`,`).
- Fecha: `dd/mm/aaaa` (también `d/m/aaaa` y con `-`); fecha y hora: lo mismo + `hh:mm` o `hh:mm:ss`. Fechas imposibles (31/02) son texto.
- Sí/No: `sí`, `si`, `no`, `verdadero`, `falso`, `true`, `false` (sin importar mayúsculas ni espacios). `0`/`1` son enteros.
- Columna: sin valores → `vacia`; una sola clase → esa; entero + decimal → `decimal`; fecha + fecha_hora → `fecha_hora`; cualquier otra combinación → `mixto` (con el desglose).
- Las conversiones **nunca** fallan en silencio: todo lo que no se pudo convertir va en `failed`.

### Fase 2

#### `core/progress.py` (backend)

```python
@dataclass(frozen=True)
class ProgressEvent:
    message: str          # castellano, para mostrar tal cual ("Analizando la columna Importe…")
    done: int             # pasos hechos
    total: int | None     # None = indeterminado (barra animada)

ProgressCallback = Callable[[ProgressEvent], None]

class CancelToken:
    def cancel(self) -> None: ...
    @property
    def cancelled(self) -> bool: ...
    def check(self) -> None: ...      # si se canceló, lanza Cancelado

class Cancelado(ExcelToolError): ...  # en core/errors.py; mensaje "Operación cancelada."
```

- Toda función larga del core acepta `progress: ProgressCallback | None = None` y `cancel: CancelToken | None = None` (keyword), llama a `progress` cada ~1 % o por columna (nunca más de ~10 veces por segundo) y a `cancel.check()` entre pasos. Cancelar tiene que responder en ≤ 2 s.
- El callback se llama **desde el hilo de trabajo**: el core no sabe nada de tkinter.

#### `core/profile.py` (backend)

```python
@dataclass
class FileInfo:
    name: str; path: Path; size: int; modified: datetime; format: FileFormat

@dataclass
class ColumnProfile:                       # PLAN §2, más las marcas de §3.1
    name: str
    position: int                          # 0 = primera
    letter: str                            # "A", "B", …
    detected_type: ColumnType
    type_breakdown: dict[ValueKind, int]
    non_empty: int
    empty: int
    distinct: int                          # valores tal cual (sin normalizar)
    is_unique: bool                        # sin vacíos y distinct == filas → "posible clave"
    few_values: bool                       # distinct <= 50 y filas >= 2 × distinct → "pocos valores"
    summable: bool                         # entero, decimal o mixto con >= 90 % numérico
    top_values: list[tuple[str, int]]      # 10 más frecuentes
    sample_values: list[str]               # hasta 5 distintos, no vacíos
    min: str | None; max: str | None       # números y fechas, como texto para mostrar
    sum: float | None; mean: float | None  # números (solo lo convertible)
    min_len: int | None; max_len: int | None   # texto
    leading_zeros: int
    untrimmed: int
    case_variants: int                     # valores que solo difieren en mayúsculas/espacios de otro ("Norte"/"NORTE"/"Norte ")
    not_converted: int                     # celdas no vacías que no son del tipo detectado (mixto/numérico)
    notes: list[str]                       # observaciones en castellano para la columna "Observaciones"

@dataclass
class TableProfile:                        # PLAN §2
    file: FileInfo
    sheets: list[SheetSummary]             # [] en CSV
    sheet: str | None
    csv_options: CsvOptions | None
    header_row: int
    rows: int; columns: int
    empty_rows: int                        # filas totalmente vacías (del medio)
    duplicate_rows: int                    # copias de una fila anterior no vacía (convención de perfil.json)
    columns_profile: list[ColumnProfile]
    warnings: list[str]                    # castellano: sin nombre, repetidas, "no son números", espacios, variantes de mayúsculas

def profile_table(data: ReadResult, *, sheets: list[SheetSummary] | None = None,
                  on_column: Callable[[ColumnProfile], None] | None = None,
                  progress=None, cancel=None) -> TableProfile: ...
```

- Recibe lo ya leído (`read_table`): **no vuelve a leer el archivo**. `on_column` se llama a medida que termina cada columna (la UI muestra primero lo general y completa las columnas de a una).
- Tipos con `core/types.detect_column_type` (nunca su propia lógica).

#### Valores distintos, filtros, orden y páginas (`core/explore.py`, backend; ADR-024)

```python
@dataclass(frozen=True)
class DistinctOptions:
    sort: Literal["cantidad", "alfabetico"] = "cantidad"
    include_empty: bool = True             # fila "(vacío)"
    normalize: bool = False                # ignorar mayúsculas y espacios de los bordes
    search: str | None = None              # filtra valores que contienen el texto (sin importar mayúsculas)
    limit: int = 100_000                   # máximo de filas devueltas

@dataclass
class DistinctResult:
    columns: list[str]
    table: pd.DataFrame                    # columnas: las elegidas + "Cantidad" (int) + "Porcentaje" (float 0–100); vacío = "(vacío)"
    distinct_total: int                    # antes de search y limit
    rows_total: int
    truncated: bool
    variant_groups: list[list[str]]        # grupos que se juntan al normalizar (para el aviso), solo con normalize=False

def value_counts(df: pd.DataFrame, columns: list[str], options: DistinctOptions = DistinctOptions(),
                 *, progress=None, cancel=None) -> DistinctResult: ...
def rows_with_value(df: pd.DataFrame, columns: list[str], values: tuple[str, ...],
                    normalize: bool = False) -> pd.Index: ...   # posiciones de fila; "(vacío)" = ""

FilterOp = Literal["igual", "distinto", "contiene", "vacio", "no_vacio", "mayor", "menor"]

@dataclass(frozen=True)
class FilterCondition:
    column: str; op: FilterOp; value: str = ""

def filter_rows(df: pd.DataFrame, conditions: list[FilterCondition],
                column_types: dict[str, ColumnType]) -> pd.Index: ...   # Y entre condiciones
def sort_rows(df: pd.DataFrame, rows: pd.Index, column: str, ascending: bool,
              column_type: ColumnType) -> pd.Index: ...
def get_page(df: pd.DataFrame, rows: pd.Index, page: int, page_size: int = 1000) -> pd.DataFrame: ...
```

- `igual`/`distinto`/`contiene` comparan texto sin importar mayúsculas. `mayor`/`menor` comparan como número o fecha si el tipo de la columna lo es (con `core/types`), si no como texto; un valor de filtro inválido para el tipo → `ExcelToolError` con mensaje claro.
- Orden: números y fechas por valor (los no convertibles y vacíos al final), texto alfabético sin importar mayúsculas.
- `get_page` devuelve la página con una columna `"N.º fila"` = posición en el archivo (1 = primera fila de datos).

#### `export_profile` y preferencias (backend)

```python
def export_profile(profile: TableProfile, distincts: list[DistinctResult], path: Path) -> list[WriteResult]: ...
    # .xlsx con hojas "General", "Columnas" y una por cada DistinctResult (nombre de hoja = columnas). Usa write_sheets.
def export_distincts(result: DistinctResult, path: Path, format: WriteFormat) -> WriteResult: ...   # usa write_table

# core/preferences.py
@dataclass
class Preferences:
    last_folder: Path | None = None
    last_csv_options: CsvOptions | None = None
    last_destination: str | None = None    # se usa en la Fase 3
def load_preferences() -> Preferences: ...      # archivo JSON en platform_utils.user_data_dir(); si falta o está roto, defaults
def save_preferences(prefs: Preferences) -> None: ...
```

#### `ui/worker.py` y errores en la UI (frontend)

```python
def run_in_background(root: tk.Misc, func: Callable[[ProgressCallback, CancelToken], T], *,
                      title: str, on_success: Callable[[T], None],
                      on_error: Callable[[BaseException], None] | None = None,
                      cancellable: bool = True, modal: bool = True,
                      on_cancel: Callable[[], None] | None = None) -> CancelToken: ...
```

- Corre `func` en un `threading.Thread` (daemon). El progreso y el resultado viajan por `queue.Queue`; la UI los lee con `after(100, …)`. **Ningún widget se toca desde el hilo.**
- Muestra una ventana de progreso con mensaje, barra (indeterminada si `total=None`) y botón "Cancelar". Al cancelar (botón o X) la ventana se cierra enseguida, se descarta el resultado aunque `func` siga corriendo y se llama a `on_cancel` una sola vez, sin `on_success` ni `on_error` (ADR-027).
- Errores: `ExcelToolError` → se muestra su `mensaje`. Cualquier otra excepción → "Ocurrió un problema inesperado. El detalle quedó guardado en el registro." + traceback completo al log. Nunca un traceback en pantalla.

#### Alcance acordado (ADR-023)
- "Exportar lo filtrado" (pestaña Datos) y el aviso "demasiado grande para Google Sheets" pasan a la Fase 3 (necesitan `save_result` y `core/limits`). En la Fase 2 solo se exporta el informe y la lista de valores distintos.
- Memoria: presupuesto de la app ≤ 1,5 GB de pico con el CSV de 2M filas (PCs de 8 GB compartidas con el sistema y otros programas). Medido antes de empezar: leer + `detect_types` = ~1,0 GB de pico, 23 s. El perfil, los distintos y la UI no pueden sumar más de ~0,5 GB: cada tarea mide y reporta su pico.
- Progreso: el helper `ThrottledProgress(progress)` de `core/progress.py` (≤ 10 eventos/s; `force()` para cambio de columna) es el que usan las funciones largas del core (F2-01).
- Páginas: `get_page` numera páginas desde 0, igual que el proveedor de la grilla `ui/grilla.py` (F2-12).

#### Control de memoria (ADR-025, backend; propuesto por el líder, nivel 2)
```python
# platform_utils.py
def available_memory() -> int | None: ...   # bytes libres; Windows: GlobalMemoryStatusEx (ctypes); Linux: MemAvailable de /proc/meminfo; None si no se puede saber

# core/memory.py
@dataclass(frozen=True)
class MemoryCheck:
    needed: int            # bytes estimados para cargar y analizar
    available: int | None
    ok: bool               # True si available es None (no se puede saber → se deja cargar, con log)
    message: str           # castellano, para mostrar tal cual si not ok

def check_memory(path: Path, *, format: str | None = None, margin: float = 0.8) -> MemoryCheck: ...
```
- Estimación = tamaño del archivo × factor por formato (CSV / xlsx comprimido / xls / xlsb), medido en F2-05 con pyarrow y con un margen para análisis. `ok` si needed ≤ available × margin.
- Mensaje (ejemplo): "Este archivo es muy grande para la memoria libre de esta computadora: necesita alrededor de 2,3 GB y hay 1,1 GB libres. Cerrá otros programas y probá de nuevo." Montos con coma decimal.

#### Aclaraciones de la Fase 2 (líder, nivel 2, ADR-027)
- `ColumnProfile.min` y `max` vienen **como texto listo para mostrar**, con formato argentino: números con coma decimal y sin separador de miles, fechas dd/mm/aaaa (igual que la lectura de la Fase 1). `sum` y `mean` siguen siendo `float` (como dice el contrato) y la UI los formatea con `ui/formato.py` (corrección de F2-02).
- `TableProfile.sheets` trae todas las hojas cuando se le pasa `sheets=list_sheets(...)`, y `sheet` es el nombre de la hoja leída (None en CSV).
- `read_table` suma `progress` y `cancel` opcionales (palabra clave):
  - **CSV:** lee por partes y llama a `cancel.check()` entre partes, así responde en ≤ 2 s.
  - **Excel:** calamine lee la hoja de una vez y no se puede cortar; el progreso es indeterminado.
- Cancelación en la UI: al tocar Cancelar, el worker **cierra la ventana enseguida y descarta el resultado** aunque el hilo siga terminando (caso Excel). El hilo es daemon y su resultado se ignora.
- (F2-05) `ReadResult.df`: columnas con dtype `str` y almacenamiento pyarrow (`pd.StringDtype("pyarrow")`), ya no `object`. Los vacíos siguen siendo `""`. En memoria, el CSV de 2M ocupa 316 MB (antes 1.177 MB).
- (F2-05) Factores de memoria medidos para `check_memory`: base ~114 MB + tamaño del archivo × factor neto. Factores: CSV 4,4; xlsx 21; xls 16; xlsb sin medir (usar el de xlsx). `check_memory` usa factores con margen para el análisis: CSV 6, xlsx/xlsm/xlsb/ods 25, xls 20.
