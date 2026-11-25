# Plan de desarrollo — Excel Join Tool (tkinter)

App de escritorio para que personas del negocio, sin saber fórmulas de Excel, puedan:
1. **Explorar un archivo:** ver toda su información de un vistazo (hojas, cantidad de filas y columnas, tipo de dato de cada columna, vacíos, duplicados), ver los **valores distintos de cualquier columna** con su cantidad y recorrer todos los datos. Ver sección 3.1.
2. **Unir (join) varios archivos en uno:** se elige un **archivo principal** y se le van sumando columnas de **uno o más archivos secundarios** que comparten una columna clave (por ejemplo `ID cliente`). Para cada archivo se eligen las columnas que se quieren conservar. El resultado es un único archivo con todo junto. Es el caso de uso central de la app (ver sección 3.2).
3. **Ver detalle o resumen:** cuando un cliente aparece varias veces en un archivo secundario (por ejemplo, una fila por compra), el usuario elige entre:
   - **Detalle:** traer todas las filas, para ver qué artículos compró cada cliente.
   - **Resumen:** agrupar por las columnas que elija y calcular sobre las columnas que elija (suma, conteo, promedio, etc.). Por ejemplo, el monto total por cliente.
   - También se puede resumir el **resultado final** (por ejemplo, total por región).
4. **Guardar resultados grandes sin problemas:** si el resultado va a ser demasiado grande para abrirlo cómodo en **Microsoft Excel o Google Sheets**, la app lo detecta **antes de guardar** y recomienda dividirlo en varios archivos (ver sección 3.3).
5. **Concatenar** varios archivos (o varias hojas) en uno solo, uno debajo del otro.
6. **Dividir (split)** archivos muy grandes en partes más chicas para poder trabajarlos.

**Formatos:** CSV y planillas de Microsoft Excel, como entrada y como salida, y mezclables entre sí (por ejemplo, unir un CSV con dos .xlsx). Detalle en la sección 1.1.

**Plataformas:** se desarrolla en **Linux**. Tiene que correr con Python en **Linux y Windows**, y compilarse como **`.exe` en Windows** (y como ejecutable en Linux). Detalle en las secciones 1.3 y 1.4.

**Todo se hace en local:** no se usa GitHub Actions ni ningún CI. Los tests, la compilación y las pruebas se corren con los scripts del repo. **Al repo se sube solo código** (sección 1.5).

---

## 1. Decisiones técnicas

| Tema | Decisión | Por qué |
|---|---|---|
| Lenguaje | Python 3.12+ | Buen soporte de PyInstaller y pandas |
| UI | `tkinter` + `ttk` (sin dependencias extra de UI) | Viene con Python, empaqueta sin problemas |
| Datos | `pandas` | Joins, concat, groupby, value_counts ya resueltos y probados |
| Lectura Excel | `python-calamine` (motor `engine="calamine"` de pandas) | Lee .xlsx, .xlsm, .xlsb, .xls y .ods con una sola librería, y entre 5 y 10 veces más rápido que openpyxl |
| Escritura xlsx | `openpyxl` (o `xlsxwriter` si importa la velocidad) | Estándar |
| CSV | Lectura y escritura propias con autodetección (ver 1.1) | Los CSV que exporta Excel en Windows con configuración regional de Argentina vienen con `;`, coma decimal y cp1252 |
| Entorno | `venv` + `pyproject.toml` + `constraints.txt`, manejados con scripts (ver 1.3) | Mismo entorno y mismas versiones en Linux y Windows |
| Empaquetado | PyInstaller | Maduro y simple |
| Compilación | **Local**, con scripts. El `.exe` se compila **en una PC con Windows** y el binario de Linux en Linux | **PyInstaller no compila de un SO a otro.** Sin CI, por decisión del proyecto |
| Tests | `pytest` (lógica sin UI + tests de UI básicos), corridos en local | La lógica se puede testear sin abrir ventanas |

### 1.1 Formatos soportados

| Formato | Leer | Escribir | Notas |
|---|---|---|---|
| `.xlsx` | ✅ | ✅ | Formato principal de salida |
| `.xlsm` (con macros) | ✅ | — | Se leen los datos y se ignoran las macros |
| `.xlsb` (binario) | ✅ | — | Común en archivos grandes |
| `.xls` (Excel 97–2003) | ✅ | — | Al guardar se ofrece .xlsx |
| `.csv` / `.txt` | ✅ | ✅ | Ver reglas abajo |
| `.ods` (LibreOffice) | ✅ | — | Gratis con calamine, por si alguien lo usa en Linux |

**Reglas para CSV (es donde más se rompe todo):**
- **Separador:** detectarlo automáticamente (`;`, `,`, tabulación, `|`) con `csv.Sniffer` sobre las primeras líneas, y mostrarlo en pantalla para que se pueda cambiar.
- **Codificación:** probar `utf-8-sig` → `utf-8` → `cp1252`. Si no, el usuario ve "Ã±" en lugar de "ñ".
- **Leer todo como texto (`dtype=str`)** por defecto. Así no se pierden ceros a la izquierda (CUIT, códigos de producto, legajos) ni se convierten fechas por error. El tipo de cada columna se **detecta** aparte (`core/types.py`, ver 3.1) y los valores se convierten solo cuando hace falta (resúmenes).
- **Decimales:** detectar si los números usan coma decimal ("1.234,56") y convertirlos solo en las columnas que se van a sumar o promediar. Si una celda no se puede convertir, se avisa ("12 celdas de *Importe* no son números: ver cuáles") en lugar de fallar o ignorarla en silencio.
- **Al guardar CSV:** por defecto `;` como separador, `utf-8-sig` (con BOM) y coma decimal. Así el archivo **abre bien con doble clic en Excel de Windows** con configuración regional argentina. Opciones:
  - **"CSV estándar (coma)"**, avanzada.
  - **"CSV para Google Sheets"**: `,` + `utf-8`. Se sugiere sola cuando el destino elegido es Google Sheets (ver 3.3).
- Archivos sin encabezado o con filas basura arriba: la misma detección de fila de encabezado que en Excel.

**Excel:**
- Listar hojas y dejar elegir una (o "todas" en concatenar).
- Las celdas combinadas y las fórmulas se leen como valores. Avisarlo una vez.
- Al guardar: los límites de tamaño se controlan **siempre antes de escribir** (ver 3.3).
- **Archivo abierto en Excel** (Windows lo bloquea): detectar el `PermissionError` y mostrar "Cerrá el archivo en Excel y volvé a intentar".

### 1.2 Límites de destino

Estos números van en `core/limits.py` y además se pueden cambiar sin recompilar con un `limites.json` junto al ejecutable (los valores por defecto están en `src/excel_tool/limites_default.json`), porque los proveedores los cambian. Por ejemplo, **Google Sheets pasó de 10 a 20 millones de celdas por planilla en septiembre de 2026**.

| Destino | Límite duro (no abre o se corta) | Límite "cómodo" (recomendación) |
|---|---|---|
| **Microsoft Excel** | 1.048.576 filas y 16.384 columnas **por hoja** | ~300.000 filas o ~50 MB por archivo (arriba de eso se vuelve lento para filtrar, ordenar o tablas dinámicas) |
| **Google Sheets** | 20.000.000 de celdas **por planilla**, contando todas las hojas y celdas vacías. Importar archivos de más de ~100 MB suele fallar | ~5.000.000 de celdas o ~50 MB por archivo |
| **Solo guardar** (otro sistema, base de datos) | Sin límite en CSV | — |

Los valores "cómodos" son una recomendación inicial: se ajustan después de la prueba con usuarios (Fase 11).

### 1.3 Entorno de desarrollo, scripts y compilación local

Todo se maneja con scripts del repo. Son lo **primero que se construye (Fase 0)**, porque el resto del equipo los usa para correr la app y los tests.

| Script | Sistema | Qué hace |
|---|---|---|
| `scripts/dev.sh` | Linux | Crea/actualiza `.venv`, instala dependencias **solo si cambiaron** y abre la app |
| `scripts/dev.ps1` + `scripts/dev.bat` | Windows | Lo mismo. El `.bat` permite **doble clic** y evita el bloqueo de PowerShell (`-ExecutionPolicy Bypass`) |
| `scripts/build_windows.ps1` + `scripts/build_windows.bat` | Windows | Compila el `.exe` con PyInstaller, lo prueba y lo empaqueta |
| `scripts/build_linux.sh` | Linux | Compila el ejecutable de Linux, lo prueba y lo empaqueta |

**Opciones de `dev` (iguales en los dos sistemas):**

| Linux | Windows | Efecto |
|---|---|---|
| *(sin opciones)* | *(doble clic)* | Prepara el entorno y abre la app |
| `--setup` | `-Setup` | Solo prepara el entorno |
| `--tests` | `-Tests` | Corre `pytest` + `ruff` |
| `--recreate` | `-Recreate` | Borra `.venv` y lo crea de cero |
| `--no-dev` | `-NoDev` | Instala solo lo necesario para usar la app (sin pytest/ruff/pyinstaller) |
| `--lock` | `-Lock` | Regenera `constraints.txt` con las versiones instaladas |
| `-- <args>` | `-AppArgs <args>` | Pasa argumentos a la app (ej. `--self-test`) |

**Comportamiento obligatorio de `dev`:**
- Busca **Python ≥ 3.12**: en Linux `python3.13`, `python3.12`, `python3`; en Windows `py -3.13`, `py -3.12`, `py -3`, `python`. Descarta el alias falso de Microsoft Store. Variable `PYTHON` para forzar uno.
- Verifica **tkinter** antes de seguir. Si falta, explica cómo instalarlo: `sudo apt install python3-tk` (Ubuntu/Debian), `python3-tkinter` (Fedora); en Windows, reinstalar desde python.org marcando "tcl/tk and IDLE".
- En Linux verifica `venv`/`ensurepip` (en Ubuntu viene aparte: `python3-venv`).
- Reinstala dependencias **solo si cambió** `pyproject.toml` o `constraints.txt` (hash guardado en `.venv`).
- Instala con `pip install -e ".[dev,build]" -c constraints.txt`: mismas versiones en Linux y Windows.
- Mensajes en castellano, claros, y código de salida ≠ 0 si algo falla.

**Comportamiento obligatorio de `build_*`:**
1. Prepara el entorno (`dev --setup`).
2. Corre los tests (se puede saltar con `-SkipTests` / `--skip-tests`).
3. Limpia `build/` y `dist/`.
4. Compila con PyInstaller usando **`excel_tool.spec`** (uno solo para los dos sistemas). Por defecto en **carpeta** (arranca más rápido y la marcan menos los antivirus); opción `-OneFile` / `--onefile`.
5. **Prueba el ejecutable** con `--self-test`: importa todas las librerías empaquetadas, escribe y lee un .xlsx y un .csv con tildes y ceros a la izquierda, y abre y cierra una ventana. Si falta algo en el paquete, falla acá y no en la PC del usuario. En Linux sin pantalla usa `xvfb-run`.
6. Empaqueta en `dist/ExcelJoinTool-<versión>-windows.zip` (o `-linux.tar.gz`) y muestra tamaño y SHA256.

**Detalles de empaquetado que tienen que estar:**
- `packaging/launcher.py` como punto de entrada de PyInstaller (no puede arrancar con `python -m`), con `multiprocessing.freeze_support()`.
- `src/excel_tool/app_entry.py` con `main()`, que acepta `--version` y `--self-test`. Lo usan `python -m excel_tool`, el comando `excel-tool` y el ejecutable.
- `console=False` (sin consola negra). Como en ese modo **no hay `stdout`**, la app **nunca escribe a la consola**: loguea a archivo en `%LOCALAPPDATA%\ExcelJoinTool\logs` (Windows) o `~/.local/share/ExcelJoinTool/logs` (Linux).
- Archivos propios (ícono, `limites_default.json`) se leen con `importlib.resources.files("excel_tool")`. Así funcionan igual en desarrollo y empaquetados, sin trucos con `sys._MEIPASS`.
- `collect_all("python_calamine")` en el spec (trae una extensión compilada).
- Excluir lo que pandas arrastra y no se usa (matplotlib, IPython, scipy, pytest, Qt).
- Sin UPX (aumenta los falsos positivos de antivirus).
- Ícono `.ico` en Windows, `.png` en Linux.
- **Linux:** compilar en la distro **más vieja** que haya que soportar (glibc es compatible hacia adelante, no hacia atrás).

### 1.4 Compatibilidad Linux ↔ Windows

Como se desarrolla en Linux, **los problemas de Windows solo aparecen al probar en Windows**. Para reducirlos:

**Reglas de código** (también en `CLAUDE.md`):
- Rutas siempre con `pathlib.Path`. Nunca armar rutas concatenando strings con `/` o `\`.
- Abrir archivos **siempre con `encoding=` explícito**. En Windows el default es cp1252 y rompe las tildes.
- CSV con `newline=""` al escribir (si no, Windows agrega líneas en blanco).
- Archivos temporales con `tempfile`, nunca rutas fijas como `/tmp`.
- "Abrir carpeta": `os.startfile` en Windows, `xdg-open` en Linux, en una sola función.
- Carpetas de usuario (logs, preferencias) con una función que resuelva `%LOCALAPPDATA%` o `~/.local/share`.
- Fuentes de tkinter: `TkDefaultFont` y similares, no nombres de fuentes que existan solo en un sistema.
- No usar comandos de sistema (`subprocess` con `ls`, `cp`, etc.) desde la app.
- Nombres de archivo de salida: limpiar caracteres prohibidos en Windows (`<>:"/\|?*`) y nombres reservados (`CON`, `NUL`, etc.), por ejemplo cuando se divide por el valor de una columna.
- Tests que necesitan pantalla (tkinter): en Linux sin pantalla se corren con `xvfb-run`.

**Archivo `.gitattributes`** para fijar los fines de línea: `.sh` con LF (si no, falla con "bad interpreter"), `.ps1` y `.bat` con CRLF, el resto LF, y los Excel como binarios.

**Pruebas en Windows (manuales, a cargo del PM):** como no hay CI, se prueba en una PC o VM con Windows en estos momentos:
- **Al cerrar la Fase 0:** `scripts\dev.bat`, `scripts\dev.bat -Tests` y `scripts\build_windows.bat`.
- **Al cerrar las Fases 2, 5 y 8:** lo mismo, más el checklist de Windows de la sección 6.
- Opcional y recomendable: una VM con Windows en la misma máquina Linux, con la carpeta del repo compartida o clonada.

### 1.5 Qué se sube al repo

**Se sube:** código fuente, tests, scripts, `pyproject.toml`, `constraints.txt`, `excel_tool.spec`, `packaging/`, `.claude/` (configuración del equipo de agentes), `docs/`, `PLAN.md`, `CLAUDE.md`, `README.md`, `.gitattributes`, `.gitignore`, el ícono, y los **resultados esperados chicos** de los tests (`tests/fixtures/esperados/*.json`).

**No se sube (va en `.gitignore`):**
- `.venv/`, `build/`, `dist/` (ejecutables, .zip, .tar.gz), cachés (`__pycache__/`, `.pytest_cache/`, `.ruff_cache/`, `*.egg-info/`).
- **Archivos de prueba generados** (`tests/fixtures/generados/`): los crea `scripts/make_fixtures.py` en cada máquina. Incluye los CSV de millones de filas.
- Salidas de pruebas manuales (`salidas/`), logs, y `.claude/settings.local.json` (configuración personal).

**No hay CI ni Actions:** no existe `.github/workflows/`. Los ejecutables se comparten por fuera del repo (por ejemplo, en una carpeta compartida).

---

## 2. Arquitectura

```
excel-join-tool/
├── CLAUDE.md                 # instrucciones para Claude Code (reglas, comandos, equipo)
├── PLAN.md                   # este plan
├── README.md                 # cómo correr y compilar
├── pyproject.toml
├── constraints.txt           # versiones fijas de dependencias (mismas en Linux y Windows)
├── excel_tool.spec           # PyInstaller (Windows y Linux)
├── .gitattributes / .gitignore
├── .claude/                  # equipo de agentes (ver docs/equipo.md)
├── docs/
│   ├── equipo.md             # roles, escalamiento, formatos de despacho e informes
│   ├── decisiones.md         # registro de decisiones (ADR)
│   ├── contratos.md          # dataclasses y firmas core ↔ UI
│   └── qa/                   # informes de QA por fase
├── packaging/
│   └── launcher.py           # punto de entrada para PyInstaller
├── scripts/
│   ├── dev.sh / dev.ps1 / dev.bat                # entorno + correr app + tests
│   ├── build_linux.sh                            # compilar en Linux
│   ├── build_windows.ps1 / build_windows.bat     # compilar .exe en Windows
│   └── make_fixtures.py                          # genera archivos de prueba
├── src/excel_tool/
│   ├── __init__.py           # __version__
│   ├── __main__.py           # python -m excel_tool
│   ├── app_entry.py          # main(): --version, --self-test, logging, abrir UI
│   ├── limites_default.json  # límites por defecto (sección 1.2)
│   ├── platform_utils.py     # carpetas de usuario, abrir carpeta, nombres de archivo seguros
│   ├── core/                 # LÓGICA PURA (sin tkinter)
│   │   ├── io/
│   │   │   ├── __init__.py   # read_table(path, opciones) / write_table(df, path, opciones)
│   │   │   ├── excel.py      # xlsx/xlsm/xlsb/xls/ods: listar hojas, leer, escribir
│   │   │   ├── csv.py        # detección de separador, encoding, decimales; lectura por chunks
│   │   │   └── header.py     # detectar la fila de encabezado
│   │   ├── types.py          # detectar tipo de columna; convertir "1.234,56", fechas dd/mm/aaaa, etc.
│   │   ├── profile.py        # información del archivo y de cada columna; valores distintos
│   │   ├── join.py           # join multi-archivo (JoinPlan) + validaciones + diagnóstico
│   │   ├── aggregate.py      # resúmenes: agrupar por columnas elegidas + operaciones por columna
│   │   ├── limits.py         # límites de Excel / Google Sheets; estimar tamaño; recomendar división
│   │   ├── output.py         # guardar un resultado: un archivo, varias hojas o varias partes
│   │   ├── concat.py         # concatenar archivos/hojas
│   │   ├── split.py          # dividir por filas, por valor de columna o por filas sin cortar grupos
│   │   └── normalize.py      # limpieza de claves (espacios, mayúsculas, tipos)
│   ├── ui/
│   │   ├── app.py            # ventana principal con menú de operaciones
│   │   ├── widgets/          # selector de archivo, columnas con checkbox, grilla de datos paginada,
│   │   │                     # panel de valores distintos, editor de resumen, paso "Guardar"
│   │   ├── explorer.py       # ventana "Ver información de un archivo"
│   │   ├── join_wizard.py
│   │   ├── concat_wizard.py
│   │   ├── split_wizard.py
│   │   └── worker.py         # hilos + cola para operaciones largas
│   └── assets/               # ícono (.ico y .png)
└── tests/
    ├── core/                 # unitarios del core (backend)
    ├── ui/                   # tests básicos de UI (frontend)
    ├── integracion/          # de punta a punta (QA)
    ├── conftest.py
    └── fixtures/
        ├── esperados/        # resultados correctos chicos (se suben)
        └── generados/        # creados por make_fixtures.py (NO se suben)
```

**Regla de oro:** `core/` no importa nada de `ui/`. La UI arma un objeto de configuración (dataclass) y se lo pasa al core. Así la lógica se testea sola y Claude Code puede trabajar cada parte por separado.

**Una sola puerta de entrada para leer:** todo pasa por `read_table()`, que devuelve un DataFrame más metadatos (formato, hoja, separador y encoding detectados). Joins, concat, split y el explorador no saben si el origen era CSV o Excel.

**Una sola puerta de salida para guardar:** todo resultado (join, concat, export del explorador, split) se guarda con `output.save_result()`. Esa función controla los límites (`limits.py`) y, si corresponde, divide con `split.py`. Así la recomendación de dividir funciona igual en toda la app.

**Un solo lugar para los tipos:** `core/types.py` decide si una columna es número, fecha, texto, etc., y convierte los valores. Lo usan el explorador (para mostrar el tipo), los resúmenes (para sumar) y el join (para avisar de claves número vs. texto). Así los tres dicen lo mismo.

**Un solo lugar para lo que depende del sistema:** `platform_utils.py` (carpetas de usuario, abrir carpeta, nombres de archivo válidos en Windows). Ningún otro módulo pregunta `sys.platform`.

**Modelo del explorador** (en `core/profile.py`):

```python
ColumnType = Literal["texto", "entero", "decimal", "fecha", "fecha_hora",
                     "si_no", "mixto", "vacia"]

@dataclass
class ColumnProfile:
    name: str
    position: int                   # letra/número de columna (A, B, C...)
    detected_type: ColumnType
    type_breakdown: dict[ColumnType, int]   # ej. {"entero": 9950, "texto": 50} → "mixto"
    non_empty: int
    empty: int                      # vacíos (celda vacía o solo espacios)
    distinct: int
    is_unique: bool                 # distintos == filas no vacías y sin vacíos → "posible clave"
    top_values: list[tuple[str, int]]  # los 10 más frecuentes
    sample_values: list[str]        # algunos ejemplos
    min: str | None; max: str | None   # números y fechas
    sum: float | None; mean: float | None  # números
    min_len: int | None; max_len: int | None  # texto
    leading_zeros: int              # valores con ceros a la izquierda ("00123")
    untrimmed: int                  # valores con espacios al inicio o al final

@dataclass
class TableProfile:
    file_name: str; file_size: int; format: str
    sheets: list[SheetSummary]      # para Excel: nombre, filas y columnas de cada hoja
    sheet: str | None               # hoja analizada
    csv_options: CsvOptions | None  # separador y encoding detectados
    rows: int; columns: int
    empty_rows: int                 # filas totalmente vacías
    duplicate_rows: int             # filas exactamente repetidas
    columns_profile: list[ColumnProfile]
    warnings: list[str]             # ej. "La columna *Importe* tiene 12 valores que no son números"
```

- `profile_table(path, read_options, progress)` arma el `TableProfile`. En archivos grandes trabaja por partes (chunks) e informa el progreso.
- `value_counts(df, column, options)` devuelve los **valores distintos** de una columna con su cantidad y porcentaje. Opciones: ordenar por cantidad o alfabéticamente, incluir vacíos, ignorar mayúsculas/espacios, filtrar por texto. Se calcula cuando el usuario la pide, no antes.
- `rows_with_value(df, column, value)` devuelve las filas que tienen ese valor (para el doble clic en un valor distinto).
- `export_profile(profile, distincts, path)` guarda el informe en .xlsx.

**Modelo del join multi-archivo** (en `core/join.py` y `core/aggregate.py`):

```python
AggFunc = Literal["sum", "count", "count_distinct", "mean", "min", "max",
                  "first", "last", "join_text"]   # join_text = "A, B, C" sin repetidos

@dataclass
class AggregationSpec:                    # un resumen
    group_by: list[str]                   # columnas por las que se agrupa (elegidas por el usuario)
    aggregations: list[tuple[str, AggFunc, str]]  # (columna, operación, nombre de la columna resultado)
                                          # ej. ("Importe", "sum", "Total comprado")

@dataclass
class SourceSpec:                         # un archivo ya leído
    name: str                             # nombre visible, ej. "ventas.csv"
    path: Path
    read_options: ReadOptions             # hoja, separador, encoding, fila de encabezado
    columns: list[str]                    # columnas a conservar (además de la clave)

@dataclass
class SecondarySpec(SourceSpec):
    key_pairs: list[tuple[str, str]]      # (col. en principal, col. en este archivo); por defecto se emparejan por nombre igual
    how: Literal["left", "inner"]         # "completar" (default) o "solo los que están en ambos"
    mode: Literal["detail", "summary", "first"]
    summary: AggregationSpec | None       # si mode == "summary"; group_by incluye siempre la clave
    prefix: str | None                    # prefijo opcional para sus columnas

@dataclass
class JoinPlan:
    principal: SourceSpec
    key_options: KeyOptions               # recortar espacios, ignorar mayúsculas, tratar como texto
    secondaries: list[SecondarySpec]      # 1..N, se aplican en orden
    output_column_order: list[str]        # orden final elegido por el usuario
    final_summary: AggregationSpec | None # resumen opcional del resultado final (ej. por Región)
    output: OutputSpec | None             # cómo se guardó (destino, formato, división) → va en la receta
```

- En modo `summary`, el resumen del secundario se calcula **antes** de unirlo al principal. Se agrupa por la clave más las columnas extra que elija el usuario (por ejemplo `ID cliente` + `Artículo` da "total por cliente y artículo").
- `aggregate.summarize(df, spec)` es la misma función para el resumen de un secundario, para el resumen final y para la futura operación "Resumir un archivo".
- `diagnose_join(plan)` devuelve un diagnóstico **por cada secundario** (coincidencias, sin pareja, claves repetidas, filas resultantes estimadas) y uno total. Avisa si hay **más de un secundario en modo detalle** (ver 3.2, paso 4). Incluye la **estimación de tamaño del resultado** (`limits.estimate`, ver abajo), calculada **sin ejecutar el join**: con las cantidades de filas por clave de cada archivo se sabe cuántas filas va a tener.
- `run_join(plan, progress)` resume cada secundario si corresponde, los une uno por uno sobre el principal, aplica el resumen final si lo hay e informa el progreso por paso.
- El `JoinPlan` se puede **guardar y volver a cargar como .json** ("receta"), para repetir el mismo cruce cada mes con archivos nuevos. La receta incluye cómo se guardó y dividió el resultado. Las rutas se guardan de forma que la receta sirva en Linux y en Windows.

**Modelo de salida y límites** (en `core/limits.py`, `core/output.py` y `core/split.py`):

```python
Destination = Literal["excel", "google_sheets", "solo_guardar"]

@dataclass
class SizeEstimate:
    rows: int
    columns: int
    cells: int                    # filas × columnas
    approx_mb_xlsx: float         # estimado por promedio de bytes por celda (muestra de filas)
    approx_mb_csv: float

@dataclass
class SizeCheck:
    level: Literal["ok", "recomendado", "obligatorio"]   # ok / conviene dividir / no entra
    reasons: list[str]            # ej. "Excel acepta hasta 1.048.576 filas por hoja; el resultado tiene 2.300.000"
    suggested_parts: int          # cantidad mínima de partes para quedar dentro del límite cómodo
    suggested_rows_per_part: int
    suggested_columns: list[ColumnSuggestion]  # columnas buenas para dividir por valor (pocos valores,
                                               # y con la parte más grande dentro del límite)

@dataclass
class SplitSpec:
    mode: Literal["rows", "rows_keep_groups", "by_value", "sheets"]
    rows_per_part: int | None
    group_column: str | None      # para "rows_keep_groups": no cortar un mismo cliente entre dos partes
    value_columns: list[str] | None  # para "by_value": una parte por Región, por Mes, etc.
    name_pattern: str             # ej. "{nombre}_parte_{n:02d}_de_{total:02d}" o "{nombre}_{valor}"

@dataclass
class OutputSpec:
    path: Path                    # archivo o carpeta destino
    format: Literal["xlsx", "csv_excel", "csv_google", "csv_standard"]
    destination: Destination
    split: SplitSpec | None
    write_index: bool             # archivo "_indice" con la lista de partes
```

- `limits.estimate(df_or_diagnosis, sample)` → `SizeEstimate`.
- `limits.check(estimate, destination, format)` → `SizeCheck`, con los límites de la tabla 1.2.
- `limits.suggest_split_columns(df, destination)` usa `value_counts` para proponer columnas con pocos valores distintos (ej. `Región`: 5 partes, la más grande de 420.000 filas).
- `split.split_dataframe(df, SplitSpec)` → iterador de partes `(nombre, df_parte)`. Esto permite **escribir cada parte a medida que se genera**, sin armar primero un archivo gigante. El mismo código lo usa el asistente de Split (que trabaja sobre un archivo ya existente, por chunks).
- `output.save_result(df, OutputSpec, progress)` escribe un archivo, varias hojas o varias partes, y verifica al final que **la suma de filas de las partes sea igual al total**. Los nombres de las partes pasan por `platform_utils` para que sean válidos en Windows.

**Operaciones largas:** se corren en un `threading.Thread`, mandan el progreso por `queue.Queue` y la UI lo lee con `after()`. Nunca se toca tkinter desde el hilo secundario.

---

## 3. Diseño de UX (pensado para usuarios no técnicos)

- Pantalla de inicio con **4 botones grandes**: "Ver información de un archivo", "Unir archivos por una columna", "Juntar archivos uno debajo del otro", "Dividir un archivo grande".
- Cada operación (salvo el explorador) es un **asistente paso a paso** con "Atrás / Siguiente".
- **Selector de archivo único** para todos los formatos (filtro "Planillas y CSV"). Según el tipo de archivo, debajo aparece "Hoja: [▼]" (Excel) o "Separador: [;] Codificación: [auto]" (CSV), con un preview de 10 filas para confirmar que se leyó bien.
- **En todos los asistentes**, al lado de cada archivo elegido hay un botón **"🔍 Ver información"** que abre el explorador sobre ese archivo. Sirve para revisar qué columna usar como clave o cuál sumar antes de seguir.
- **Paso "Guardar" común a todas las operaciones** (widget reutilizable), con el control de tamaño y la recomendación de dividir de la sección 3.3.
- Lenguaje de negocio, no técnico:
  - "left join" → **"Mantener todas las filas del principal y completar cuando haya datos"** (opción por defecto)
  - "inner join" → **"Quedarme solo con las filas que aparecen también en este archivo"**
  - Anti-join (fase posterior): **"Filas del principal que NO están en este archivo"** (muy pedido para conciliaciones)
  - "group by" → **"Agrupar por"**; "aggregate" → **"Calcular"**; "distinct" → **"Valores distintos"**; "null" → **"Vacío"**; "split" → **"Dividir en partes"**
- Mostrar siempre un **preview** de las primeras 50 filas antes de guardar.
- **Avisos que evitan errores típicos:**
  - Tipos distintos: "En un archivo el código es número (123) y en el otro texto ("00123")". Se ofrece normalizar.
  - Resumen por archivo: "*ventas.csv*: 9.820 de 10.000 clientes encontraron datos; 180 quedaron vacíos (ver cuáles)".
  - Columna a sumar con valores que no son números.
  - Resultado demasiado grande para el destino elegido (ver 3.3).
- Nombres de columna repetidos entre archivos: se agrega como sufijo el nombre del archivo (por ej., `Teléfono (crm.xlsx)`), no `_x`/`_y`.
- Recordar la última carpeta usada, las últimas opciones de CSV y el **último destino elegido** (Excel o Google Sheets).
- Mensajes de error en castellano y sin traceback. El detalle técnico va a un archivo de log.

### 3.1 Explorador: "Ver información de un archivo"

Una ventana con el archivo elegido y **4 pestañas**. Se calcula en segundo plano con barra de progreso. En archivos grandes, primero aparecen los datos generales y después se completa columna por columna.

**Pestaña 1 — General**
- Nombre, formato, tamaño en disco, fecha de modificación.
- **Excel:** lista de **todas las hojas** con la cantidad de filas y columnas de cada una. Un clic en una hoja cambia el análisis a esa hoja.
- **CSV:** separador y codificación detectados (se pueden cambiar y se vuelve a analizar).
- **Cantidad de filas** y **cantidad de columnas**.
- Filas totalmente vacías y **filas repetidas** (idénticas en todas las columnas), con "ver cuáles".
- Fila donde se detectó el encabezado.
- **Avisos** en lenguaje simple: "La columna *Importe* tiene 12 valores que no son números", "La columna *Código* tiene valores con espacios al final", "Hay 2 columnas sin nombre", "Hay 2 columnas con el mismo nombre", **"Este archivo es demasiado grande para abrirlo en Google Sheets"**.

**Pestaña 2 — Columnas** (una fila por columna, tabla ordenable)

| Columna | Tipo | Vacíos | Distintos | Ejemplos | Mín. | Máx. | Observaciones |
|---|---|---|---|---|---|---|---|
| ID cliente | Número entero | 0 | 10.000 | 1001, 1002, 1003 | 1001 | 11000 | 🔑 Posible clave (no se repite) |
| Código | Texto | 0 | 9.998 | 00123, 00124 | | | ⚠️ 2 repetidos · conserva ceros a la izquierda |
| Fecha alta | Fecha | 15 (0,2%) | 1.230 | 05/03/2024 | 01/01/2019 | 20/11/2026 | |
| Importe | Mixto: 99,5% decimal | 3 | 8.412 | 1.234,56 | 0,50 | 98.000,00 | ⚠️ 50 no son números (ver cuáles) |
| Región | Texto | 0 | 5 | Norte, Sur | | | Pocos valores: ideal para agrupar o dividir |

- **Tipos posibles** (en castellano): Texto, Número entero, Número decimal, Fecha, Fecha y hora, Sí/No, **Mixto** (con el % de cada tipo y "ver cuáles"), Vacía.
- El tipo se calcula **por columna**, mirando todos sus valores, no solo los primeros. Una columna como `00123` se marca como Texto (código), no como número, para no sugerir que se pueden perder los ceros.
- Para números: además suma y promedio (en el panel de detalle al hacer clic en la columna).
- Para texto: largo mínimo y máximo.
- Marcas útiles para los otros asistentes: **🔑 Posible clave** (sin vacíos ni repetidos), **Pocos valores** (sirve para agrupar o dividir), **Se puede sumar**.
- Doble clic en una columna → abre la pestaña 3 con esa columna.

**Pestaña 3 — Valores distintos (distinct por columna)**
- Desplegable para elegir la columna (o varias a la vez, para ver **combinaciones distintas**, por ejemplo `Región` + `Vendedor`).
- Tabla: **Valor | Cantidad de filas | %**, con una barra visual del %. Incluye la fila "(vacío)" si los hay.
- Ordenar por cantidad (de mayor a menor, por defecto) o alfabéticamente.
- **Buscador** para filtrar valores.
- Opciones: "ignorar mayúsculas y espacios". Esta opción sirve para ver si "Norte", "NORTE" y "Norte " son el mismo valor escrito distinto. Si es así, se avisa.
- Total arriba: "5 valores distintos en 10.000 filas".
- **Doble clic en un valor** → pestaña 4 filtrada por ese valor ("ver las filas de Región = Norte").
- **Exportar** la lista de distintos a Excel/CSV o **copiar** al portapapeles (para pegarla en un mail o en otro Excel).
- Si la columna tiene muchísimos distintos (más de 100.000), se muestran los primeros y se ofrece exportar la lista completa.

**Pestaña 4 — Datos**
- Grilla con **todas las filas**, paginada (de a 1.000), con número de fila. No se cargan todas en pantalla a la vez, así no se cuelga con archivos grandes.
- **Filtro simple:** `[columna ▼] [es igual a / contiene / está vacío / mayor que / menor que ▼] [valor]`, con "+ Agregar filtro". Muestra "Mostrando 1.230 de 10.000 filas".
- Ordenar haciendo clic en el encabezado de la columna.
- Ocultar o mostrar columnas.
- **Exportar lo filtrado** a .xlsx o .csv (pasa por el paso "Guardar" de la sección 3.3). Es una forma simple de sacar "todos los clientes de la región Norte" sin fórmulas.

**Botón "Exportar informe"**: guarda un .xlsx con una hoja "General", una hoja "Columnas" y una hoja por cada columna de la que se hayan pedido los valores distintos.

### 3.2 Asistente de Join (principal + N archivos)

Ejemplo guía: `clientes.xlsx` (principal) + `ventas.csv` (una fila por artículo comprado: `ID cliente`, `Fecha`, `Artículo`, `Cantidad`, `Importe`) + `cobranzas.xlsx`.

1. **Archivo principal:** elegirlo → hoja o separador → confirmar la fila de encabezados con el preview. Se explica: "Este es el archivo base".
2. **Columna clave:** elegir la columna del principal por la que se va a unir (ej. `ID cliente`). Se puede elegir más de una (ej. `ID cliente` + `Sucursal`). En el desplegable se marcan con 🔑 las columnas que el explorador detecta como posible clave.
3. **Archivos a sumar:** botón "+ Agregar archivo" (se repite las veces que haga falta). Se ven en una lista que se puede reordenar o quitar. Para cada uno:
   - Hoja o separador, con preview.
   - **Clave:** se empareja sola si hay una columna con el mismo nombre (sin importar mayúsculas ni espacios). Si no la hay, se elige de un desplegable (ej. `ID cliente` ↔ `Nro Cliente`).
   - **Tipo de unión:** "completar" (por defecto) o "solo los que están en ambos".
   - Indicador ✅/⚠️ con el diagnóstico rápido (% de coincidencia, si hay claves repetidas).

4. **¿Detalle o resumen?** Se pregunta por cada archivo secundario donde la clave se repite, con un ejemplo de sus propios datos ("El cliente 1045 aparece 12 veces en *ventas.csv*"). **No hay opción por defecto**: los dos usos son igual de comunes, así que el usuario elige siempre.

   - **🧾 Ver el detalle (una fila por cada registro)**
     *"Para saber qué compró cada cliente".* Se traen todas las filas: el cliente 1045 aparece 12 veces, una por artículo, con sus datos del principal repetidos en cada fila. Se muestra cuántas filas va a tener el resultado y, **si va a ser grande, se adelanta el aviso**: "El resultado va a tener ~2.300.000 filas: no entra en una sola hoja de Excel. Al guardar te vamos a proponer dividirlo".

   - **📊 Ver un resumen**
     *"Para saber cuánto compró cada cliente".* Se abre el **editor de resumen** (un widget reutilizable):
     - **Agrupar por:** la clave ya viene marcada y no se puede quitar. Se pueden sumar otras columnas del archivo (checkboxes). Por ejemplo:
       - solo `ID cliente` → una fila por cliente
       - `ID cliente` + `Artículo` → una fila por cliente y artículo
       - `ID cliente` + `Mes` → una fila por cliente y mes (si hay una columna de fecha, ofrecer "agrupar por mes/año de *Fecha*")
     - **Calcular:** una tabla donde cada fila es `[columna ▼] [operación ▼] [nombre del resultado]`, con "+ Agregar cálculo". Operaciones: **Sumar, Contar filas, Contar distintos, Promedio, Mínimo, Máximo, Primero, Último, Listar valores** (ej. "Lápiz, Cuaderno, Goma").
       - Ejemplo: `Importe → Sumar → "Total comprado"`, `Artículo → Contar distintos → "Artículos distintos"`, `Fecha → Máximo → "Última compra"`.
       - El nombre del resultado se completa solo ("Suma de Importe") y se puede editar.
       - El desplegable de columnas muestra el tipo detectado. Sumar y Promedio solo se ofrecen para columnas numéricas o mixtas (en ese caso, con aviso).
     - Preview del resumen de ese archivo antes de unirlo.

   - **1️⃣ Traer solo la primera fila** (opción secundaria, en "Más opciones").

   ⚠️ **Más de un archivo en modo detalle:** si `ventas.csv` y `cobranzas.xlsx` están los dos en "detalle", las filas se multiplican entre sí: 12 compras × 5 pagos = 60 filas para ese cliente, con datos que no tienen relación entre sí. La app lo detecta y explica:
   > "Si traés el detalle de dos archivos a la vez, las filas se combinan todas con todas y los totales quedan inflados. Te recomendamos ver el detalle de uno solo y el resumen de los otros."

   Ofrece "Cambiar *cobranzas.xlsx* a resumen" con un clic. Si el usuario insiste, se permite, pero mostrando la cantidad de filas resultante.

5. **Opciones de limpieza de la clave** (para todos los archivos): recortar espacios, ignorar mayúsculas y tratar como texto, activadas por defecto.

6. **Elegir columnas:** un panel con una pestaña o sección por archivo (principal + cada secundario), con checkboxes, "todas/ninguna" y buscador. En los archivos resumidos aparecen las columnas del resumen ("Total comprado", etc.). La clave aparece una sola vez en el resultado. Abajo se ve la **lista final de columnas, que se puede reordenar** con ↑/↓ o arrastrando. Un **indicador de tamaño** se actualiza a medida que se marcan columnas: "~2.300.000 filas × 18 columnas = 41 millones de celdas", con un semáforo para el destino elegido. Así el usuario ve que quitar columnas también achica el resultado.

7. **(Opcional) Resumir el resultado final:** "¿Querés agrupar el resultado?". Usa el mismo editor de resumen, pero con **todas las columnas del resultado**, de cualquier archivo. Ejemplos:
   - Agrupar por `Región` (de clientes) → Sumar `Importe` (de ventas): total vendido por región.
   - Agrupar por `Vendedor` + `Artículo` → Sumar `Cantidad`.

8. **Preview + diagnóstico** (total y por archivo) → botón **"🔍 Ver información del resultado"** (abre el explorador sobre el resultado antes de guardarlo) → **paso Guardar** (sección 3.3). Opcionales:
   - En .xlsx, guardar **detalle y resumen en hojas distintas del mismo archivo** (ej. hoja "Detalle" + hoja "Resumen por cliente"), porque muchas veces se quieren las dos. Si el detalle es grande, se recomienda guardar **el resumen en un archivo y el detalle dividido en partes**.
   - Hojas extra con "Sin coincidencia en *ventas.csv*" para revisar.
   - **"Guardar esta receta"** (.json) para repetir el mismo cruce con archivos nuevos.

**Control de filas:** si ningún secundario está en modo detalle y no hay resumen final, el resultado tiene **exactamente las mismas filas que el archivo principal**. La app lo verifica y avisa si no se cumple. En modo detalle, se muestra antes de guardar cuántas filas va a tener el resultado y por qué.

### 3.3 Paso "Guardar" y recomendación de dividir

Es el último paso de **todas** las operaciones (unir, concatenar, exportar desde el explorador). Es un solo widget reutilizable que llama a `output.save_result`.

**1. ¿Dónde vas a abrir el archivo?** (se recuerda la última elección)
- 🟩 **Microsoft Excel**
- 🟦 **Google Sheets**
- 📁 **Solo guardarlo** (para otro sistema; no hay límites)

**2. Control de tamaño**, calculado con `limits.check`. Pasa antes de escribir nada:

| Semáforo | Cuándo | Qué hace la app |
|---|---|---|
| 🟢 OK | Dentro del límite cómodo | Guarda normal en un solo archivo |
| 🟡 Conviene dividir | Pasa el límite cómodo pero no el duro | **Recomienda** dividir (opción preseleccionada), pero deja guardar en un solo archivo |
| 🔴 No entra | Pasa el límite duro (ej. más de 1.048.576 filas para Excel, más de 20 M de celdas para Google Sheets) | **Obliga** a elegir: dividir, achicar o guardar como CSV en "Solo guardarlo" |

Mensaje de ejemplo (🔴, destino Excel):
> **El resultado es demasiado grande para abrirlo en Excel.**
> Tiene 2.300.000 filas y 18 columnas. Excel acepta hasta 1.048.576 filas por hoja, y con más de ~300.000 se pone lento.
> **Te recomendamos dividirlo en 8 archivos de ~290.000 filas cada uno.**

**3. Opciones para dividir** (con una vista previa de las partes que van a salir: nombre, filas y tamaño estimado de cada una):
- **Por cantidad de filas** (sugerida): la app propone el número de partes. Si el resultado está en modo detalle, viene activada la opción **"No cortar un mismo cliente entre dos archivos"** (se agrupa por la clave): todas las compras del cliente 1045 quedan en la misma parte, aunque las partes no tengan exactamente la misma cantidad de filas.
- **Por valor de una columna:** la app **sugiere columnas** con pocos valores distintos donde todas las partes entran en el límite (ej. "`Región`: 5 archivos, el más grande de 420.000 filas ✅", "`Mes`: 12 archivos, el más grande de 230.000 filas ✅"). Si alguna parte igual queda demasiado grande (ej. Región Norte con 1,5 M), se ofrece **dividir esa parte además por cantidad de filas**.
- **En varias hojas del mismo archivo** (solo Excel y solo si el problema es el límite de filas por hoja, no el tamaño total): una hoja por parte. Para Google Sheets no se ofrece, porque el límite es por planilla y varias hojas no lo resuelven.

**4. Otras salidas** que la app también sugiere:
- **Achicar el resultado:** "Volver a elegir columnas" o "Cambiar *ventas.csv* a resumen", con el tamaño que quedaría ("Con resumen por cliente: 10.000 filas 🟢").
- **Guardar como un solo CSV** ("Solo guardarlo"), aclarando que no se va a poder abrir completo en Excel ni en Google Sheets.

**5. Cómo se guardan las partes**
- En una **carpeta nueva** con el nombre del resultado: `resultado_ventas/resultado_ventas_parte_01_de_08.xlsx`, o `resultado_ventas_Norte.xlsx` si se divide por valor (con el nombre limpiado para que sea válido en Windows).
- **Cada parte lleva el encabezado** y, en Excel, la primera fila congelada y el filtro activado.
- Archivo **`_indice.xlsx`** con la lista de partes: nombre, filas, desde/hasta, y el valor si se dividió por columna. Sirve para encontrar, por ejemplo, "en qué archivo está el cliente 1045".
- Las partes se escriben **a medida que se generan**, sin crear antes un archivo gigante.
- Formato según destino: Excel → .xlsx (o "CSV para Excel"); Google Sheets → .xlsx o "CSV para Google Sheets" (`,` + utf-8), y se avisa si alguna parte supera ~100 MB, que es lo que suele fallar al importar.
- Al terminar: "Se guardaron 8 archivos con 2.300.000 filas en total (coincide con el resultado ✅)" + botón **"Abrir carpeta"**.

**6. Receta:** la división elegida (modo, columna, filas por parte) se guarda en la receta, así el mes siguiente sale igual.

### 3.4 Asistente de Concatenar
- Agregar N archivos de cualquier formato mezclado (y elegir hojas, o "todas las hojas").
- Las columnas se alinean por nombre (ignorando mayúsculas y espacios). Se muestra cuáles faltan en cada archivo.
- Opción: agregar una columna "Origen" con el nombre del archivo/hoja.
- Opción: eliminar filas duplicadas.
- Guardar con el paso común de la sección 3.3 (juntar muchos archivos es una de las formas más comunes de pasarse del límite).
- Al terminar, opción **"Usar este resultado como archivo en un cruce"**: por ejemplo, juntar 12 CSV mensuales de ventas y después cruzarlos con clientes.

### 3.5 Asistente de Split
- Para dividir un archivo **que ya existe** (el paso Guardar de 3.3 divide resultados nuevos). Usa el mismo `split.py`.
- Primero muestra la información del archivo y el semáforo para Excel y Google Sheets, con el número de partes sugerido.
- Por **cantidad de filas** (por ej. 100.000 por archivo), por **cantidad de filas sin cortar grupos**, o por **valor de una columna** (un archivo por sucursal, por mes, etc.). Al elegir la columna, se muestran sus valores distintos con la cantidad de filas de cada uno, para saber cuántos archivos van a salir.
- Salida: varios archivos (.xlsx o .csv) o varias hojas del mismo archivo, con índice.
- Repetir el encabezado en cada parte.
- Para archivos enormes: leer en modo streaming (CSV en chunks, Excel fila a fila) para no cargar todo en memoria. CSV → CSV es el caso más común y debe andar con archivos de varios GB.

---

## 4. Fases de trabajo

Cada fase sigue el flujo de `docs/equipo.md`: el líder técnico arma el plan de la fase, el PM lo aprueba, backend/frontend/QA ejecutan, QA verifica, el líder commitea y el PM prueba a mano.

> **Por qué el explorador va antes que el join:** es la función más simple de punta a punta. Valida la lectura de todos los formatos y obliga a construir la base de la UI (ventana principal, worker con progreso, grilla de datos). Además, deja una app ya útil para mostrar a los usuarios temprano.

### Fase 0 — Setup, scripts y compilación (1½–2 días)

Objetivo: que cualquiera pueda clonar el repo y, con un solo comando, correr la app y los tests en Linux y en Windows, y compilar el `.exe` en Windows.

- **Backend:**
  - `pyproject.toml` (dependencias + extras `dev` y `build`, ruff, pytest), `constraints.txt` (con `dev --lock`).
  - Esqueleto de `src/excel_tool/`: `__init__.py` (`__version__`), `__main__.py`, `app_entry.py` (`--version`, `--self-test`, logging a archivo), `platform_utils.py`, `limites_default.json`.
  - Scripts de la sección 1.3: `dev.sh`, `dev.ps1`, `dev.bat`, `build_linux.sh`, `build_windows.ps1`, `build_windows.bat`.
  - `excel_tool.spec` y `packaging/launcher.py` según 1.3.
  - `README.md` con cómo correr, testear y compilar en cada sistema.
  - Tests en `tests/core/`: `--version`, `--self-test`, `platform_utils`.
- **Frontend:** `ui/app.py` mínimo (ventana con título y versión) + `tests/ui/` con un test que la abre y la cierra.
- **QA:** `scripts/make_fixtures.py`, que genera en `tests/fixtures/generados/`:
  - .xlsx con varias hojas, .xls, .xlsb (si no se pueden generar, uno chico en `tests/fixtures/esperados/` como excepción, aprobada por el PM)
  - CSV con `;` + cp1252 + coma decimal (como el que exporta Excel en Windows en Argentina)
  - CSV con `,` + utf-8, CSV con BOM, CSV con tabulaciones
  - **Archivo "para explorar"** con una columna de cada tipo: entero, decimal con coma, fecha dd/mm/aaaa, fecha y hora, Sí/No, texto, código con ceros a la izquierda, una columna mixta, una vacía, una sin nombre, dos con el mismo nombre, valores "Norte"/"NORTE"/"Norte ", filas vacías y filas repetidas. Con `tests/fixtures/esperados/perfil.json`.
  - **Escenario de 3 archivos:**
    - `clientes.xlsx`: 1 fila por cliente, con `Región` y `Vendedor`
    - `ventas.csv`: varias filas por cliente con `Fecha`, `Artículo`, `Cantidad`, `Importe`; `;` y coma decimal; algunos importes inválidos
    - `cobranzas.xlsx`: varias filas por cliente; columna clave con otro nombre (`Nro Cliente`), con ceros a la izquierda y espacios
    - Algunos clientes faltan en cada secundario
    - Resultados esperados chicos en `tests/fixtures/esperados/*.json`, calculados a mano
  - **Escenario grande:** `clientes_grande.csv` (200k clientes) + `ventas_grande.csv` (~2,5 M de filas, con una región mucho más grande que las demás) para que el join en detalle supere el límite de Excel y probar la división.
  - Casos difíciles: claves duplicadas, espacios, ceros a la izquierda, ñ y tildes, filas basura antes del encabezado
  - Un CSV de 2M filas y un xlsx de 500k filas para performance
  - **Generación determinística** (semilla fija): en cualquier máquina sale lo mismo, así los esperados sirven.
  - QA verifica los scripts en Linux.
- ✅ **Criterios de aceptación:**
  - En Linux, con el repo recién clonado: `scripts/dev.sh` crea el entorno y abre la ventana; `scripts/dev.sh --tests` pasa; correrlo dos veces no reinstala nada.
  - `scripts/build_linux.sh` genera el `.tar.gz` y el autotest pasa.
  - `python scripts/make_fixtures.py` genera todo en `tests/fixtures/generados/` y `git status` no muestra esos archivos (están ignorados).
  - **El PM, en Windows:** doble clic en `scripts\dev.bat` abre la ventana; `scripts\dev.bat -Tests` pasa; `scripts\build_windows.bat` genera el `.zip`, el autotest pasa y el `.exe` abre con doble clic en una PC sin Python.

### Fase 1 — Core de I/O y tipos (2 días)
- `core/io/`: `read_table` y `write_table` para todos los formatos de la sección 1.1, autodetección de CSV, listar hojas (con filas y columnas de cada una), preview de N filas, detectar la fila de encabezado, escribir varias hojas en un mismo .xlsx (con encabezado congelado y filtro).
- `core/types.py`: detección del tipo de una columna (con desglose para "mixto"), conversión de números "1.234,56" / "1,234.56", fechas dd/mm/aaaa, Sí/No.
- ✅ Tests:
  - Cada fixture se lee con las tildes correctas y sin perder ceros a la izquierda.
  - Un CSV guardado con las opciones por defecto se vuelve a leer igual.
  - El tipo detectado de cada columna del archivo "para explorar" coincide con el esperado.
  - El xlsx de 500k filas se lee en un tiempo razonable.

### Fase 2 — Explorador (core + UI base) (3–4 días)
- **2a.** `core/profile.py`: `profile_table`, `value_counts` (una o varias columnas, con opciones), `rows_with_value`, filtros simples, `export_profile`. Por chunks para archivos grandes.
- **2b.** UI base: ventana principal con los 4 botones (los que no existen todavía, deshabilitados), `worker.py` con barra de progreso y cancelar, **grilla de datos paginada** (widget reutilizable sobre `ttk.Treeview`), selector de archivo con panel de hoja/separador.
- **2c.** Ventana del explorador con las 4 pestañas de la sección 3.1.
- ✅ Tests del core contra `esperados/perfil.json` (filas, vacíos, repetidos, distintos, tipos, posible clave, "Norte"/"NORTE"). Desde la UI: explorar el archivo "para explorar", el CSV de 2M filas (que no se cuelgue y muestre progreso) y un .xlsx con varias hojas, cambiando de hoja.
- 🪟 **Prueba en Windows (PM):** `dev.bat -Tests`, abrir el explorador con un CSV exportado por Excel de Windows, y compilar el `.exe`.

### Fase 3 — Core de salida: límites, división y guardado (2 días)
- `core/limits.py`: límites de la tabla 1.2 (de `limites_default.json`, con override por `limites.json`), `estimate`, `check`, `suggest_split_columns`.
- `core/split.py`: `split_dataframe` con los modos `rows`, `rows_keep_groups`, `by_value` (y subdivisión de una parte que siga grande) y `sheets`, como iterador. `split_file` por chunks para archivos existentes.
- `core/output.py`: `save_result` (un archivo / varias hojas / varias partes en carpeta + `_indice.xlsx`), verificación de que la suma de filas coincida, nombres de archivo válidos en Windows.
- ✅ Tests:
  - Con el escenario grande, `check` da 🔴 para Excel y la cantidad de partes sugerida es correcta.
  - Un resultado de 400k filas da 🟡 para Excel y 🟢 para "Solo guardar".
  - Ninguna parte supera el límite.
  - `rows_keep_groups` nunca corta un cliente entre dos partes.
  - `by_value` con la región grande la subdivide.
  - La suma de filas de las partes es igual al total.
  - El índice lista bien las partes.
  - Cambiar un límite en `limites.json` cambia la recomendación.
  - Dividir por un valor con caracteres prohibidos en Windows (`Norte/Sur`, `CON`) genera nombres válidos.
- 💡 El explorador ya puede usar el paso Guardar para "Exportar lo filtrado".

### Fase 4 — Core de resúmenes y join multi-archivo (3–4 días)
- **4a.** `normalize.py`: limpieza de claves.
- **4b.** `aggregate.py`: `summarize(df, AggregationSpec)` con todas las operaciones (usando `core/types.py` para convertir), reporte de celdas no numéricas, agrupar por mes/año de una fecha, nombres automáticos de columnas resultado.
- **4c.** `join.py`: `JoinPlan`/`SecondarySpec`, emparejamiento automático de claves, claves múltiples, modos detalle/resumen/primera por secundario, joins encadenados, detección de "más de un detalle", resumen final, selección y orden de columnas, sufijos legibles, `diagnose_join` por archivo con **estimación de tamaño sin ejecutar el join**, control de filas, reportes de no coincidentes, guardar y cargar la receta en .json (incluido el `OutputSpec`).
- ✅ Tests con el escenario de 3 archivos:
  - **Detalle de ventas + resumen de cobranzas:** una fila por compra; cada fila tiene el total cobrado del cliente.
  - **Resumen de ventas por cliente:** mismas filas que `clientes.xlsx`; `Total comprado` coincide con el esperado.
  - **Resumen por cliente + artículo:** coincide con el esperado.
  - **Resumen final por `Región`:** coincide con el esperado.
  - **Detalle en dos archivos:** el diagnóstico avisa y la cantidad de filas estimada es correcta.
  - `Nro Cliente` "  00123" engancha con `ID cliente` 123.
  - Importes inválidos → se reportan, no rompen la suma.
  - Una receta guardada y vuelta a cargar produce el mismo resultado (incluida la división).
  - Una receta guardada en Linux se puede cargar en Windows (rutas).
  - **Escenario grande:** la estimación de filas de `diagnose_join` coincide con las filas reales.

### Fase 5 — Asistente de Join (UI) + paso Guardar (5–6 días)
- Widgets nuevos:
  - Lista de columnas con checkboxes e **indicador de tamaño** con semáforo.
  - Lista reordenable.
  - **Editor de resumen** (agrupar por + tabla de cálculos).
  - **Paso Guardar** de la sección 3.3: destino, semáforo, opciones de división con vista previa de las partes, sugerencia de columnas, "achicar el resultado", progreso por parte y "Abrir carpeta".
  - Botón "🔍 Ver información" junto a cada archivo, que reutiliza el explorador.
- Asistente de Join completo según la sección 3.2, incluido "+ Agregar archivo", el paso detalle/resumen, el aviso de "más de un detalle", el resumen final y el guardado en varias hojas.
- ✅ Hacer desde la UI, sin tocar código:
  - Las 4 variantes del escenario (detalle, resumen por cliente, resumen por cliente+artículo, resumen final por región).
  - El escenario grande en detalle → la app recomienda dividir y se guardan las partes correctas, para Excel y para Google Sheets.
  - El checklist de la sección 6.
- 🪟 **Prueba en Windows (PM):** el escenario de 3 archivos completo con el `.exe` compilado, incluido "Abrir carpeta" y un archivo abierto en Excel al guardar.

### Fase 6 — Concatenar (1 día)
- Core + asistente (con el paso Guardar común) + "usar este resultado en un cruce".

### Fase 7 — Split de archivos existentes (1 día)
- Asistente de la sección 3.5. El core ya está hecho en la Fase 3; falta `split_file` en streaming probado con archivos de varios GB.

### Fase 8 — Empaquetado final (1 día)
- Ícono definitivo (`.ico` y `.png`), "Acerca de" con versión, ajustar los `excludes` del spec para achicar el paquete, revisar el tamaño y el tiempo de arranque del `.exe`.
- Probar `--onefile` vs carpeta y decidir con el PM cuál se entrega.
- `limites.json` de ejemplo junto al ejecutable, con instrucciones.
- ✅ El `.exe` abre en una Windows limpia sin Python instalado. El binario de Linux abre en la distro más vieja que se soporte.
- 🪟 **Prueba en Windows (PM):** checklist completo de la sección 6 con el `.exe` final.
- 💡 Se puede hacer una entrega intermedia solo con el explorador después de la Fase 2 o 3: los scripts de compilación ya existen desde la Fase 0.

### Fase 9 — Pulido (1–2 días)
- Logs, mensajes de error, recordar carpetas, opciones y destino, "Acerca de" con versión, manual corto (PDF o pantalla de ayuda) con el explorador, el ejemplo de los 3 archivos en sus variantes detalle y resumen, y cómo trabajar con un resultado dividido en partes.

### Fase 10 (opcional) — "Resumir un archivo" como operación propia + anti-join
- Quinto botón en la pantalla de inicio (o pestaña "Resumir" en el explorador) que abre el editor de resumen sobre un solo archivo (reutiliza `aggregate.py` y el widget).
- Opción de anti-join en el asistente de unir.

### Fase 11 — Prueba con usuarios reales
- 2 o 3 personas del negocio con sus archivos reales, usando el `.exe`. Pedir sobre todo:
  - CSV exportados de sus sistemas.
  - Un caso real de 3 o más archivos: uno de "qué compró cada cliente" y otro de "cuánto compró".
  - Un caso que hoy no puedan abrir por tamaño.
- Anotar dónde se traban. Ajustar los textos, el flujo y los **límites "cómodos"** de `limites_default.json`.

---

## 5. CLAUDE.md y equipo

- Las reglas para Claude Code están en **`CLAUDE.md`** (raíz del repo).
- Roles, escalamiento de preguntas, formatos de despacho e informes: **`docs/equipo.md`**.
- Configuración de los agentes: **`.claude/`**.
- Decisiones tomadas: **`docs/decisiones.md`**. Contratos core ↔ UI: **`docs/contratos.md`**.

---

## 6. Checklist de prueba manual (por release)

**Formatos**
- [ ] Abrir un CSV exportado por Excel de Windows (`;`, cp1252): las ñ y tildes se ven bien
- [ ] Abrir un CSV con `,` y utf-8
- [ ] Abrir .xlsx con varias hojas, .xls viejo y .xlsb
- [ ] Los códigos "00123" siguen siendo "00123" después de cualquier operación
- [ ] Un CSV guardado por la app se abre con doble clic en Excel y las columnas quedan separadas
- [ ] Un "CSV para Google Sheets" se importa bien en Google Sheets (tildes y columnas)
- [ ] Concatenar CSV y xlsx mezclados

**Explorador**
- [ ] General: cantidad de filas y columnas correctas; en Excel se listan todas las hojas con sus filas y columnas
- [ ] Cambiar de hoja desde la lista y que se vuelva a analizar
- [ ] CSV: cambiar el separador detectado y que se vuelva a analizar
- [ ] Filas vacías y filas repetidas detectadas, con "ver cuáles"
- [ ] Tipos: entero, decimal con coma, fecha dd/mm/aaaa, Sí/No, texto, código "00123" como Texto, columna mixta con % y "ver cuáles", columna vacía
- [ ] 🔑 Posible clave marcada en `ID cliente`; no marcada en una columna con 2 repetidos
- [ ] Valores distintos de `Región` con cantidad y %, ordenando por cantidad y alfabéticamente, con buscador
- [ ] "Norte" / "NORTE" / "Norte " → aviso y la opción "ignorar mayúsculas y espacios" los junta
- [ ] Combinaciones distintas de 2 columnas
- [ ] Doble clic en un valor distinto → ver sus filas
- [ ] Exportar y copiar la lista de distintos
- [ ] Pestaña Datos: recorrer páginas, ordenar, filtrar por 2 condiciones, exportar lo filtrado
- [ ] CSV de 2M filas: no se cuelga, muestra progreso, se puede cancelar y avisa que no entra en Google Sheets
- [ ] Exportar informe a .xlsx
- [ ] "🔍 Ver información" desde el asistente de unir y sobre el resultado antes de guardar

**Unir archivos — detalle**
- [ ] `clientes` + `ventas` en detalle → una fila por compra con los datos del cliente; se ve qué artículos compró cada uno
- [ ] Se muestra la cantidad de filas del resultado antes de guardar
- [ ] `ventas` y `cobranzas` los dos en detalle → aparece el aviso y "cambiar a resumen" funciona con un clic

**Unir archivos — resumen**
- [ ] `ventas` en resumen agrupando por `ID cliente`, Sumar `Importe` → una fila por cliente con el total correcto
- [ ] Agrupar por `ID cliente` + `Artículo` → total por cliente y artículo correcto
- [ ] Agrupar por mes de `Fecha`
- [ ] Varios cálculos a la vez (Sumar Importe, Contar distintos Artículo, Máximo Fecha) con nombres editados
- [ ] "Listar valores" de `Artículo` muestra "Lápiz, Cuaderno, Goma" sin repetidos
- [ ] Sumar una columna con valores no numéricos → avisa cuáles
- [ ] Resumen final por `Región` con columnas de dos archivos distintos
- [ ] Guardar detalle y resumen en dos hojas del mismo .xlsx

**Unir archivos — general**
- [ ] `clientes` + `ventas` + `cobranzas` por `ID cliente` en modo resumen → un solo archivo con las mismas filas que `clientes`
- [ ] Clave con nombre distinto (`Nro Cliente`) → se elige a mano y funciona
- [ ] La clave se empareja sola con distinta mayúscula o espacios
- [ ] Elegir y reordenar columnas de los 3 archivos; la clave aparece una sola vez
- [ ] Columnas con el mismo nombre en dos archivos → sufijo con el nombre del archivo
- [ ] Clave numérica vs. texto ("00123" vs 123) y espacios al final
- [ ] 2 columnas clave
- [ ] 5 o más archivos secundarios
- [ ] Guardar la receta (con resúmenes y división incluidos), abrirla con archivos del mes siguiente y obtener el resultado
- [ ] Hoja de "sin coincidencia" por archivo

**Resultados grandes (paso Guardar)**
- [ ] Escenario grande en detalle → el aviso aparece ya en el paso 4 y otra vez al guardar
- [ ] El indicador de tamaño del paso 6 baja al desmarcar columnas
- [ ] Destino Excel + más de 1.048.576 filas → 🔴, obliga a elegir, sugiere la cantidad de partes correcta
- [ ] Destino Excel + 400.000 filas → 🟡, recomienda dividir pero deja guardar en un archivo
- [ ] Destino Google Sheets + más de 20 M de celdas → 🔴; no ofrece "varias hojas"
- [ ] "Solo guardarlo" → guarda un CSV único sin bloquear
- [ ] Dividir por filas con "no cortar un mismo cliente" → ningún cliente aparece en dos partes
- [ ] Dividir por `Región` → sugiere la columna y subdivide la región grande
- [ ] "Cambiar a resumen" desde el paso Guardar → el semáforo pasa a 🟢
- [ ] Carpeta con partes numeradas, encabezado en cada una, filtro activado e `_indice.xlsx` correcto
- [ ] Mensaje final con la suma de filas y "Abrir carpeta"
- [ ] Abrir una parte en Excel y otra importada en Google Sheets sin problemas

**Otras operaciones**
- [ ] Concatenar 5 archivos con columnas en distinto orden y una columna faltante
- [ ] Concatenar 12 CSV mensuales → se pasa del límite → recomienda dividir → usar el resultado en un cruce
- [ ] Split de un CSV de 2M filas en partes de 500k
- [ ] Split por valor de columna (muestra cuántos archivos van a salir)

**Entorno y Windows**
- [ ] Repo recién clonado en Windows: doble clic en `scripts\dev.bat` → crea el entorno y abre la app
- [ ] `scripts\dev.bat -Tests` pasa en Windows
- [ ] `scripts\build_windows.bat` genera el .zip y el autotest pasa
- [ ] El .exe abre con doble clic en una PC sin Python, sin consola negra
- [ ] El log se guarda en `%LOCALAPPDATA%\ExcelJoinTool\logs`
- [ ] Archivo abierto en Excel mientras se intenta guardar → mensaje claro
- [ ] Rutas con tildes y espacios (`C:\Users\José\Mis Documentos\`)
- [ ] "Abrir carpeta" funciona en Windows y en Linux
- [ ] Una receta creada en Linux funciona en Windows

---

## 7. Ejemplos de mensajes del PM al líder técnico

- "Arrancá la Fase 0. Mostrame el plan de la fase antes de despachar."
- "Probé en Windows: `dev.bat -Tests` falla con este error: <pegar>. Pasáselo a quien corresponda."
- "Aprobado el cierre de la Fase N. Arrancá la N+1."
- "Antes de seguir, mostrame el estado de las tareas y si hay algo bloqueado."
- "Quiero cambiar <X>. Decime el impacto antes de tocar nada."

---

## 8. Decisiones

El registro completo está en `docs/decisiones.md`. Resumen:

**Tomadas**
- ✅ Formatos: CSV + planillas de Microsoft Excel (.xlsx, .xlsm, .xlsb, .xls), como entrada y salida, mezclables. Se suma lectura de .ods sin costo extra.
- ✅ El join es **1 archivo principal + N secundarios** unidos por una columna clave común (ej. `ID cliente`), eligiendo las columnas de cada uno. Resultado: un único archivo.
- ✅ Por cada secundario con claves repetidas, el usuario elige entre **detalle** (traer todas las filas: "qué compró cada cliente") y **resumen** (agrupar y calcular: "cuánto compró"). No hay opción por defecto: los dos casos son igual de frecuentes.
- ✅ En el resumen, **el usuario elige por qué columnas agrupar** (la clave más las que quiera) **y qué calcular sobre qué columnas** (sumar, contar, promedio, etc.). También se puede resumir el resultado final. Entra en la v1.
- ✅ **Explorador de un archivo** en la v1: información general (hojas, filas, columnas, vacíos, repetidos), tipo de dato por columna, valores distintos por columna (o combinación de columnas) con cantidad, y grilla de datos con filtros. Se construye antes que el join.
- ✅ **Control de tamaño al guardar** en todas las operaciones: el usuario elige si lo va a abrir en Excel o en Google Sheets. Si el resultado es grande, la app recomienda (🟡) u obliga (🔴) a dividirlo en varios archivos: por cantidad de filas sin cortar clientes, o por valor de una columna, con índice de partes. También ofrece achicar el resultado (menos columnas o pasar a resumen). Los límites son configurables.
- ✅ **Equipo:** líder técnico Opus 5.5 (effort medio, no toca código) + backend, frontend y QA en Sonnet 5.5. El PM decide los cambios de arquitectura. Ver `docs/equipo.md`.
- ✅ **Plataformas:** se desarrolla en Linux; corre con Python en Linux y Windows; `.exe` para Windows compilado en Windows. Scripts de entorno y compilación para cada sistema (Fase 0).
- ✅ **Todo en local, sin GitHub Actions ni CI.** Al repo se sube solo código; los archivos de prueba se generan con un script y los ejecutables se comparten por fuera del repo. Las pruebas en Windows las hace el PM en los puntos indicados.

**Abiertas**
1. ¿Tamaño típico de los archivos? (define si alcanza con pandas en memoria o si hace falta streaming/DuckDB; afecta al explorador, a los valores distintos y a la división)
2. ¿Qué destino usan más, Excel o Google Sheets? (define el destino por defecto la primera vez)
3. ¿Hay que firmar el .exe? (sin firma, Windows SmartScreen muestra una advertencia)
4. ¿Qué distro y versión de Linux tienen los usuarios finales, si alguno usa Linux? (define dónde compilar el binario de Linux)
5. ¿Qué PC o VM con Windows se usa para las pruebas y la compilación del `.exe`?
6. ¿Conviene entregar una primera versión solo con el explorador (después de la Fase 2 o 3) para que los usuarios lo prueben?
