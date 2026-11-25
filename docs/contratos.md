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
| `ReadOptions`, `CsvOptions`, `read_table`, `write_table` | 1 | Por definir |
| `ColumnType`, funciones de `core/types.py` | 1 | Por definir |
| `TableProfile`, `ColumnProfile`, `SheetSummary`, `value_counts`, `rows_with_value` | 2 | Borrador en PLAN.md §2 |
| Eventos de progreso y cancelación (`worker`) | 2 | Por definir |
| Errores del core hacia la UI (excepciones y mensajes) | 2 | Por definir |
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
