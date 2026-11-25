# Excel Join Tool

App de escritorio tkinter para usuarios de negocio NO técnicos: explorar un archivo (info general, tipos por columna,
valores distintos, ver y filtrar datos), unir varios archivos por una columna clave (1 principal + N secundarios,
cada uno en modo detalle o resumen), concatenar y dividir archivos CSV y Excel. Si un resultado es demasiado grande
para Excel o Google Sheets, la app recomienda dividirlo en varios archivos.

Plan completo: `PLAN.md`. Equipo y flujo de trabajo: `docs/equipo.md`.

## Equipo
- La sesión principal es el **líder técnico** (Opus 5.5): planifica, coordina y NO escribe código.
- Agentes (Sonnet 5.5): **backend** (core/, scripts, empaquetado), **frontend** (ui/), **qa** (fixtures, tests de integración, verificación).
- Cada agente edita solo sus carpetas (lo fuerza un hook). Contratos core ↔ UI en `docs/contratos.md`.
- Dudas: nivel 1 decide el agente, nivel 2 el líder, nivel 3 (arquitectura, alcance, dependencias, lo que ve el usuario) el **PM**.
- Los agentes no commitean. Commitea el líder cuando QA aprueba. **Nadie hace push**: lo hace el PM.
- Decisiones en `docs/decisiones.md`. Formatos de despacho e informes en `docs/equipo.md`.

## Entorno y plataformas
- Se desarrolla en **Linux**. Tiene que correr en **Linux y Windows** y compilarse como `.exe` en Windows.
- **Todo local. No hay GitHub Actions ni CI**: no crear `.github/workflows/`.
- **Al repo se sube solo código.** Nunca commitear: `.venv/`, `build/`, `dist/`, ejecutables, `tests/fixtures/generados/`, logs, salidas de prueba.

## Reglas de código
- `src/excel_tool/core/` es lógica pura: NUNCA importar tkinter ahí.
- Toda lectura pasa por `core/io.read_table`. Todo resultado se guarda con `core/output.save_result` (que controla límites y divide). Nada de `pd.read_csv` / `pd.read_excel` / `to_excel` / `to_csv` sueltos en otros módulos.
- Leer siempre como texto (`dtype=str`). La detección de tipos y la conversión a número/fecha están SOLO en `core/types.py`.
- Reportar siempre las celdas que no se pudieron convertir.
- CSV de salida por defecto: `;`, `utf-8-sig`, coma decimal (para que abra bien en Excel de Windows en Argentina). Para Google Sheets: `,` + `utf-8`.
- Los límites de Excel y Google Sheets están SOLO en `core/limits.py` (valores en `limites_default.json`, override con `limites.json`). Nunca hardcodear 1048576 ni 20_000_000 en otro lado.
- Controlar el tamaño ANTES de escribir. Al dividir, escribir cada parte a medida que se genera y verificar que la suma de filas de las partes sea igual al total.
- El explorador nunca carga todas las filas en la grilla: paginar. Los valores distintos se calculan a pedido.
- El join trabaja siempre sobre un `JoinPlan` (principal + lista de secundarios). Nunca asumir que hay solo 2 archivos.
- Cada secundario tiene modo `detail` | `summary` | `first`. Los resúmenes usan SIEMPRE `aggregate.summarize` (también el resumen final).
- Sin secundarios en `detail` y sin resumen final, el resultado DEBE tener las mismas filas que el principal: verificarlo en código y en tests.
- Más de un secundario en `detail` → el diagnóstico tiene que avisarlo.
- Toda función nueva del core lleva test, probando al menos un CSV y un xlsx.
- La UI no hace cálculos: arma un dataclass de config y llama al core.
- Operaciones largas → `ui/worker.py` (hilo + queue). Nunca tocar widgets desde otro hilo.
- Textos de UI en castellano rioplatense, sin jerga técnica (no decir "join", "merge", "group by", "distinct", "split", "null", "DataFrame", "encoding", "dtype").
- Errores al usuario: mensaje amigable. Detalle técnico → log.
- Solo ttk/tkinter estándar, sin librerías de UI externas.

## Reglas de compatibilidad Linux ↔ Windows
- Rutas siempre con `pathlib.Path`. Nunca concatenar rutas con `/` o `\`.
- Abrir archivos SIEMPRE con `encoding=` explícito. CSV con `newline=""` al escribir.
- Temporales con `tempfile`, nunca `/tmp` fijo.
- Todo lo que depende del sistema (carpetas de usuario, abrir carpeta, nombres de archivo válidos en Windows) va SOLO en `src/excel_tool/platform_utils.py`. Ningún otro módulo pregunta `sys.platform`.
- La app nunca escribe a stdout/stderr (en el `.exe` no hay consola): loguear a archivo.
- Archivos propios (ícono, límites) con `importlib.resources.files("excel_tool")`, nunca `sys._MEIPASS`.
- Fuentes de tkinter genéricas (`TkDefaultFont`). Nada de comandos de sistema con `subprocess` (única excepción: `xdg-open` en `platform_utils.open_folder`, ADR-012).
- Scripts `.sh` con LF; `.ps1` y `.bat` con CRLF (lo fija `.gitattributes`).

## Comandos (Linux)
- Entorno + abrir la app: `scripts/dev.sh`
- Solo preparar el entorno: `scripts/dev.sh --setup`
- Tests + lint: `scripts/dev.sh --tests` (sin pantalla: `xvfb-run -a scripts/dev.sh --tests`)
- Tests sueltos: `.venv/bin/python -m pytest tests/core -q`
- Lint: `.venv/bin/ruff check . && .venv/bin/ruff format .`
- Archivos de prueba: `.venv/bin/python scripts/make_fixtures.py`
- Compilar ejecutable de Linux: `scripts/build_linux.sh`

## Comandos (Windows, los corre el PM)
- Entorno + abrir la app: doble clic en `scripts\dev.bat`
- Tests: `scripts\dev.bat -Tests`
- Compilar el `.exe`: `scripts\build_windows.bat`

(Estos scripts se crean en la Fase 0. Hasta entonces no existen.)

## Flujo
- Antes de implementar algo no trivial, proponé el plan y esperá el OK.
- Una tarea = un commit con mensaje en castellano (lo hace el líder).
- Al terminar una tarea, correr tests y lint.
