# Registro de decisiones

Formato: ADR-NNN · fecha · nivel (2 = líder, 3 = PM) · decisión · motivo · alternativas descartadas · impacto

Las decisiones nuevas se agregan al final. No se borran: si una cambia, se agrega una nueva que la reemplaza y se indica "reemplaza ADR-NNN".

- ADR-001 · 2026-11-25 · 3 · Formatos: CSV + Excel (.xlsx/.xlsm/.xlsb/.xls), entrada y salida, mezclables; lectura de .ods.
- ADR-002 · 2026-11-25 · 3 · Join = 1 principal + N secundarios por columna clave, eligiendo columnas por archivo.
- ADR-003 · 2026-11-25 · 3 · Por secundario: modo detalle o resumen, elegido siempre por el usuario (sin default).
- ADR-004 · 2026-11-25 · 3 · Resumen con agrupar-por y cálculos elegidos por el usuario; también sobre el resultado final. En v1.
- ADR-005 · 2026-11-25 · 3 · Explorador de un archivo en v1, construido antes que el join.
- ADR-006 · 2026-11-25 · 3 · Control de tamaño antes de guardar (Excel / Google Sheets / solo guardar) con recomendación de dividir; límites configurables.
- ADR-007 · 2026-11-25 · 3 · Equipo: líder técnico Opus 5.5 (effort medio, no toca código) + backend, frontend y QA en Sonnet 5.5. El PM decide los cambios de arquitectura.
- ADR-008 · 2026-11-25 · 3 · Plataformas: desarrollo en Linux; la app corre con Python en Linux y Windows; .exe compilado en Windows con PyInstaller. Scripts de entorno (dev) y compilación (build) para cada sistema, en la Fase 0.
- ADR-009 · 2026-11-25 · 3 · Todo local: sin GitHub Actions ni CI. Al repo se sube solo código. Archivos de prueba generados por script (no se suben); solo se suben resultados esperados chicos. Ejecutables fuera del repo. Pruebas en Windows a cargo del PM al cerrar las Fases 0, 2, 5 y 8.
- ADR-010 · 2026-11-25 · 3 · Los agentes no commitean; commitea el líder cuando QA aprueba. Nadie hace push salvo el PM.
- ADR-011 · 2026-11-25 · 3 · La configuración `.claude/` (agentes y hook `limitar_rutas.py`) la arma el líder como tarea F0-00, por única vez y con OK del PM. Cambios posteriores a `.claude/` los hace o aprueba el PM. El hook también bloquea `git commit/push/reset/...` a backend, frontend y qa, y `.github/` a todos. `settings.json` lo crea el PM (los permisos de comandos los decide él).
- ADR-012 · 2026-11-25 · 2 · `platform_utils.open_folder` usa `os.startfile` en Windows y `xdg-open` vía `subprocess` en Linux. Es la única excepción a "nada de `subprocess`" y vive solo en `platform_utils`. Motivo: lo pide PLAN §1.4; la regla apunta a comandos como `ls`/`cp`. Descartado: `webbrowser.open` (puede abrir el navegador en vez del explorador de archivos).
- ADR-013 · 2026-11-25 · 2 · `dev.sh --tests` usa `xvfb-run -a` solo si no hay `DISPLAY` y `xvfb-run` existe. Los tests que necesitan pantalla se saltan con un aviso claro si no hay pantalla.
- ADR-014 · 2026-11-25 · 2 · `make_fixtures.py` sin opciones genera todo (criterio de la Fase 0). `--sin-grandes` saltea el escenario grande y los archivos de rendimiento, para el día a día.
- ADR-015 · 2026-11-25 · 2 · Escritura de .xlsx solo con `openpyxl` por ahora. `xlsxwriter` se evalúa en la Fase 3 si hace falta velocidad (agregarla sería nivel 3).
- ADR-016 · 2026-11-25 · 3 · Los fixtures `.xls` y `.xlsb` no se generan (no hay escritor sin sumar dependencias). El PM los guarda desde Excel en Windows y van chicos a `tests/fixtures/esperados/` como excepción. Hasta que lleguen, los tests que los usan se saltan con aviso.
- ADR-017 · 2026-11-25 · 3 · **Reemplaza ADR-016.** Los fixtures `.xls` y `.xlsb` se generan con `make_fixtures.py`, como los demás. `.xls`: con `xlwt`, solo como dependencia de desarrollo (extra `dev`, no va en el ejecutable). `.xlsb`: no hay librería que lo escriba, así que QA arma un escritor mínimo propio dentro de `make_fixtures.py` (zip + registros BIFF12) y lo valida leyéndolo con calamine. Si `xlwt` no anda en Python 3.12, el `.xls` también se escribe a mano. Riesgo: que Excel real no abra el `.xlsb` generado; el PM lo comprueba en la prueba de Windows.
- ADR-018 · 2026-11-25 · 3 · Windows: el PM tiene una PC con Windows (otra máquina) y prueba al final de la fase. La Fase 0 cierra primero en Linux; la parte de Windows de los criterios de aceptación queda pendiente hasta esa prueba.
- ADR-019 · 2026-11-25 · 2 · La regla "solo `platform_utils` pregunta `sys.platform`" aplica al código de la app (`src/excel_tool/`). Scripts de desarrollo y QA (`scripts/`, `tests/`) y el `.spec` pueden consultarlo. Motivo: no se empaquetan ni los usa el usuario final (observación 1 de QA, Fase 0).

## Preguntas abiertas
1. Tamaño típico de los archivos (pandas en memoria vs. streaming/DuckDB).
2. Destino más usado: Excel o Google Sheets (default de la primera vez).
3. ¿Firmar el .exe? (sin firma, Windows SmartScreen muestra advertencia).
4. Distro/versión de Linux de los usuarios finales, si alguno usa Linux.
5. PC o VM con Windows para pruebas y compilación. **Bloquea el cierre de la Fase 0** (sin eso cierra solo la parte Linux).
6. ¿Entrega temprana solo con el explorador?
