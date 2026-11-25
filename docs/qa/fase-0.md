# QA Fase 0 (Linux)

## Resultado: APROBADO (Linux). Windows: pendiente PM
## Suite: pytest 62 passed / 0 failed / 1 skipped (con grandes; el skip es el test del caso "sin grandes") · sin los 2 tests de pantalla: 60 passed · ruff check OK · ruff format --check OK (21 archivos) · git status limpio de generados: SÍ

## Criterios de aceptación de la fase
- [x] Repo recién clonado (copia de `git ls-files -co --exclude-standard`, sin .venv/build/dist/generados): `scripts/dev.sh --setup` crea el entorno (18 s); la segunda corrida dice "Dependencias al día" (0,1 s), sin reinstalar.
- [x] `xvfb-run -a scripts/dev.sh --tests` en la copia: 60 passed, 3 skipped (sin fixtures grandes generados), ruff OK.
- [x] `timeout 8 xvfb-run -a scripts/dev.sh` → código 124 (la ventana quedó abierta).
- [x] `scripts/build_linux.sh` completo (con tests): genera `dist/ExcelJoinTool-0.1.0-linux.tar.gz` (40 MB, SHA256 mostrado), autotest del ejecutable "SELF-TEST OK" (22 s).
- [x] `scripts/make_fixtures.py` en la copia: 21 archivos, 302 MB, 1 min 39 s. `git status -uall` no muestra nada de `generados/`, `dist/`, `build/`, `.venv`. En el repo real, `git status --ignored` marca como ignorados generados, build, dist, .venv, cachés.
- [x] Suite completa en el repo real (con grandes): 62 passed. `--version` y argumento desconocido (código 2) verificados, con log escrito en `~/.local/share/ExcelJoinTool/logs/excel_tool.log`.
- [x] Reglas de CLAUDE.md (grep + lectura):
  - tkinter en `core/`: ninguno (`core/__init__.py` vacío).
  - `sys.platform`/`os.name`: solo en `platform_utils` y `excel_tool.spec` (aceptado). Ver observaciones 1 y 2.
  - print/stdout/stderr: solo `_print` de `app_entry`, protegido por `if stream is not None` (contrato).
  - `open(`: todos con `encoding=` (los 3 de make_fixtures lo llevan en la línea siguiente). El log lleva `encoding="utf-8"`.
  - Rutas con "/" o "\\": ninguna armada a mano (solo la regex de caracteres prohibidos). Sin `/tmp` fijo (usa `tempfile`).
  - `subprocess`: solo `platform_utils.open_folder` (ADR-012), con DEVNULL y sin shell.
  - 1048576 / 20000000: solo en `limites_default.json`, `CLAUDE.md`, `docs/contratos.md` y el esperado `tests/fixtures/esperados/grandes.json` (ver observación 3). Ninguno en código.
  - `.github/`: no existe. `_MEIPASS`: no aparece.
  - Fuentes de `ui/app.py`: solo nombradas (`TkHeadingFont`, `TkSmallCaptionFont`, `TkTextFont`).
  - Archivos propios (ícono, límites) vía `importlib.resources.files("excel_tool")`.
- [x] Fines de línea y permisos: `dev.sh` y `build_linux.sh` LF y 775; `dev.ps1`, `dev.bat`, `build_windows.ps1`, `build_windows.bat` con CRLF en todas las líneas; `git check-attr` da eol=lf para .sh/.py/.spec y eol=crlf para .ps1/.bat.
- [x] Contratos § Fase 0: `main(argv=None) -> int` (sin args abre la app con import diferido, `--version` imprime `Excel Join Tool 0.1.0` y devuelve 0, `--self-test` con los 4 pasos devuelve 0/1, desconocido devuelve 2 con uso a stderr); `ui.app.run()` + `crear_ventana()` separada; ícono con importlib.resources y tolerante a fallas; `platform_utils` con las 4 funciones; `limites_default.json` idéntico al contrato.
- [x] `.gitignore` cubre PLAN §1.5: .venv, build, dist, egg-info, `__pycache__`, cachés, `generados/`, `salidas/`, `*.log`, `.claude/settings.local.json` (6 rutas de prueba confirmadas con `git check-ignore`). Los `.tar.gz`/`.zip` caen bajo `dist/`.
- [ ] Windows (dev.bat, -Tests, build_windows.bat, .exe en PC sin Python) → pendiente PM (ADR-018).

## Bugs
Ninguno.

## Observaciones (no bloquean)
1. `scripts/make_fixtures.py:1153` consulta `sys.platform` (para el pico de memoria, `resource` en Linux/macOS). La regla dice que solo `platform_utils`; es una herramienta de QA, no de la app, pero conviene que el líder lo acepte o lo pida distinto (nivel 2). Revisar que corra en Windows: PM lo prueba en el paso de fixtures (opcional).
2. Los tests (`test_app_entry`, `test_app.py`) usan `sys.platform` para decidir si hay pantalla. Aceptable en tests.
3. `grandes.json` contiene 1048576 como valor esperado (resultado calculado de forma independiente); es un esperado de test, no código.
4. `app_entry._paso_archivos` usa `pd.read_csv`/`to_excel` directo: permitido por contrato en la Fase 0, debe pasar a `core/io` en la Fase 1.
5. Los modos de ejecutable git (+x en .sh) están en el disco; al commitear con `git add` quedan 100755. Verificar con `git ls-files -s scripts/*.sh` después del commit.
6. Sin `tests/fixtures/generados/`, 3 tests se saltean (no fallan): correcto.

## Para probar a mano (PM)
- Abrir `tests/fixtures/generados/formatos/chico.xlsb` con Excel real (escritor propio, ADR-017).

## Para probar en Windows (PM), si la fase lo pide
- `scripts\dev.bat -Help`
- `scripts\dev.bat -Setup` dos veces (la segunda debe decir "Dependencias al dia")
- `scripts\dev.bat -Tests`
- `scripts\dev.bat -AppArgs --version`
- Doble clic en `scripts\dev.bat`: abre la ventana (si falla, la consola queda con pausa).
- `scripts\build_windows.bat` → `dist\ExcelJoinTool-0.1.0-windows.zip` con autotest OK; repetir con `-OneFile`; copiar la carpeta a una PC sin Python y abrir el .exe con doble clic.
- Puntos de riesgo: argumentos vacíos, `-AppArgs --...`, pausa en doble clic, alias falso de Python de la Store, `constraints.txt` generado en Linux (puede faltar algún paquete exacto en Windows), SmartScreen (el autotest del build queda esperando si pregunta).
- Opcional: `python scripts\make_fixtures.py` (ver observación 1) y la suite con los fixtures generados.
