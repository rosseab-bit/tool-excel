# Excel Join Tool

App de escritorio para explorar, unir, juntar y dividir archivos Excel y CSV sin usar fórmulas.

El plan está en [`PLAN.md`](PLAN.md) y la forma de trabajo del equipo en [`docs/equipo.md`](docs/equipo.md).

## Requisitos

- **Python 3.12 o más nuevo, con tkinter.**
- **Linux:** `sudo apt install python3-tk python3-venv` (Ubuntu/Debian; la versión de venv puede llamarse `python3.12-venv`). En Fedora: `sudo dnf install python3-tkinter`. Para correr los tests sin pantalla: `sudo apt install xvfb`.
- **Windows:** instalador de [python.org](https://www.python.org/downloads/), marcando **"tcl/tk and IDLE"** y **"Add python.exe to PATH"** (o usando el launcher `py`, que se instala con él).
- Los scripts buscan Python solos. Para forzar uno: variable `PYTHON` (ej.: `PYTHON=/usr/bin/python3.12 scripts/dev.sh`).

## Linux

| Tarea | Comando |
|---|---|
| Abrir la app (prepara el entorno si hace falta) | `scripts/dev.sh` |
| Solo preparar el entorno | `scripts/dev.sh --setup` |
| Correr los tests y revisar el código | `scripts/dev.sh --tests` |
| Pasarle opciones a la app | `scripts/dev.sh -- --version` o `-- --self-test` |
| Regenerar `constraints.txt` | `scripts/dev.sh --lock` |
| Borrar `.venv` y empezar de cero | `scripts/dev.sh --recreate` |
| Instalar solo lo necesario para usar la app | `scripts/dev.sh --no-dev` |
| Crear los archivos de prueba | `.venv/bin/python scripts/make_fixtures.py` (con `--sin-grandes` salta los archivos grandes) |
| Compilar, carpeta | `scripts/build_linux.sh` |
| Compilar, un solo archivo | `scripts/build_linux.sh --onefile` |

`scripts/dev.sh --help` muestra todas las opciones. Las dependencias se reinstalan solo si cambió `pyproject.toml` o `constraints.txt`.

La compilación corre los tests (se saltean con `--skip-tests`), prueba el ejecutable con `--self-test` y deja el resultado en `dist/ExcelJoinTool-<versión>-linux.tar.gz`, mostrando tamaño y SHA256.

**Binario de Linux:** conviene compilarlo en la distro más vieja que haya que soportar. El binario es compatible hacia adelante con versiones más nuevas de glibc, pero no hacia atrás.

## Windows

Todos los scripts se pueden abrir con doble clic o desde una consola. Usan `-ExecutionPolicy Bypass`, así que no hace falta tocar la configuración de PowerShell.

| Tarea | Comando |
|---|---|
| Abrir la app (prepara el entorno si hace falta) | `scripts\dev.bat` |
| Solo preparar el entorno | `scripts\dev.bat -Setup` |
| Correr los tests y revisar el código | `scripts\dev.bat -Tests` |
| Pasarle opciones a la app | `scripts\dev.bat -AppArgs --version` o `-AppArgs --self-test` |
| Regenerar `constraints.txt` | `scripts\dev.bat -Lock` |
| Borrar `.venv` y empezar de cero | `scripts\dev.bat -Recreate` |
| Instalar solo lo necesario para usar la app | `scripts\dev.bat -NoDev` |
| Crear los archivos de prueba | `.venv\Scripts\python.exe scripts\make_fixtures.py` (con `--sin-grandes` salta los archivos grandes) |
| Compilar, carpeta | `scripts\build_windows.bat` |
| Compilar, un solo `.exe` | `scripts\build_windows.bat -OneFile` |

`scripts\dev.bat -Help` muestra todas las opciones. La compilación corre los tests (se saltean con `-SkipTests`), prueba el `.exe` con `--self-test` y deja el resultado en `dist\ExcelJoinTool-<versión>-windows.zip`, mostrando tamaño y SHA256.

## Prueba en Windows al cerrar una fase

No hay CI: la prueba la hace a mano el PM en una PC o VM con Windows al cerrar la Fase 0 y, después, las Fases 2, 5 y 8 (PLAN §1.4). Con el repo en la PC:

1. `scripts\dev.bat` — tiene que crear el entorno, instalar las dependencias y abrir la ventana de la app.
2. `scripts\dev.bat -Tests` — tests y ruff en verde; el último mensaje es "Todo bien".
3. `scripts\build_windows.bat` — termina sin errores, el self-test del `.exe` da OK y se crea el `.zip` en `dist\`.
4. Descomprimir el `.zip` en una **PC sin Python instalado** y abrir `ExcelJoinTool.exe` con doble clic: tiene que abrirse la ventana sin consola negra.
5. A partir de la Fase 2, además, el checklist de Windows de PLAN §6.

**SmartScreen:** el `.exe` no está firmado, así que Windows puede mostrar "Windows protegió su PC". Hay que elegir "Más información" y después "Ejecutar de todas formas". Si aparece durante la compilación, el autotest del build puede quedar colgado esperando: hay que aceptar el aviso para que continúe.

## Qué no se sube al repo

Solo código: nada de `.venv/`, `build/`, `dist/`, ejecutables, archivos de prueba generados ni logs. Detalle en [PLAN §1.5](PLAN.md).
