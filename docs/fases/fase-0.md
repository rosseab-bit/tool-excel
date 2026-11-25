# Fase 0 — Plan aprobado (2026-11-25)

Fuente: PLAN.md §1.3, §1.4, §1.5 y Fase 0. Contratos: `docs/contratos.md` § Fase 0. Decisiones: ADR-011 a ADR-018 (ADR-017 reemplaza a ADR-016).

## Estado

| # | Agente | Tarea | Depende de | Estado |
|---|---|---|---|---|
| F0-00 | líder + PM | `.claude/` (agentes + hook) · borrar duplicados de la raíz · `settings.json` (PM) | — | Hecho |
| F0-01 | backend | `pyproject.toml`, `.gitignore`, `.gitattributes`, esqueleto del paquete, `app_entry`, `platform_utils`, `limites_default.json`, tests en `tests/core/` | F0-00 | Hecho (32 tests). F0-01b: ruff excluye `*.md` |
| F0-02 | backend | `scripts/dev.sh` (todas las opciones) + `constraints.txt` con `--lock` | F0-01 | Hecho |
| F0-03 | backend | `scripts/dev.ps1` + `dev.bat` (sin poder probarlos; justificar en el informe) | F0-02 | Parcial: escritos y revisados a mano, falta prueba en Windows |
| F0-04 | backend | `excel_tool.spec`, `packaging/launcher.py`, `build_linux.sh`, `build_windows.ps1/.bat` | F0-02, F0-10 | Parcial: Linux OK (carpeta 40 MB, onefile 39 MB, autotest OK); Windows sin ejecutar |
| F0-05 | backend | `README.md` | F0-03, F0-04 | Hecho |
| F0-10 | frontend | `ui/app.py` mínimo (`run()`), íconos `.png` + `.ico`, `tests/ui/` | F0-01 | Hecho (self-test completo da 0) |
| F0-20 | qa | `make_fixtures.py` base + `conftest.py`: formatos (incluye `.xls` con xlwt y `.xlsb` con escritor propio, ADR-017), variantes CSV, "para explorar" + `perfil.json` | F0-02 | Hecho (49 tests) |
| F0-21 | qa | Escenario de 3 archivos + casos difíciles + esperados | F0-20 | Hecho (59 tests) |
| F0-22 | qa | Escenario grande + archivos de rendimiento | F0-20 | Hecho (62 tests con grandes; 1 min 40 s, ~300 MB) |
| F0-30 | qa | Verificación → `docs/qa/fase-0.md` | todo | APROBADO (Linux); Windows pendiente PM |

Paralelo a partir de F0-02: backend (F0-03 → F0-05), frontend (F0-10), qa (F0-20 → F0-22).

## Pendientes del PM
- Crear `.claude/settings.json` (ver abajo) y reiniciar la sesión para que aparezcan los agentes.
- ~~`.xls`/`.xlsb` desde Excel~~ → se generan (ADR-017).
- Windows: PC aparte, prueba al final de la fase (ADR-018). Prueba mínima de `dev` (F0-03):
  `scripts\dev.bat -Help` · `scripts\dev.bat -Setup` (dos veces: la segunda dice "Dependencias al dia") ·
  `scripts\dev.bat -Tests` · `scripts\dev.bat -AppArgs --version` · doble clic en `dev.bat` (abre la ventana).
  Puntos de riesgo: argumentos vacíos, `-AppArgs --...`, la pausa al hacer doble clic, el alias falso de Python de la Store,
  `constraints.txt` generado en Linux (puede faltar algún paquete exacto para Windows).
  Compilación (F0-04): `scripts\build_windows.bat` → `dist\ExcelJoinTool-0.1.0-windows.zip` con autotest OK; repetir con `-OneFile`;
  copiar la carpeta a una PC sin Python y abrir el .exe con doble clic. Riesgo: si SmartScreen pregunta, el autotest del build queda esperando.

`.claude/settings.json` propuesto:

```json
{
  "agent": "lider-tecnico",
  "permissions": {
    "allow": [
      "Bash(.venv/bin/python:*)", "Bash(.venv/bin/ruff:*)", "Bash(.venv/bin/pytest:*)",
      "Bash(scripts/dev.sh:*)", "Bash(scripts/build_linux.sh:*)", "Bash(xvfb-run:*)",
      "Bash(git status:*)", "Bash(git diff:*)", "Bash(git log:*)", "Bash(git show:*)"
    ],
    "deny": ["Bash(git push:*)"]
  }
}
```

## Despacho F0-01 (listo para enviar)

```
Tarea: F0-01 — Base del proyecto y esqueleto del paquete
Contexto: PLAN.md §1 (tabla), §1.3 ("Detalles de empaquetado"), §1.4, §1.5, §2 (árbol) y Fase 0. ADR-012, ADR-015.
Contrato: docs/contratos.md § Fase 0 (__init__, app_entry, platform_utils, limites_default.json). No cambiar firmas sin consultar.
Archivos que podés tocar: pyproject.toml, .gitignore, .gitattributes, src/excel_tool/__init__.py, __main__.py,
  app_entry.py, platform_utils.py, limites_default.json, src/excel_tool/core/__init__.py, tests/__init__.py (si hace falta),
  tests/core/.
Criterio de aceptación:
  - pyproject.toml: Python >=3.12; dependencias pandas, python-calamine, openpyxl; extras dev (pytest, ruff, xlwt — ADR-017) y build
    (pyinstaller); script `excel-tool = excel_tool.app_entry:main`; versión dinámica desde __init__; paquete en src/;
    incluye limites_default.json y assets/ como datos del paquete; config de ruff (line-length 100) y pytest (testpaths, marker `pantalla`).
  - .gitignore según PLAN §1.5 (incluye tests/fixtures/generados/, salidas/, *.log, .claude/settings.local.json).
  - .gitattributes según PLAN §1.4: *.sh LF, *.ps1 y *.bat CRLF, resto LF, xlsx/xls/xlsb/xlsm/ods/ico/png binarios.
  - app_entry.main cumple el contrato (sin argumentos, --version, --self-test, argumento desconocido). Si
    excel_tool.ui.app todavía no existe, el self-test marca ese paso como FALLA (no lo saltea).
  - platform_utils cumple el contrato. Tests de safe_filename con todos los casos (prohibidos, reservados con y sin
    extensión, vacío, largo, tildes/ñ se conservan), de user_data_dir/log_dir con monkeypatch de entorno y HOME en
    tmp_path, y de open_folder con subprocess/os.startfile mockeados (rama Windows y Linux).
  - Tests de --version y --self-test llamando a main() en proceso (el paso de ventana con marker `pantalla`).
  - Crear .venv a mano para esta tarea (python3.12 -m venv .venv && .venv/bin/pip install -e ".[dev]");
    dev.sh viene en F0-02.
  - pytest y ruff check + ruff format --check pasan (xvfb-run -a si no hay pantalla).
Fuera de alcance: scripts/, ui/, assets/, spec, packaging, README, constraints.txt.
```

## Notas para fases siguientes
- Fase 1: calamine con `dtype=str` puede devolver los números enteros de Excel como "100.0" (QA, F0-20). `read_table` tiene que devolver "100". Los fixtures actuales evitan enteros en celdas numéricas de Excel: la Fase 1 tiene que agregar uno a propósito.
- Fase 1: `CalamineSheet.total_height` da una fila de menos en xlsx escritos con openpyxl `write_only` (sin `<dimension>`). No usarlo para contar filas al listar hojas.
- Fase 1: `app_entry._paso_archivos` (self-test) usa pandas directo; pasarlo a `core/io` (`read_table`/`write_table`).
- Fase 2: convenciones de conteo de `perfil.json` (QA, F0-20): distintos sin normalizar ("Norte", "NORTE" y "Norte " son 3), filas repetidas = copias de una fila anterior no vacía, filas vacías aparte. "Código" con ceros a la izquierda es texto.
- Fase 4: limpieza de claves con ceros a la izquierda ("00012" vs "12"): decidir si es opción del usuario. Los esperados del escenario de 3 archivos traen las dos variantes.
- PM en Windows: abrir `tests/fixtures/generados/formatos/chico.xlsb` con Excel real (generado con escritor propio, ADR-017).
