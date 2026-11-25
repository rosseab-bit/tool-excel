# Fase 1 — Core de I/O y tipos (cerrada 2026-11-25; Windows pendiente PM)

Fuente: PLAN.md §1.1, §2 (`core/io`, `core/types.py`) y Fase 1. Contratos: `docs/contratos.md` § Fase 1 (cerrados). Notas heredadas: `docs/fases/fase-0.md` § "Notas para fases siguientes".
Ahorro de tokens (ADR-020): una tarea = un agente nuevo, con registro de avance (docs/equipo.md §4.4).

## Tareas

| # | Agente | Tarea | Depende de | Estado |
|---|---|---|---|---|
| F1-10 | qa | Fixtures nuevos para la Fase 1 + esperados | — | Hecho (71 tests; error de fórmula = vacío, ADR-022) |
| F1-01 | backend (A) | `core/errors.py`, `core/io/csv.py` (separador, encoding), `core/io/header.py` + tests | — | Hecho (2M filas en 3,2 s) |
| F1-04 | backend (B) | `core/types.py` (detección + conversiones) + tests | — | Hecho (66 tests; 2M filas en 4,5 s) |
| F1-02 | backend (A) | `core/io/excel.py` + `read_table` / `list_sheets` en `core/io/__init__.py` + tests con todos los fixtures | F1-01, F1-10 | Hecho (225 tests; 500k xlsx en ~9 s) |
| F1-03 | backend (A) | `write_table` / `write_sheets` + ida y vuelta + pasar el self-test a `core/io` | F1-02, F1-04 | Hecho (242 tests) |
| F1-20 | qa | Tests de integración + verificación → `docs/qa/fase-1.md` | todo | APROBADO (309 tests) |

Paralelo al arrancar: F1-10 (qa), F1-01 (backend A) y F1-04 (backend B). No se pisan: A toca `core/errors.py` y `core/io/`, B toca `core/types.py`; los tests van en archivos distintos de `tests/core/`.
Frontend: sin tareas en esta fase.

## F1-10 — fixtures nuevos (qa)
- Excel con **números enteros en celdas numéricas** (100 → "100"), decimales, fechas y fecha-hora como fecha real de Excel, booleanos, una fórmula (con valor calculado guardado) y celdas combinadas. En .xlsx y, si se puede, en .xls.
- CSV con separador `|`; CSV con estilo `1,234.56`; CSV solo con encabezado; CSV vacío (0 bytes); archivo con extensión .xlsx que no es un Excel (dañado).
- Esperados: valores exactos como texto según las reglas del contrato (enteros sin ".0", fechas `dd/mm/aaaa`, etc.) y el tipo esperado por columna.

## Criterios de aceptación de la fase (PLAN Fase 1 + contrato)
- Cada fixture se lee con `read_table` con tildes y ñ correctas y sin perder ceros a la izquierda; las filas basura se descartan y el encabezado detectado coincide con `dificiles.json`.
- `list_sheets` da nombre, filas y columnas correctas de cada hoja (incluido el xlsx de 500k).
- Un CSV guardado con las opciones por defecto se vuelve a leer igual (todos los formatos de salida, ida y vuelta).
- El tipo detectado de cada columna de "para explorar" coincide con `perfil.json` (desglose de la mixta incluido).
- Conversión de números, fechas y Sí/No con las celdas no convertibles reportadas.
- Rendimiento (medido y reportado; el test falla solo si se pasa del doble): xlsx de 500k filas leído en ≤ 30 s; CSV de 2M leído en ≤ 20 s; `detect_types` sobre el CSV de 2M en ≤ 15 s.
- Ningún `pd.read_csv` / `read_excel` / `to_csv` / `to_excel` fuera de `core/io` (el self-test pasa a usar `core/io`).
- pytest + ruff completos pasan; `git status` limpio de generados.

## Notas
- Windows: un archivo abierto en Excel podría dar un `OSError` distinto de `PermissionError` y mostrarse como "No se pudo abrir el archivo" en vez de "Cerrá el archivo en Excel". Probarlo en Windows (abrir un .xlsx en Excel y leerlo con la app) y refinar si pasa.
- Fase 2: advertencia de fórmulas/celdas combinadas no implementada (calamine no las expone).
- `list_sheets` tarda ~8 s con una hoja de 500k filas (lee la hoja para contar filas reales). Tenerlo en cuenta en el explorador (Fase 2): mostrarlo con progreso.
- Fase 3: escribir xlsx de 500k × 10 tarda 57 s (texto) a 80 s (con tipos) con openpyxl write_only; CSV 2–10 s. Evaluar `xlsxwriter` (ADR-015, dependencia nueva = nivel 3) y guardar siempre con progreso en el worker.
- Fase 3: `WriteResult` no tiene `warnings` (el ajuste de nombre de hoja va al log). Si el paso Guardar lo tiene que mostrar, agregar el campo al contrato.
- `writer.py` cierra generadores internos de openpyxl vía `ws._rows` (atributo privado) si falla el guardado. Revisar si se actualiza openpyxl.
- Fase 2: `detect_types` sobre 2M filas midió 4,5 s (backend) y 14,7 s (QA), al borde del objetivo de 15 s. Vigilar: el explorador lo corre en archivos grandes (con progreso, por partes).
- Fase 2/4: `ConversionResult.failed` no incluye las celdas vacías (contrato). Si la UI avisa "N celdas de Importe no son números", las vacías se informan aparte.
- Escribir en una ruta que es una carpeta da un mensaje genérico ("Revisá que haya espacio…"). Afinar en la Fase 3 (paso Guardar).
