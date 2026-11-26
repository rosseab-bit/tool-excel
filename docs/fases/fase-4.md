# Fase 4 — Core de resúmenes y unir varios archivos (cerrada en Linux 2026-11-26; Windows pendiente PM)

Fuente: PLAN.md §2 (modelo del join), §3.2 y Fase 4. Contratos: `docs/contratos.md` § Fase 4 (propuestos). Escenario: `tests/fixtures/generados/escenario_3/` + `esperados/escenario_3/`.

## Tareas

| # | Agente | Tarea | Depende de | Estado |
|---|---|---|---|---|
| F4-01 | backend | `core/normalize.py`: limpieza de claves (`KeyOptions`: recortar espacios, ignorar mayúsculas, ceros a la izquierda) + emparejar columnas clave por nombre | — | Hecho (15 tests; cobranzas 5 / 27 coincidencias según la opción; 2M en 0,24 s) |
| F4-02 | backend | `core/aggregate.py`: `summarize` (sumar, contar, contar distintos, promedio, mín, máx, primero, último, listar valores), agrupar por mes/año de una fecha, nombres automáticos, reporte de celdas no numéricas | — | Hecho (36 tests; 2M suma por Región 7 s / 956 MB, por ID 13 s / 1.275 MB) |
| F4-03 | backend | `core/join.py`: `JoinPlan`, `diagnose_join` (por archivo, estimación de filas y tamaño **sin unir**, aviso de más de un detalle, estimación de memoria), `run_join` (detalle / resumen / primera fila, encadenado, resumen final, orden de columnas, control de filas), con progreso y cancelar | F4-01, F4-02 | Hecho (45 tests; grande 200k × 2,5M: detalle 5 s / 937 MB, estimación = reales) |
| F4-05 | backend | Acelerar `types.to_number` / `to_date` convirtiendo solo los valores distintos (hoy 6,7 s sobre 2M, el grueso de `summarize`) | F4-02 | Hecho (to_number 2M 6,7 → 2,4 s; summarize por Región 7,2 → 2,6 s; guardar CSV con tipos 29 → 17 s) |
| F4-04 | backend | Recetas: guardar y cargar `JoinPlan` + `OutputSpec` en .json (rutas que sirvan en Linux y Windows, versión del formato) | F4-03 | Hecho (47 tests; `RecipeLoad`) |
| F4-10 | frontend | Editor de resumen reutilizable (§3.2 paso 4) contra el contrato, adelantado de la Fase 5 | contrato | Hecho contra el contrato (21 tests; `ui/editor_resumen.py`, botón \"Resumir…\" en Datos); falta probar con `aggregate` real |
| F4-20 | qa | Integración con el escenario de 3 archivos + escenario grande (estimación = filas reales, memoria) + verificación → `docs/qa/fase-4.md` | todo | APROBADO (1152 tests; `docs/qa/fase-4.md`) |

Paralelo: F4-01 y F4-02 juntos (y F4-10 si se aprueba); F4-03 después; F4-04 al final.

## Criterios de aceptación (PLAN Fase 4)
- Detalle de ventas + resumen de cobranzas: una fila por compra, cada una con el total cobrado del cliente.
- Resumen de ventas por cliente: mismas filas que clientes; `Total comprado` coincide con el esperado. Por cliente + artículo y resumen final por Región coinciden.
- Dos archivos en detalle: el diagnóstico avisa y la cantidad estimada de filas es correcta.
- Ceros a la izquierda (ADR-034): sin la opción, cobranzas coincide como `solo_recortar_espacios`; con la opción, como `sin_ceros_a_la_izquierda`; el diagnóstico sugiere activarla.
- Importes inválidos se reportan y no rompen la suma.
- Sin secundarios en detalle y sin resumen final, el resultado tiene exactamente las filas del principal (verificado en código).
- Receta guardada y cargada da el mismo resultado (incluida la división); una receta de Linux carga en Windows.
- Escenario grande: la estimación de filas de `diagnose_join` coincide con las reales; pico de memoria dentro de 1,5 GB o aviso previo si no alcanza.

## Notas
- 2026-11-26: F4-01, F4-02 y F4-10 bloqueadas: el hook del líder se aplica también a los subagentes desde que se retomó la sesión. Arreglo en `.claude/hooks/limitar_rutas.py`, a cargo del PM (ADR-011). Se relanzan con lo ya diseñado.
- F4-02, diseño a retomar:
  - grupos en orden de primera aparición;
  - números con coma decimal, sin miles, de 2 a 6 decimales;
  - columnas por período "<col> (mes)" / "<col> (año)";
  - `first`/`last` = primer/último valor con contenido;
  - `join_text` con tope de 1000 caracteres y aviso;
  - tipo de salida entero/decimal según los valores válidos.
- F4-02 (nivel 2, líder): un importe vacío **no** cuenta como inválido (igual que `ConversionResult.failed`, Fase 1). En el escenario, `not_numeric["Importe"]` = 3 ("n/a", "12,34,5", "abc"); el esperado `resumen_ventas.json` lista 4 porque incluye el vacío del cliente 7. Los tests esperan 3 inválidos + 1 vacío. QA ajusta la nota del esperado en F4-20 si hace falta.
- F4-10, diseño a retomar:
  - `_core()` perezoso, y si falta `aggregate` → `ExcelToolError`;
  - `name=None` salvo edición;
  - "Contar filas" deshabilita la columna;
  - Sumar/Promedio inválido al cambiar de columna → "Contar distintos";
  - vista previa con las primeras 200.000 filas;
  - `abrir_resumen` (Toplevel + "Guardar el resumen…" → `abrir_guardar` con `column_types`);
  - botón "Resumir…" en `ui/datos.py`.
- F4-01: "123.0" no se iguala a "123" (podría juntar códigos distintos); si un CSV trae enteros con ".0" no van a emparejar. Posible opción futura (nivel 3). Normalizar una clave de 2M usa ~120 MB extra: el join normaliza de a una columna.
- F4-01 sí pudo escribir sus archivos a pesar del problema del hook; F4-02 y F4-10 no. El bloqueo no es parejo.
- F4-02 → F4-03: el join excluye las claves vacías antes de resumir un secundario (no emparejan). Helpers públicos: `result_columns(spec)`, `group_column_name`. Resumir 2M con 2M grupos llega a 1,28 GB de pico: vigilar la memoria del join.
- F4-05: `test_fase3_reglas.py::test_los_limites_estan_solo_en_core_limits` falla por un falso positivo: encuentra "300000" dentro de "0,30000000000000004" en un docstring de `aggregate.py`. QA lo corrige en F4-20: buscar solo constantes numéricas en el AST, no texto.
- F4-05: guardar con tipos sigue en 17 s (formato de fechas y `to_csv` en `io/writer.py`); `to_date` de fecha-hora ~3 s. Posibles mejoras más adelante.
- F4-03: `prefix` va delante de todas las columnas del secundario; `output_column_order` ordena pero no quita (para quitar, `columns`); con resumen final, `rows_result` = grupos. Memoria del resultado: `memory.check_result_memory` (32 B por celda, medido en Linux; revalidar en Windows).
- QA F4-20: APROBADO CON OBSERVACIONES (1134 tests). Arreglos antes del commit: F4-A (ventana Resumir a 768 px, hecho: helper `ui/desplazable.py`), F4-B (NaT, meses cronológicos, plurales), F4-C (ceros en resúmenes sin pareja, ADR-037). QA saca la marca xfail de `test_fase4_ui.py` y ajusta lo que dependa del orden de meses o de los vacíos.
- F4-C hecho: ceros en Sumar/Contar sin pareja. Consecuencia aceptada: en un resumen final sobre esas columnas, el 0 entra en promedio, mínimo, máximo y contar distintos (ej. el promedio de "Total comprado" por región cuenta a los clientes sin compras como 0). Fallan 14 tests de QA en `test_fase4_*` que esperaban vacío → QA los ajusta.

## Estado al cerrar (2026-11-26)
- Fases 0–4 cerradas en Linux. Windows: checklist en docs/qa/fase-2.md (secciones A–G, pasos 1–50), en la PC propia del PM.
- No correr dos suites a la vez: la suite regenera `tests/fixtures/generados/` y una corrida paralela puede fallar `test_generacion_deterministica`.
- Próximo: Fase 5 (asistente para unir archivos, UI + paso Guardar ya hecho en la Fase 3 + editor de resumen ya hecho en la Fase 4).
