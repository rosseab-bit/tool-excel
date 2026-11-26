# Fase 5 — Asistente para unir archivos (cerrada en Linux 2026-11-26; Windows pendiente PM)

Fuente: PLAN.md §3.2, §3.3 (puntos 4 y 8), Fase 5 y §6 (checklist). Contratos: `docs/contratos.md` § Fase 5 (propuestos). Ya hecho: paso Guardar (F3), editor de resumen (F4), todo el core del join y recetas (F4).

## Tareas

| # | Agente | Tarea | Depende de | Estado |
|---|---|---|---|---|
| F5-01 | backend | "Achicar el resultado": `join.suggest_reductions(plan, tables)` → alternativas con su tamaño (pasar un detalle a resumen por la clave, quitar columnas pesadas), sin unir | — | Hecho (13 tests; grande en 0,9 s; suma mixtas sumables) |
| F5-02 | backend | Guardar varias tablas en un libro: `output.save_workbook` (hoja "Detalle" + "Resumen…" + "Sin coincidencia en …"), con control de tamaño por hoja y la regla de §3.3 (si el detalle es grande, resumen en un archivo y detalle en partes) | — | Hecho (12 tests; `save_workbook` + `plan_workbook`) |
| F5-03 | backend | `join.final_column_names(plan, tables)`: nombres finales de las columnas sin unir (pedido de F5-13) | F5-01 | Hecho (6 tests; misma regla que `run_join`; < 0,5 s) |
| F5-04 | backend | Bug: una columna llamada "Cantidad" o "Porcentaje" choca con las columnas de `value_counts` (rompe `suggest_split_columns` y probablemente Valores distintos). Además `profile_table` sin ruta de archivo | F5-14 | Hecho (716 tests core; `count_column`/`percent_column`; `profile_table(display_name=)`) |
| F5-10 | frontend | Base del asistente: ventana con pasos (Anterior / Siguiente, barra de pasos), estado compartido `EstadoUnion` que arma el `JoinPlan`, botón "Unir archivos" habilitado en el inicio | contrato | Hecho (140 tests ui; `ui/unir/estado.py`, `asistente.py`, `pasos/`) |
| F5-11 | frontend | Pasos 1–3: archivo principal (selector + 🔑 en la clave), columnas clave (1 o más), archivos a sumar ("+ Agregar archivo", reordenar, quitar, clave emparejada sola o a mano, completar / solo los que están en ambos, indicador ✅/⚠️ con `diagnose_join`, "🔍 Ver información" que abre el explorador) y opciones de limpieza de la clave (incluye ceros a la izquierda) | F5-10 | Hecho (17 tests; `paso1_principal.py`, `paso2_clave.py`, `paso3_sumar.py`, `widgets_archivos.py`) |
| F5-12 | frontend | Paso 4: detalle / resumen / primera fila por secundario, **sin opción elegida de entrada**, con ejemplo de sus datos ("El cliente 1045 aparece 12 veces"), editor de resumen, aviso de "más de un detalle" con "Cambiar a resumen" en un clic, aviso adelantado de tamaño | F5-10 | Hecho (12 tests; `paso4_modo.py`) |
| F5-13 | frontend | Paso 5: elegir columnas por archivo (casillas, todas/ninguna, buscador), lista final reordenable (↑/↓), indicador de tamaño con semáforo que se actualiza | F5-10 | Hecho (14 tests; `paso5_columnas.py`, `lista_casillas.py`, `lista_ordenable.py`, `tamano.py`) |
| F5-14 | frontend | Pasos 6–7: resumen final opcional (editor sobre todas las columnas), vista previa + diagnóstico, "🔍 Ver información del resultado", Guardar (una tabla con el paso Guardar, o libro con hojas extra), "achicar el resultado", "Guardar esta receta" y "Abrir una receta" desde el inicio | F5-01, F5-02, F5-11..13 | Hecho (19 tests; 202 tests ui; pasos 6-7, explorador con `tabla=`, `abrir_guardar(salida_inicial=)`) |
| F5-15 | frontend | UI usa `count_column`/`percent_column` y `profile_table(display_name=)` | F5-04 | Hecho (203 tests ui) |
| F5-16 | frontend | F5-C (variables tk sin `master=`), F5-A UI (tope de parte al no cortar grupos; al cambiar de destino se vuelve a la división sugerida), F5-B UI (soltar resultados, no repetir diagnósticos) | F5-20 | Hecho (247 tests; recorrido grande 1.443–1.674 → ~1.200 MB, estable) |
| F5-05 | backend | F5-A core (`SizeCheck.max_rows_per_part`), F5-B core (resumen final sin armar el detalle completo; menos pico en `profile_table`) | F5-20 | Hecho (730 tests core; resumen final 996 → 741 MB; `max_rows_per_part`) |
| F5-20 | qa | Escenario de 3 archivos y escenario grande hechos **desde la UI** (xvfb), checklist §6 (unir) automatizable, verificación → `docs/qa/fase-5.md` | todo | APROBADO tras arreglos (1336 tests; memoria ~1,2 GB; `docs/qa/fase-5.md`) |

Paralelo: F5-01, F5-02 y F5-10 juntos; F5-11, F5-12 y F5-13 en paralelo sobre la base de F5-10; F5-14 al final.

## Criterios de aceptación (PLAN Fase 5)
- Desde la UI, sin tocar código: las 4 variantes del escenario (detalle; resumen por cliente; por cliente + artículo; resumen final por región).
- Escenario grande en detalle: la app avisa en el paso 4 y al guardar, recomienda dividir y guarda las partes correctas, para Excel y para Google Sheets.
- Checklist §6 "Unir archivos" (detalle, resumen, general, resultados grandes) cumplido.
- Receta guardada y abierta con archivos del mes siguiente da el resultado.
- 🪟 Prueba en Windows (PM, al final de todo, ADR-038): escenario de 3 archivos con el `.exe`, "Abrir carpeta" y archivo abierto en Excel al guardar.

## Notas
- F5-10: `al_salir()` se llama también al volver atrás. Cada paso se crea una vez (no se pierde lo cargado). `estado.faltantes` / `estado.completar(nombre, fuente)` para recetas con archivos del mes nuevo. Los nombres de fuente tienen que ser únicos (lo garantiza el paso 3). `estado.py` es compartido: los pasos no lo modifican en paralelo; si necesitan algo, lo piden.
- F5-12: si la clave no se repite, el modo queda en "primera fila" sin preguntar (se reevalúa al volver al paso). El editor de resumen se abre en una ventana aparte. `tests/ui/test_unir_asistente.py::test_avanzar_y_retroceder_sin_perder_nada` asumía pasos provisorios: se ajusta en F5-14 cuando estén todos los pasos.
- F5-13: las columnas de un secundario en resumen se ven en el paso 5 pero se cambian en el paso 4 (editando el resumen); se deja así. `tamano.nombres_finales` usa una unión de 300 filas hasta que exista `final_column_names` (F5-03); después se reemplaza en F5-14.
- F5-11: ✅ desde 80 % de coincidencia. "Ver información" hoy relee el archivo (el explorador solo acepta una `Seleccion`): F5-14 agrega al explorador la opción de recibir `df` + perfil (hace falta también para "Ver información del resultado", que no tiene archivo).
- F5-14: la receta guarda el `OutputSpec` solo si se guardó con "Guardar el resultado" (un libro de varias hojas no se modela en `OutputSpec`). Resumen final con resultado > 200.000 filas: el editor trabaja sobre una muestra de 5.000 filas del principal (con aviso); QA lo prueba con el grande.
- F5-05: `profile_table` de 2,5M sigue sumando ~215 MB (columna con 2,2M valores distintos); bajarlo más requiere procesar por bloques (más lento). Se deja.
- F5-16: al cerrar Guardar se conserva el resultado unido para poder guardarlo para otro destino sin volver a unir; se suelta al volver atrás o antes de unir de nuevo. Caché de diagnósticos por firma del plan en `EstadoUnion`. Se arreglaron además 23 variables de tk sin `master=` en otras pantallas.
- QA F5-20b (para la Fase 9): con "sin cortar" la división usa el tope por parte, así que la última parte puede quedar más chica (Google Sheets: 4 × 625.000 + 20.032). Correcto pero menos parejo; se podría repartir mejor.

## Estado al cerrar (2026-11-26)
- Fases 0–5 cerradas en Linux: explorador, guardar/dividir y el asistente para unir completo con recetas.
- Windows: checklist en docs/qa/fase-2.md, secciones A–H (pasos 1–58), en la PC del PM al final (ADR-038).
- Próximo: Fase 6 (Concatenar), Fase 7 (Dividir un archivo existente), Fase 8 (empaquetado), Fase 9 (pulido).
