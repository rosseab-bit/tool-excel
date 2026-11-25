# Fase 2 — Explorador (cerrada en Linux 2026-11-25; Windows pendiente PM)

Fuente: PLAN.md §3 (intro), §3.1 y Fase 2. Contratos: `docs/contratos.md` § Fase 2 (propuestos). Notas heredadas: `docs/fases/fase-1.md` § Notas.
Ahorro de tokens (ADR-020): una tarea = un agente nuevo, con registro de avance.

## Tareas

| # | Agente | Tarea | Depende de | Estado |
|---|---|---|---|---|
| F2-01 | backend | `core/progress.py` + `Cancelado` + `core/preferences.py` + tests | — | Hecho (214 tests core; `ThrottledProgress`) |
| F2-05 | backend | pyarrow (ADR-025): dependencia, `.spec`, textos en formato compacto en `read_table`, suite completa, medir memoria/tiempo/tamaño del `.exe` antes y después | F2-01 | Hecho (340 tests; pico CSV 2M 906 MB; .exe +87 MB, ADR-028) |
| F2-06 | backend | Control de memoria antes de cargar: `platform_utils.available_memory()` + `core/memory.py` (`check_memory`) + tests | F2-05 | Hecho (25 tests; self-test exige pyarrow) |
| F2-07 | backend | `read_table` con `progress`/`cancel` (CSV por partes, ≤ 2 s) (ADR-027) | F2-05 | Hecho (lectura CSV 2M: 3,5 s, 571 MB pico; cancela por bloques) |
| F2-15 | frontend | Worker: al cancelar, cerrar enseguida y descartar el resultado (ADR-027) | F2-10 | Hecho (45 tests ui; `on_cancel`) |
| F2-08 | backend | `excel_tool.spec`: la poda de F2-05 rompía pyarrow en el ejecutable (lo detectó el self-test de F2-06); corregir y volver a medir el tamaño | F2-06 | Hecho (SELF-TEST OK con pyarrow real; carpeta 206 MB, tar.gz 76 MB) |
| F2-02 | backend | `profile_table` (TableProfile/ColumnProfile) + tests contra `perfil.json` y 2M filas | F2-05 | Hecho (28 tests; CSV 2M leer+perfil 34 s y 1,08 GB pico; xlsx 500k 18 s, 765 MB) |
| F2-03 | backend | `value_counts`, `rows_with_value`, `filter_rows`, `sort_rows`, `get_page` en `core/explore.py` (ADR-024) + tests | F2-05 | Hecho (35 tests; peor caso +412 MB; filtro/orden numérico 2M ~5 s) |
| F2-04 | backend | `export_profile`, `export_distincts` (en `core/profile.py`) + tests; además F2-09: `cancel` en `detect_column_type` y sacar el hilo auxiliar de `profile_table` | F2-02, F2-03 | Hecho (informe xlsx + lista de distintos; `detect_column_type` con `cancel`; CSV 2M 32,8 s, 1.122 MB) |
| F2-10 | frontend | `ui/worker.py` (ventana de progreso, cancelar, errores) + tests | F2-01 | Hecho (18 tests ui) |
| F2-11 | frontend | (después de F2-10, usa el worker) Pantalla de inicio (4 botones grandes, 3 deshabilitados) + selector de archivo (hoja / separador / codificación / fila de encabezado + preview 10 filas, recuerda la última carpeta) | F2-01 | Hecho (26 tests ui; `ui/selector.py`; usa doble de `check_memory` hasta F2-06) |
| F2-12 | frontend | Grilla paginada reutilizable (`ttk.Treeview`, 1.000 por página, N.º de fila, orden por encabezado, ocultar columnas) | — | Hecho (`ui/grilla.py`, proveedor de páginas base 0; menú de clic derecho mirar a ojo en Windows) |
| F2-13 | frontend | Explorador: pestañas General y Columnas | F2-02, F2-10, F2-11 | Hecho contra el contrato (43 tests ui; `ui/explorador.py`, `ui/formato.py`); falta probar con el core real |
| F2-14 | frontend | Explorador: pestañas Valores distintos y Datos + exportar informe y lista de distintos | F2-03, F2-04, F2-12, F2-13 | Hecho (18 tests; `ui/distintos.py`, `ui/datos.py`, `ui/exportar.py`) |
| F2-16 | frontend | Test de UI que falla en la suite completa (estado compartido entre tests) + limpieza común en `tests/ui/conftest.py` | F2-14 | Hecho (63 tests ui, ~18 s, 5 corridas y orden inverso) |
| F2-17 | backend | Test de rendimiento de tipos por el camino real (`read_table`) y tope = doble del objetivo | F2-05 | Hecho (`detect_types` 2M por `read_table`: 14,5 s; aviso > 15 s, falla > 30 s) |
| F2-20 | qa | Integración (core + UI con xvfb) + verificación → `docs/qa/fase-2.md` | todo | APROBADO CON OBSERVACIONES (564 tests; `docs/qa/fase-2.md`) |

Paralelo: backend F2-01 → (F2-02 ‖ F2-03 en dos agentes) → F2-04. Frontend F2-12 arranca ya; F2-10 y F2-11 apenas está F2-01 (o contra el contrato). F2-13/F2-14 cuando el core que usan esté.

Orden actualizado (ADR-025): F2-05 → (F2-02 ‖ F2-03 ‖ F2-06) → F2-04. El selector de archivo (F2-11) usa `check_memory` antes de cargar y muestra el cartel.

## Criterios de aceptación (PLAN Fase 2 + contrato)
- Core contra `perfil.json`: filas, columnas, vacíos, filas vacías, repetidas, distintos, tipos, posible clave, variantes "Norte"/"NORTE"/"Norte ".
- Desde la UI (xvfb y a mano): explorar "para explorar" (CSV y xlsx), un xlsx con varias hojas cambiando de hoja, el CSV de 2M filas sin colgarse, con progreso y cancelable (responde en ≤ 2 s).
- Perfil completo del CSV de 2M (leer + analizar) en ≤ 60 s en esta máquina; memoria pico medida y reportada.
- La grilla nunca carga más de una página; valores distintos solo a pedido.
- Textos sin jerga técnica; errores amigables; nada de tracebacks en pantalla.
- 🪟 Cierre con prueba en Windows (ADR-009): `dev.bat -Tests`, explorar un CSV exportado por Excel de Windows, compilar el `.exe` (y lo pendiente de las Fases 0 y 1).

## Notas
- F2-11: aviso de pytest `main thread is not in main loop` (un `__del__` de variable tk en `test_ver_informacion_carga_archivo`). Inofensivo según frontend; QA lo revisa en F2-20.
- F2-11: un `messagebox` sin reemplazar en tests bajo xvfb provoca segfault: los tests de UI siempre lo reemplazan.
- F2-05: `test_rendimiento_2m_filas` (`detect_types` ≤ 15 s) dio 15,6 s una vez con la máquina cargada; ahora ~14 s. QA revisa el margen en F2-20.
- F2-05: el self-test del `.exe` no comprueba que pyarrow esté dentro; se agrega en F2-06.
- F2-05/F2-06: el tamaño de 187 MB de ADR-028 se midió con pyarrow roto dentro del ejecutable (pandas volvía en silencio a textos de Python). F2-08 lo vuelve a medir.
- `test_rendimiento_2m_filas` volvió a pasarse (15,27 s) con varios agentes usando la CPU a la vez.
- F2-08, Windows: listar `dist\ExcelJoinTool\_internal\pyarrow\*.dll`, confirmar `SELF-TEST OK` en `build_windows.bat`; si falla por DLL, revisar el filtro del `.spec` (no sacar nada que enlace `arrow_python.dll`).
- F2-03: `filter_rows`/`sort_rows` no tienen `cancel`; sobre columnas numéricas con 2M de valores distintos tardan ~5 s → la UI los corre en el worker.
- F2-03: el orden alfabético usa el orden de pyarrow (Unicode): "Ñ" y las vocales con tilde quedan después de la Z. Revisar en la Fase 9 si molesta.
- F2-02: mientras `detect_column_type` no acepte `cancel`, `profile_table` corre la detección en un hilo auxiliar y lo abandona al cancelar (sigue gastando CPU unos segundos). Se arregla en F2-04/F2-09.
- F2-04: los textos del informe (tipos, marcas) están duplicados en `core/profile.py` y `ui/formato.py` (el core no puede importar la UI). Si cambian en un lado, cambiarlos en el otro.
- F2-04: `profile_table` y `detect_column_type` factorizan cada columna dos veces; se cumple la meta igual. Optimización posible más adelante.
- F2-14: la tabla de distintos muestra hasta 5.000 filas (aviso: exportar para ver todo). Filtrar y ordenar no se pueden cancelar (el core no lo permite); la ventana aparece si tardan más de 0,3 s.
- F2-16: la falla venía de `tk.Variable.__del__` ejecutado por el recolector de basura dentro de un hilo del worker sin `mainloop` (solo pasa en tests). En la app hay `mainloop` y Tcl lo deriva al hilo principal. Riesgo residual: al cerrar la app con una operación en curso. QA lo prueba a mano (Linux y Windows): cerrar la ventana durante un análisis largo.
- QA F2-20 (para fases siguientes): OBS-1 el hilo de lectura del CSV tarda 2,3 s en terminar tras cancelar (la UI se cierra en < 0,2 s); OBS-2 no hay fixture `.ods`; OBS-3 "N bytes" en archivos < 1 KB; OBS-8 el perfil de 2M tarda ~29 s (la mitad del margen): vigilar en PCs lentas.
- Prueba de Windows: seguir el checklist de `docs/qa/fase-2.md` (incluye lo pendiente de las Fases 0 y 1).
