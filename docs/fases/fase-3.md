# Fase 3 — Salida: límites, división y guardado (cerrada en Linux 2026-11-25; Windows pendiente PM)

Fuente: PLAN.md §1.2, §2 (modelo de salida), §3.3 y Fase 3. Notas heredadas: `docs/fases/fase-1.md` y `fase-2.md` § Notas. ADR-023 pasó acá "Exportar lo filtrado" y el aviso de Google Sheets.

## Tareas

| # | Agente | Tarea | Depende de | Estado |
|---|---|---|---|---|
| F3-01 | backend | `core/limits.py`: carga de límites (`limites_default.json` + `limites.json` junto al ejecutable), `SizeEstimate`, `estimate`, `check`, `suggest_split_columns` (usa `explore.value_counts`) | — | Hecho (2,3M → 🔴 Excel, 8 partes; estimación de MB: CSV −0,1 %, xlsx +4,9 %) |
| F3-02 | backend | `core/split.py`: `split_dataframe` (modos `rows`, `rows_keep_groups`, `by_value` con subdivisión, `sheets`) como iterador; nombres válidos en Windows vía `platform_utils` | — | Hecho (24 tests; 2M por valor 0,7 s, +228 MB) |
| F3-03 | backend | `core/output.py`: `save_result` (un archivo / varias hojas / partes en carpeta + `_indice.xlsx`), escribe cada parte a medida, verifica suma de filas, progreso y cancelar; xlsx con encabezado congelado y filtro; `warnings` en el resultado | F3-01, F3-02 | Hecho (60 tests; 2M en 7 partes: CSV 4,4 s / 612 MB, xlsx 232 s / 791 MB con openpyxl) |
| F3-04 | backend | Medición (sin agregar dependencia): escribir xlsx de 500k con openpyxl vs `xlsxwriter` (tiempo, memoria, tamaño del `.exe`) → decisión del PM (ADR-015) | — | Hecho: xlsxwriter 1,6–1,8× más rápido (500k con tipos 79 → 45 s; 300k 47 → 26,5 s), −150/−240 MB con tipos, mismo contenido, +0,5 MB el .exe → decide el PM |
| F3-06 | backend | Escritor de .xlsx con `xlsxwriter` (ADR-033): dependencia, `.spec`, `writer.py`, casos borde ("=", caracteres de control, varias hojas, encabezados vacíos, fechas < 1900, archivo bloqueado) | F3-03 | Hecho (791 tests; 500k con tipos 79 → 47 s; 2M en 7 partes xlsx 232 → 136 s; SELF-TEST OK, 206 MB / 73 MB) |
| F3-05 | backend | `split_file` para CSV por partes (sin cargar todo en memoria); Excel por partes queda para la Fase 7 | F3-02 | Hecho (34 tests; CSV 2M en partes: 8 s, 217 MB pico) |
| F3-10 | frontend | Paso "Guardar" reutilizable (§3.3): destino, semáforo, opciones de división con vista previa de partes, progreso, resultado con "Abrir carpeta" | F3-01..03 (contra contrato) | Hecho contra el contrato (23 tests; `ui/guardar.py`); falta prueba con `save_result` real |
| F3-11 | frontend | Explorador: "Exportar lo filtrado" con el paso Guardar + semáforo Excel / Google Sheets en la pestaña General | F3-10 | Hecho contra el contrato (12 tests; Canvas con scroll en Guardar, semáforo en General) |
| F3-20a | qa | Fixtures de la Fase 3 (nombres raros, 400k, 1,6M con región > 1,1M) + 18 tests de integración de limits/split | F3-01, F3-02 | Hecho (197 tests integración; sin bugs) |
| F3-20 | qa | Fixtures (valores con `/`, `CON`, `Norte/Sur`; resultado de 400k filas; región grande) + integración + verificación → `docs/qa/fase-3.md` | todo | APROBADO CON OBSERVACIONES (891 tests; `docs/qa/fase-3.md`) |

Paralelo: F3-01, F3-02 y F3-04 arrancan juntos; F3-10 contra el contrato apenas esté aprobado.

## Criterios de aceptación (PLAN Fase 3)
- Escenario grande: `check` da 🔴 para Excel y la cantidad de partes sugerida es correcta. 400k filas: 🟡 Excel, 🟢 "Solo guardar".
- Ninguna parte supera el límite; `rows_keep_groups` nunca corta un grupo; `by_value` subdivide la región grande; suma de filas de las partes = total; el índice lista bien las partes.
- Cambiar un límite en `limites.json` cambia la recomendación. Ningún límite hardcodeado fuera de `core/limits.py`.
- Valores con caracteres prohibidos en Windows (`Norte/Sur`, `CON`) dan nombres de archivo válidos.
- Memoria: dividir y guardar el CSV de 2M no pasa de 1,5 GB de pico (ADR-023).
- (Si se aprueba F3-10/F3-11) desde el explorador se exporta lo filtrado, con semáforo y división.

## Notas
- F3-02 → F3-03: con 0 filas `plan_split` da 0 partes (save_result tiene que escribir igual un archivo con solo encabezado o no dividir). En `rows_keep_groups`, un grupo más grande que el límite va entero en su parte: `save_result` lo avisa en `warnings`. En `sheets` el nombre ya es la hoja final (no volver a pasarlo por `nombre_de_hoja`). Controlar el largo de la ruta completa (260 caracteres en Windows).
- F3-10 → F3-11: `abrir_guardar` suma `column_types: dict | None = None` y lo pasa al `OutputSpec`; el explorador le pasa los tipos detectados del perfil, así los números y fechas salen como número y fecha en Excel (cierra lo que ADR-021 dejó para la Fase 3). La ventana es alta y no tiene barra de desplazamiento: agregarla o ajustarla para pantallas de 768 px de alto.
- F3-10: no cubre "Achicar el resultado" ni guardar la división en la receta (§3.3 puntos 4 y 6): van con el asistente de unir (Fase 5) y las recetas.
- F3-01: `limits.estimate` usa `df.to_csv` sobre un buffer en memoria para medir la muestra; va contra la regla de CLAUDE.md (nada de `to_csv` fuera de `core/io`). F3-03 agrega un helper en `core/io` y `limits` lo usa.
- F3-01: la estimación de xlsx se calibró escribiendo texto; con `column_types` el tamaño real puede variar.
- F3-20a: `by_value` distingue mayúsculas ("Sur" y "sur" son dos partes; el segundo nombre lleva " (2)"). Se deja así (valor exacto, igual que el resto de la app); si molesta, se puede ofrecer "juntar escrituras parecidas" también al dividir (Fase 9).
- F3-03: escritura atómica (`.tmp` + `os.replace`); en Windows probar a mano un archivo abierto en Excel, una ruta larga y el antivirus reteniendo el `.tmp`.
- F3-05 → Fase 7 (asistente Dividir): `split_file` admite hasta 50 valores en `by_value` y 50 partes en `rows_keep_groups`; en `rows_keep_groups` cada parte conserva el orden original del archivo (distinto de `split_dataframe`, que agrupa). `output.py` importa privados de `split.py` (`_Nombres`, `_limite`): hacerlos públicos en una limpieza. `CSV_ESTILOS` está duplicado en `csv.py` y `writer.py` (un test avisa si divergen).
- F3-06: fechas anteriores a 1900 se guardan como texto con aviso; celdas de más de 32.767 caracteres se recortan con aviso. `writer.py` usa `ws._opt_close()` (privado de xlsxwriter) al cancelar: revisar si se actualiza xlsxwriter. openpyxl queda solo en el extra `dev` (tests y `make_fixtures`).
- QA F3-20b (para fases siguientes): OBS-1 guardar con tipos es más lento (2M en CSV 29 s; en xlsx 173 s): la UI siempre pasa tipos; ver si conviene optimizar la conversión. OBS-2 en CSV con tipos la fecha-hora sale con segundos. OBS-4 `SplitSpec("sheets")` ignora `value_columns` sin avisar. OBS-6 Cancelar desde la ventana de progreso al guardar: probar a mano. OBS-7 la suite completa tarda ~6 min.

## Estado al pausar (2026-11-25)
- Fases 0–3 cerradas en Linux y commiteadas. Windows: instalación y 545 tests confirmados; falta compilar y probar el `.exe` y los checklists (docs/qa/fase-2.md, secciones A–F, pasos 1–43) en la PC propia del PM.
- Pendiente del PM: agregar `"docs/progreso/",` a backend, frontend y qa en `.claude/hooks/limitar_rutas.py`.
- Próximo: Fase 4 (core de resúmenes y unir varios archivos, PLAN.md). El líder presenta el plan y espera el OK antes de despachar.
