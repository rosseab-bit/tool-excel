# QA Fase 4 (F4-20): Core de resúmenes y unir varios archivos

## Resultado: APROBADO (F4-A, F4-B y F4-C resueltos)
## Suite (re-verificación F4-20b): pytest 1152 passed / 0 failed / 2 skipped / 0 xfailed (`xvfb-run -a scripts/dev.sh --tests`) · ruff OK (check y format) · git status limpio de generados: SÍ

## Re-verificación F4-20b (después de F4-A, F4-B y F4-C)
- **F4-A resuelto:** `test_la_ventana_de_resumir_entra_en_una_pantalla_de_768` daba XPASS; saqué la marca `xfail` y pasa como test normal.
- **F4-B resuelto:** meses en orden cronológico y avisos en singular/plural. Ajusté los textos esperados (`test_fase4_join.py`, `test_fase4_ui.py`:
  "Hay 1 celda vacía de Importe: no se suma", "3 celdas con contenido que no se pudieron leer"). Test nuevo
  `test_ventas_agrupadas_por_mes_salen_en_orden_cronologico` (CSV y xlsx): meses del escenario en orden calendario contra un recuento
  independiente, "(sin fecha)"/"(vacío)" al final, y un caso que cruza de año (nov 2023, dic 2023, ene 2024, feb 2024, "(sin fecha)").
- **F4-C anotado (ADR-037):** sin pareja, Sumar/Contar filas/Contar distintos dan "0"; el resto queda vacío. Ajusté 12 tests (CSV y xlsx):
  clientes 5, 17 y 28 con "Total comprado" y "Compras" en 0 (y "Última compra" vacía), clientes sin cobranzas con "Total cobrado" y "Cobros" en 0,
  y en el xlsx guardado la celda es el número 0, no vacía. Los conteos "5 clientes con cobro sin la opción / 27 con la opción" ahora cuentan
  los que tienen un total distinto de 0. Los demás valores (sumas, filas, 88 filas del diagnóstico) no cambiaron.
- **Esperados:** `resumen_final_por_region.json` NO necesita cambios: las sumas por región no cambian con el 0 y "Clientes" es contar
  distintos de la clave del principal (no del secundario). Su `_nota` ya decía que los clientes sin ventas suman 0. No hay ningún
  esperado que use promedio/mínimo/máximo sobre esas columnas, así que no hubo que tocar el generador.
- Ningún fallo quedó sin explicar con estos tres cambios.
- Nota de proceso: una corrida de la suite dio un fallo de `test_generacion_deterministica` mientras otra corrida corría en paralelo
  (las dos regeneran `tests/fixtures/generados/`); sola pasa (2 veces) y la corrida limpia dio 1152 passed. No es un bug de producto.

Corridas: 365 s (6 min 4 s), 364,3 s y 364,6 s (la del build). Ninguna falla intermitente.
Los 2 skips son los de siempre (`test_fixtures.py:185` y `test_fase2_explorador.py:282`, sin `.ods`). El xfailed es el
BUG F4-A (abajo), marcado `strict`: cuando se arregle el test pasa a fallar hasta que se saque la marca.
`scripts/build_linux.sh`: **SELF-TEST OK**, carpeta 206 MB, `.tar.gz` 73 MB, SHA256 `2fb9beef...0ad73`, 406 s con la suite adentro.
`git status`: sin generados sin ignorar (`dist/`, `build/` y `tests/fixtures/generados/` siguen ignorados; lo nuevo es código,
`docs/` y un esperado de 1 archivo).

## Tests nuevos y cambios (solo en mis carpetas)
- `tests/integracion/test_fase4_join.py` (45 tests, todos con el core real; las vistas de `csv` y `xlsx` por parametrización: las
  ventas como `ventas.csv` o guardadas como `.xlsx`). Recorren `read_table` -> `diagnose_join` -> `run_join` -> `save_result`
  -> relectura (xlsx y CSV; se compara contenido celda por celda, números como `Decimal`). Esperados: los JSON del generador o
  recuentos independientes con `csv`, `Decimal` y openpyxl (nunca con el producto).
- `tests/integracion/test_fase4_ui.py` (3 tests + 1 xfail): explorador -> Datos -> "Resumir…" con xvfb y el core real.
- `tests/integracion/test_fase4_reglas.py` (14 tests): reglas de CLAUDE.md sobre el AST de `src/`.
- `tests/integracion/medicion_fase4.py` (no es un test; lo usa el test del escenario grande) y
  `tests/integracion/receta_cruzada.py` (no es un test: la receta Linux <-> Windows para el PM; probada acá en Linux: carpeta movida y datos en otra carpeta).
- Arreglo pendiente mío: `test_fase3_reglas.py::test_los_limites_estan_solo_en_core_limits` ya no da el falso positivo.
  Ahora mira solo constantes numéricas del AST (`ast.Constant` int/float), no texto ni docstrings ni comentarios; un test nuevo
  (`test_el_detector_de_limites_mira_solo_numeros_del_ast`) prueba que detecta `1_048_576`, `300000.0`, etc., y que ignora el
  `"0,30000000000000004"` del docstring de `aggregate.py`.
- `scripts/make_fixtures.py` + `esperados/escenario_3/resumen_ventas.json`: distingue vacío de inválido. `importes_invalidos`
  ahora cuenta solo los que traen algo ilegible (3: "n/a", "abc", "12,34,5"), `importes_vacios` el vacío (1), y
  `celdas_invalidas` sigue listando las 4 con `"vacia": true/false`; la `_nota` lo explica. `test_escenarios.py` ajustado.
  Regeneré todos los fixtures (incluidos los grandes) con el script: de lo versionado cambió solo ese JSON.
- `docs/qa/fase-2.md`: sección G (pasos 44 a 50) del checklist de Windows, sin duplicar lo anterior.

## Criterios de aceptación de la fase (`docs/fases/fase-4.md`)
- [x] **Detalle de ventas + resumen de cobranzas: una fila por compra con el total cobrado.** 88 filas = 85 compras de
  clientes + 3 clientes sin compras (las 2 ventas de IDs que no son clientes no entran); por cliente: filas = compras
  del esperado (o 1), `Total cobrado` = `suma_monto` del esperado y `Cobros` = `filas`, vacío donde no hay cobranzas; los importes
  originales salen todos (recuento con `csv`); no hay aviso de varios detalles. Guardado en xlsx y CSV y releído igual.
  (`test_detalle_de_ventas_mas_resumen_de_cobranzas`).
- [x] **Resumen de ventas por cliente: mismas filas que clientes, `Total comprado` igual al esperado.** 30 filas en el orden del
  principal, `Total comprado`, `Compras`, `Artículos distintos` y `Última compra` iguales al esperado para los 27 clientes con
  compras; los 3 sin compras (5, 17, 28) quedan vacíos; tipos `decimal` y `fecha`.
- [x] **Cliente + artículo.** Una fila por par cliente/artículo (recuento con `csv` + `Decimal`) más una por cliente sin compras;
  suma y cantidad de cada par iguales; filas = estimación de `diagnose_join` (ADR-036).
- [x] **Resumen final por Región.** 5 filas; `Total vendido` por región = `total_importe_por_region`; clientes por región =
  `clientes_por_region`; y total por región + lo excluido (IDs 98/99, 8.155,88) = suma de todos los importes válidos del CSV.
- [x] **Dos archivos en detalle: aviso + filas estimadas correctas.** `many_details` True, aviso con "detalle" en el
  diagnóstico y en el resultado; `rows_result` = `filas_resultado_ventas_detalle_y_cobranzas_detalle` del esperado, con la opción de
  ceros y sin ella; filas por cliente = filas de ventas x filas de cobranzas.
- [x] **Ceros a la izquierda (ADR-034).** Sin la opción: cobranzas empareja 5 claves (`solo_recortar_espacios`), `keys_matched_without_zeros`
  = 27 y el aviso "si sacás los ceros a la izquierda, coinciden 27 claves en vez de 5" (una sola vez). Con la opción: 27
  (`sin_ceros_a_la_izquierda`), sin sugerencia, sin pareja = [3, 17, 22], huérfanas = 1 clave (las 2 filas de "00077", en `unmatched`).
  Totales: con la opción los 27 igualan `suma_monto`; sin ella son parciales y se recuentan con openpyxl (solo las filas con la
  clave exacta). Detalle sin la opción: filas = `solo_recortar_espacios`.
- [x] **Importes inválidos se reportan y no rompen la suma.** `not_numeric == {"Importe": 3}`; avisos "3 celda(s) con contenido
  que no se pudo leer" y "1 celda(s) vacía(s)" por separado (también en `run_join` y en la UI: "3 celdas de Importe no son
  números y no se sumaron"); los totales siguen siendo los del esperado. El esperado distingue ahora 3 inválidos + 1 vacío.
- [x] **Sin detalle y sin resumen final, mismas filas que el principal (ADR-036).** Resumen + resumen, resumen + primera, primera +
  primera y con `prefix`: 30 filas, mismo orden de IDs, diagnóstico = real. Los casos que cambian las filas a propósito (cliente +
  artículo, `inner` en resumen = 27, `inner` en detalle = compras de clientes) coinciden con la estimación. El control corta de
  verdad: simulando un error (una fila de más con la estimación coincidiendo) `run_join` falla con "no tiene las mismas filas ...
  No se guardó nada".
- [x] **Receta guardada y cargada: mismo resultado, incluida la división.** Plan con 2 resúmenes, orden de columnas y salida
  CSV en partes de 12: `cargada.plan == plan`, `missing == []`, mismo `df` y mismos tipos, y los 3 archivos de partes son **iguales
  byte a byte**. Rutas con `/` y relativas (sin `\` ni rutas absolutas del equipo); la carpeta entera se copia a otro lugar y
  carga igual. **Receta con rutas de Windows cargada en Linux** (`D:/Mis datos/...`, `C:\Users\Ana Pérez\...`, `\\servidor\...`):
  carga sin error, `missing` = los 3 nombres, avisa, conserva resumen y división, y al elegir los archivos de este mes corre igual.
  Recetas rotas, de otra versión o inexistentes: error amigable ("receta ...").
- [x] **Escenario grande (200.000 clientes x 2.500.000 ventas, proceso aparte con `/usr/bin/time -v`; incluye leer, `diagnose_join` y `run_join`):**

  | Modo | Filas estimadas | Filas reales | Leer | Diagnóstico | Unir | Pico |
  |---|---|---|---|---|---|---|
  | detalle (7 columnas) | 2.519.998 | 2.519.998 (= `grandes.json`) | 2,2 s | 1,5 s | 4,9 s | **889 MB** |
  | resumen por cliente (5 columnas) | 200.000 | 200.000 | 2,2 s | 1,6 s | 4,6 s | **634 MB** |

  Dentro de los 1,5 GB (ADR-023) con 650 MB de margen en el peor caso; `memoria_ok` True. Se mide en cada corrida de la suite
  (`test_escenario_grande_...`, 2 tests, ~20 s).
- [x] **Guardar en xlsx con `column_types` del join (openpyxl):** ID cliente como número, `Total comprado` como decimal real (float
  igual al esperado), `Compras` como número, `Última compra` como **fecha real** (`data_type "d"`, formato `dd/mm/yyyy`), nombres como
  texto, encabezado congelado en `A2`; los clientes sin compras quedan con celdas vacías. CSV: coma decimal y mismos valores.

## UI con el core real y xvfb (`test_fase4_ui.py`)
- [x] `ventas.csv` en el explorador -> Datos -> "Resumir…" (texto exacto) -> "ID cliente" + Sumar Importe + Contar filas -> frase "Una fila
  por cada ID cliente" -> "Ver cómo queda": 29 filas (27 clientes + IDs 98 y 99) y cada total/cantidad igual al esperado; aparece
  "3 celdas de Importe no son números y no se sumaron" y el vacío en los avisos del core; la grilla muestra el resumen.
- [x] "Guardar el resumen…" -> `abrir_guardar` (nombre `ventas_resumen`, 29 filas) -> guardar -> releer: mismos IDs y totales; con
  openpyxl, los totales son decimales reales y las cantidades números. Sin carteles de error.
- [x] Agrupar "Fecha" por mes: la columna ofrece "por día / por mes / por año"; sale `Fecha (mes)` con 2024-01 a 2024-06,
  cantidad y suma por mes iguales al recuento con `csv` (87 filas en total).
- [x] Con un filtro en Datos ("Artículo es igual a Tornillo") el resumen usa solo las filas filtradas (sin grupos: una fila con
  el total). Sin nada elegido, "Ver cómo queda" avisa y no calcula.
- [x] Sin jerga en ningún texto de la ventana (ni etiquetas ni listas desplegables).

## Reglas de CLAUDE.md (`test_fase4_reglas.py` + `test_fase3_reglas.py`)
- [x] Los resúmenes pasan por `aggregate.summarize`: `join.py` no tiene `groupby`/`agg`/`pivot_table`/`resample` y llama a `summarize` en
  3 lugares (secundario en resumen, conteo del final y resumen final); `groupby` solo existe en `core/aggregate.py` y en
  `core/output.py` (para repartir filas al dividir); la UI no calcula, pide `core.summarize(` por `_core()`.
- [x] Nada de tkinter en el core (todo `core/`, y los 6 módulos de la fase por nombre), ni `excel_tool.ui` importado.
- [x] Textos de la UI sin jerga (ampliada con "group by", "aggregate", "aggregation", "join", "merge"), también las frases de error y de
  aviso de `aggregate`, `join`, `normalize` y `recipe`; el detector se prueba con cada palabra.
- [x] Límites solo en `core/limits.py` (test arreglado; además se mira por AST `aggregate`, `join`, `normalize`, `recipe` y `memory`).
- [x] Cada función nueva del core con test de CSV y de xlsx: la integración corre todo con las dos variantes.

## Bugs (ordenados por gravedad)
**[RESUELTO] BUG F4-A (media, UI): la ventana "Resumir" se agranda más que una pantalla de 768 px.** Se abre con 652 px, pero al agregar dos
cálculos y apretar "Ver cómo queda" (con el aviso de celdas que no se sumaron) crece sola a **790 px** de alto (`winfo_reqheight`) y
el pie con "Guardar el resumen…" y "Cerrar" queda más abajo del borde de una notebook de 1366x768 (peor con la barra de tareas).
Reproducir: `xvfb-run -a .venv/bin/python -m pytest tests/integracion/test_fase4_ui.py --runxfail -k 768`, o a mano: abrir
`tests/fixtures/generados/escenario_3/ventas.csv` (punto y coma, Windows), Datos, "Resumir…", tildar "ID cliente", dos cálculos,
"Ver cómo queda". Esperado: la ventana tiene un tope de alto (por ejemplo, una parte con desplazamiento como la del paso
Guardar, Fase 3) y el pie siempre visible. El test está marcado `xfail(strict=True)` y lo dice. Dueño: frontend.

**[RESUELTO] BUG F4-B (baja, core): `DeprecationWarning` de NumPy en `core/aggregate.py:426`.** `np.full(ng, np.datetime64("NaT"), dtype="datetime64[ns]")`
usa una unidad genérica ("The 'generic' unit for NumPy timedelta is deprecated, and will raise an error in the future"). Aparece en
`tests/core/test_aggregate.py::test_first_last_numero_y_fecha_con_formato` y hoy no rompe nada, pero con un NumPy más nuevo
`first`/`last` sobre fechas va a fallar. Arreglo: `np.datetime64("NaT", "ns")`. Dueño: backend.

## Observaciones (no son bugs; para el líder)
1. **Los meses salen en orden de aparición, no ordenados** (diseño de las Notas de F4-02): con `ventas.csv` la vista previa muestra
   2024-01, 2024-02, 2024-06, 2024-05, 2024-03, 2024-04. Un usuario de negocio probablemente espera orden cronológico; decide el
   líder/PM (el texto `2024-03` ya ordena bien, ADR-034).
2. **Clientes sin compras: las columnas del resumen quedan vacías, no en 0** (también `Compras`). Coincide con el esperado ("los
   clientes sin ventas suman 0" solo en el resumen final por Región). Conviene que la UI de la Fase 5 lo explique.
3. El aviso de importes vacíos sale como "1 celda(s) vacía(s)" (con el "(s)"): funciona, pero es poco natural en castellano. Frase
   para pulir en la Fase 5.
4. `diagnose_join` solo advierte de claves repetidas en el principal cuando hay un resumen (probado a mano: con 2 filas de cliente
   "1" el aviso sale y dice que los totales cuentan de más). No encontré caso sin aviso.
5. La memoria medida es el pico del proceso entero (leer las dos tablas + diagnóstico + unir), sin guardar. Guardar no lo sube
   (Fase 3). El valor de 32 B/celda sigue sin validarse en Windows: pasos 48 del checklist.
6. Todavía no hay pantalla para unir ni para recetas: las recetas y el join se prueban por el core. La Fase 5 tiene que agregar los
   tests de UI equivalentes.
7. "123.0" no se iguala con "123" (nota de F4-01): no lo probé; sigue como opción futura (nivel 3).

## Sondeos extra con el core (a mano, sin bugs): clave repetida en el principal con resumen (avisa); clave vacía en el principal
(queda vacía, no empareja); columna inexistente en el secundario (error amigable "No encuentro la columna «zz» en «s»"); nombres
repetidos entre archivos (`v`, `v (s)`); clave compuesta (2 columnas); fechas ilegibles al agrupar por mes (grupo "(sin fecha)" y
aviso; vacíos en "(vacío)"); `group_by` vacío (una fila) y sin nada (error amigable); tabla vacía (0 filas); número con 'e' ("1e3")
queda como no numérico.

## Para probar en Windows (PM)
Agregado en `docs/qa/fase-2.md`, sección **G** (pasos 44 a 50): Resumir con Excel real, receta de Linux abierta en Windows y al revés
(`tests/integracion/receta_cruzada.py crear|verificar`), rutas en otra unidad o de red (y el caso "falta el archivo"), memoria del
join en Windows (32 B por celda, a revalidar con `medicion_fase4.py`) y el editor de resumen en 1366x768 (BUG F4-A).

## Veredicto
APROBADO. La Fase 4 cumple todos los criterios de aceptación, ejecutados de verdad con el core real en CSV y xlsx. Los dos defectos de mi
informe (F4-A ventana de 768 px, F4-B orden de meses / plurales) están resueltos y verificados, y F4-C (ceros sin pareja, ADR-037) está
cubierto por los tests. Sin observaciones abiertas del core. Pendiente solo lo de Windows (PM): memoria del join y el editor de resumen en 1366x768.
