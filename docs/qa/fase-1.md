# QA Fase 1 (F1-20): Core de I/O y tipos

## Resultado: APROBADO
## Suite: pytest 309 passed / 0 failed / 1 skipped · ruff OK · git status limpio de generados: SÍ
(El skip es `tests/integracion/test_fixtures.py:185`: cubre el caso "faltan los grandes" y los grandes existen.
Corrida con los grandes generados. No se regeneraron fixtures. `--self-test` devuelve 0.)

## Criterios de aceptación de la fase
- [x] Cada fixture se lee con `read_table` con tildes y ñ correctas y sin perder ceros: `test_todos_los_fixtures_se_leen_sin_mojibake`, `test_tildes_enie_y_ceros_intactos` (xls, xlsb, xlsx, csv coma/`;` cp1252/tab/BOM). Todas las celdas son `str`.
- [x] Filas basura descartadas y encabezado igual a `dificiles.json` (`con_basura.csv` y `.xlsx`: fila 4, 6 filas de datos).
- [x] Valores y errores según `fase1.json`: excel_tipos xlsx y xls (enteros sin ".0", fechas, booleanos, fórmula, error = vacío, combinadas), hoja "Con vacias", pipe, estilo_punto, solo_encabezado; `vacio.csv` y `danado.xlsx` dan `ArchivoIlegible`.
- [x] `list_sheets`: varias_hojas.xlsx (Clientes 12x4, Ventas 25x6, Resumen año 2x2), excel_tipos.xlsx (Tipos 4x8, Con vacias 2x2), CSV (una entrada) y el xlsx de 500k (500000 filas).
- [x] Ida y vuelta: 8 fixtures x 4 formatos (xlsx, csv_excel, csv_google, csv_standard) leen igual; `write_sheets` ida y vuelta; escritura con tipos informa las celdas no convertibles.
- [x] Tipos de `para_explorar` (csv y xlsx) coinciden con `perfil.json`, incluida la mixta (`entero` 158 / `texto` 38) y la cantidad de vacíos; los tipos de `fase1.json` también.
- [x] Conversiones con celdas no convertibles reportadas (escenario 3: inválidas = `resumen_ventas.json`, sumas por cliente exactas, última fecha por cliente; `to_bool`, `to_date`).
- [x] Rendimiento (falla solo si pasa del doble):
  - `read_table` xlsx 500k: 8,5 s (objetivo 30)
  - `list_sheets` xlsx 500k: 7,7 s
  - `read_table` CSV 2M: 3,2 s (objetivo 20)
  - `detect_types` CSV 2M: 14,7 s (objetivo 15; **justo en el límite**, ver Observaciones)
- [x] Reglas (grep): sin `read_csv/read_excel/to_csv/to_excel/ExcelWriter` fuera de `core/io`; sin tkinter en `core`; lectura con `dtype=str`/`object`; todo `open` con `encoding` (la única excepción es `path.open("rb")` binario en `core/io/csv.py:26`, legítima); los límites 1048576/16384/20000000 solo están en `limites_default.json`; los `except` de `core/io` envuelven en `core/errors` (probado: archivo inexistente, extensión .pdf, bytes ilegibles, xlsx vacío, destino inexistente, `list_sheets` inexistente).
- [x] `--self-test` bajo xvfb: OK, código 0.
- [x] pytest + ruff completos pasan.

## Bugs
Ninguno bloqueante.

## Observaciones (no son bugs, para el líder)
- OBS-1 `detect_types` sobre 2M filas tarda 14,7 s, al borde del objetivo de 15 s (la corrida de backend dijo 4,5 s; puede variar con la máquina, o el CSV 2M tiene más columnas que el usado antes). Si pasa de 15 s en otra PC no rompe el test (límite 30 s), pero conviene vigilarlo en Fase 2.
- OBS-2 `to_number(...).failed` no incluye las celdas vacías (según contrato: "NO vacías"). El esperado de `resumen_ventas.json` sí cuenta la vacía de la fila 48; el test las suma aparte. Si la UI debe avisar "celdas vacías en Importe", hay que calcularlo fuera de `failed`.
- OBS-3 Escribir en una ruta que es una carpeta devuelve el `ExcelToolError` genérico ("Revisá que haya espacio en el disco…"), mensaje poco preciso pero seguro.
- OBS-4 `ReadOptions(sheet=...)` con un CSV se ignora (como dice el contrato).

## Para probar a mano (PM)
- Abrir la app (`scripts/dev.sh`) debe seguir funcionando (la Fase 1 no tiene pantallas nuevas).
- Opcional, en una consola Python con `.venv/bin/python`: `from excel_tool.core.io import read_table; from pathlib import Path; print(read_table(Path("ruta/a/tu_archivo.xlsx")).df.head())` con un archivo real tuyo: revisar tildes, ñ y códigos con ceros (00123).
- Dar un archivo real con filas de título arriba y ver que `header_row` detecte bien el encabezado.

## Para probar en Windows (PM)
- `scripts\dev.bat -Tests` debe pasar (incluye los tests de integración nuevos `tests/integracion/test_fase1_io.py`; los de rendimiento se saltan si no están los archivos grandes).
- Abrir un .xlsx en Excel (sin cerrarlo) y leerlo con `read_table`: debe salir el mensaje "Cerrá el archivo en Excel y volvé a intentar." Si en cambio sale "No se pudo abrir el archivo", es el caso anotado en `docs/fases/fase-1.md` (un `OSError` distinto de `PermissionError`): avisar para refinar. Probar lo mismo con un .csv abierto en Excel.
- Guardar un resultado como `csv_excel` con `write_table` (por ejemplo desde `scripts/make_fixtures.py` o la consola) y abrirlo con doble clic en Excel: tiene que verse en columnas separadas, con tildes y ñ bien y "00123" (puede mostrarse como 123 en Excel por ser número; el archivo conserva el 00123).
- `scripts\build_windows.bat` y correr el `.exe` con `--self-test`: debe devolver 0 (ahora usa `core/io`).
