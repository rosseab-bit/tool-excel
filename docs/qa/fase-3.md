# QA Fase 3 (F3-20b): Salida: límites, división y guardado

## Resultado: APROBADO CON OBSERVACIONES
## Suite: pytest 891 passed / 0 failed / 2 skipped (3 corridas de `xvfb-run -a scripts/dev.sh --tests` + 1 dentro de `scripts/build_linux.sh`, todas iguales) · ruff OK · git status limpio de generados: SÍ

Corridas: 354,7 s, 354,6 s, 354,0 s y 355,9 s (la del build). Ninguna falla intermitente del producto.
Los 2 skips son los mismos de la Fase 2 (`test_fixtures.py:185`, que cubre el caso "faltan los grandes", y
`test_fase2_explorador.py:282`, sin fixture `.ods`).
Un incidente mío, ya corregido: una primera corrida de `build_linux.sh` SIN xvfb falló en uno de mis tests nuevos
(`test_limites_json_cambia_el_semaforo_...`): medía la posición del botón en coordenadas de pantalla, y con una
pantalla real el administrador de ventanas ubica la ventana donde quiere (y=854). Ahora se mide contra el borde de la
ventana. Pasa con y sin xvfb. No es un bug del producto.
`scripts/build_linux.sh`: **SELF-TEST OK** (incluye xlsxwriter y pyarrow reales), carpeta 206 MB, `.tar.gz` 73 MB,
SHA256 `caeac783...c4447733`.

## Tests nuevos (tests/integracion/)
- `test_fase3_output.py` (74 tests): `save_result` y `split_file` de punta a punta con el core real. Se guarda,
  se relee TODO con `read_table` y se compara con el original. Cubre cada modo (un archivo, partes por filas, sin cortar
  grupos, por valor, por valor con subdivisión, varias hojas) en xlsx y en los tres estilos de CSV, con y sin
  `column_types`. Verifica contenido igual, suma de filas, encabezado en cada parte, `_indice.xlsx`, nombres válidos en
  Windows, y con openpyxl: encabezado congelado (`A2`), filtro, números y fechas reales, texto con ceros y textos con
  "=" / "+" / "-" / "@" como texto. Además: carpeta existente -> " (2)" / " (3)" sin tocar lo anterior ni archivos
  ajenos, cancelar a mitad (partes, un archivo, hojas, `split_file`) sin dejar nada ni tocar lo que existía, 🔴 sin
  división -> error amigable, límites de `limites.json` en acción, tabla vacía, ruta larga (se acortan los nombres), y
  los archivos grandes (`fase3_dividir_1600k.csv` y `fase3_400k.csv`).
- `test_fase3_ui.py` (9 tests): explorador + paso Guardar con xvfb y el core real (sin dobles de `save_result`, perfil
  ni límites). Solo se reemplazan los diálogos del sistema y `platform_utils.open_folder`.
- `test_fase3_reglas.py` (17 tests): reglas de CLAUDE.md sobre el AST de `src/` (ver más abajo).
- `medicion_fase3.py` (no es un test): tiempo y pico de memoria con `/usr/bin/time -v`.

## Criterios de aceptación de la fase (`docs/fases/fase-3.md`)
- [x] **Escenario grande:** `fase3_dividir_1600k.csv` (1.600.000 x 4): `check` da 🔴 para Excel; partes sugeridas = max(ceil(1,6M / 300.000) = 6,
  las que pidan los MB), con `suggested_rows_per_part` <= 300.000 (`test_escenario_grande_es_rojo_para_excel`). "Solo
  guardar" en CSV es 🟢; en xlsx rige el límite duro (ceil(1,6M / 1.048.575) = 2 partes). **400k:** 🟡 Excel, 🟢 "Solo
  guardar" (CSV y xlsx), 2 partes sugeridas. En la UI real el explorador lo muestra igual: 400k = Excel amarillo
  ("CONVIENE DIVIDIR", `#f9a825`) y Google Sheets verde, sin botón de dividir; 1,6M = Excel rojo (`#c62828`) con el
  botón "Dividir y guardar…" y Google Sheets amarillo (`test_semaforo_de_400k_...`, `test_semaforo_de_1600k_...`).
- [x] **Ninguna parte supera el límite; `rows_keep_groups` nunca corta un grupo; `by_value` subdivide la región grande;
  suma = total; el índice lista bien las partes.** Con `save_result` sobre 200 filas, en xlsx y en los 3 CSV:
  `test_partes_por_filas` (70/70/60, nombres `salida_parte_01_de_03`), `test_sin_cortar_grupos` (ningún valor de Región
  en dos partes), `test_por_valor` (7 partes = 7 valores, filas por valor contra un `Counter`), `test_por_valor_con_subdivision`
  (partes <= 20 y ceil por valor), `_verificar_indice` (Archivo, Filas, Desde, Hasta, Valor iguales a `res.parts`).
  Con 1,6M (`split_file`, CSV): `Norte` (1.200.000) se subdivide en 2 y las otras 4 regiones van enteras = 6 partes,
  todas <= 1.048.575; las filas por región coinciden con `fase3_grandes.json` (recuento independiente con el módulo
  `csv`); cada archivo tiene exactamente las filas que declara y el encabezado; índice con suma 1.600.000.
  Por filas de 400.000 = 4 partes y la primera y última fila de cada parte son las del original.
  Límite duro: con `limites.json` en 101 / 51 filas ninguna parte lo pasa.
- [x] **Cambiar un límite en `limites.json` cambia la recomendación:** en el core (`test_limite_de_limites_json_...`) y
  en la UI real: con `{"excel": {"max_rows": 101, "comfortable_rows": 50}}`, 200 filas pasan a 🔴 en General (botón
  "Dividir y guardar…"), la ventana Guardar muestra "NO ENTRA EN UN SOLO ARCHIVO", deshabilita "Todo en un solo
  archivo" y el resultado guardado tiene partes <= 100 filas. Un `limites.json` roto vuelve a los valores por defecto.
  **Ningún límite fuera de `core/limits.py`:** `test_los_limites_estan_solo_en_core_limits` (1.048.575/6, 20.000.000,
  16.384, 5.000.000, 300.000), que además prueba que el regex detecta lo que debe.
- [x] **Valores con caracteres prohibidos en Windows:** `nombres_raros.csv` (20 valores: `Norte/Sur`, `Norte_Sur`, `CON`,
  `nul`, `a:b`, `Zona*1`, `aux.txt`, `COM1`, `Zona "x"? <2>`, `Dir\Sub|1`, `Punto.`, valores de 200 letras, vacío, `Sur`
  y `sur`): en CSV y xlsx salen 20 archivos con nombres válidos (sin `<>:"/\|?*`, sin espacio o punto al final, ni
  `CON`/`NUL`/`COMn` puros), sin repetir (`nombres_raros_Norte_Sur (2).csv`, `..._sur (2).csv`), la ruta < 260 y
  los archivos del disco = las partes + el índice. También como hojas de un xlsx (<= 31, sin `[]:*?/\`).
- [x] **Memoria (ADR-023, <= 1,5 GB), CSV de 2M** (esta máquina: 8 núcleos, 24 GB; pico = "Maximum resident set size" de
  `/usr/bin/time -v`; 7 partes de 300.000 filas; reproducir con
  `/usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase3.py <escenario>`):

  | Escenario | Tiempo | Pico |
  |---|---|---|
  | `split_file` por filas (CSV -> 7 partes CSV, sin cargar el archivo) | 7,7 s | **218 MB** |
  | `split_file` por Región con subdivisión a 1M (5 partes) | 7,8 s | **255 MB** |
  | `save_result` CSV, 7 partes, tras leer y perfilar (como la UI) | leer 3,3 + perfil 28,7 + guardar 29,2 = 61,2 s | **1.148 MB** |
  | `save_result` CSV por Región con subdivisión (5 partes), tras perfilar | 62,0 s (guardar 29,9) | **1.148 MB** |
  | `save_result` xlsx, 7 partes, tras leer y perfilar, con tipos | leer 3,3 + perfil 28,7 + guardar 172,6 = 204,7 s | **1.147 MB** |

  Cumple con 350 MB de margen. El pico es el del perfil (1,15 GB, igual que la Fase 2): guardar no lo sube, ni en
  CSV ni en xlsx. Sin perfil (solo `split_file`) el pico es 5 veces menor.
- [x] **Desde el explorador se exporta lo filtrado, con semáforo y división (UI real):** `para_explorar.csv`, pestaña
  Datos, "Importe es mayor que 5000" (109 filas, el mismo número que la Fase 2), el botón dice "Exportar lo filtrado…" (sin
  filtro dice "Guardar estas filas…"), Excel / xlsx por defecto, nombre sugerido `para_explorar_filtrado`, guardar en
  `tmp_path`. Relectura: las 109 filas EXACTAS (IDs y Regiones iguales a los del CSV filtrado con el módulo `csv`);
  con openpyxl: ID como número, Importe como decimal real (3322,83), Fecha como fecha real, "Código" como texto
  con el cero ("02055"), encabezado congelado, filtro, 110 filas con el encabezado. Mensaje: "Se guardó 1 archivo con 109
  filas en total (coincide con el resultado ✅)".
- [x] **"Dividir y guardar…" con el archivo de 1,6M, por valor de Región:** abre el paso Guardar, 🔴, "Todo en un solo
  archivo (no entra)" deshabilitado, las columnas sugeridas se buscan solas, se elige Región y "CSV para Excel", la vista
  previa lista las partes (una fila por parte, "Se van a generar N archivos."), se guarda: carpeta `ventas/` con las
  partes y `_indice.xlsx`, suma 1.600.000, filas por región = recuento independiente, todas <= 1.048.575, y el cartel
  "Se guardaron N archivos con 1.600.000 filas en total (coincide con el resultado ✅)". "Abrir carpeta" llama a
  `platform_utils.open_folder(<carpeta de las partes>)` (con monkeypatch). Sin ningún cartel de error.
- [x] **Recordar el destino:** la primera vez Excel + xlsx (ADR-032); al elegir Google Sheets queda en las preferencias
  (`last_destination`), el formato pasa a "CSV para Google Sheets", "En varias hojas" desaparece, el archivo sale con `,`
  y utf-8 sin BOM; al volver a abrir Exportar el destino es Google Sheets (y la carpeta del último guardado).
- [x] **La ventana Guardar entra en 768 px:** altura pedida <= 768 y botones "Guardar…", "Cerrar" y "Abrir carpeta" dentro de
  los 768 en las combinaciones destino x forma de dividir (3 destinos x 5 formas) y con el resultado a la vista;
  el cuerpo tiene tope de 470 px y se desplaza con barra y rueda.
- [x] **Cancelar a mitad no deja nada ni toca lo que existía** (`save_result`: partes por filas / por valor / sin cortar
  grupos en xlsx y CSV, un archivo xlsx y CSV, varias hojas; `split_file` por filas y por valor con el de 1,6M): se
  compara byte a byte una foto de la carpeta antes y después (archivo ajeno, carpeta existente con un archivo, el
  destino previo con "ORIGINAL"): idéntica. No queda carpeta nueva, ni `.tmp`, ni `.part`. Cancelar antes de empezar también.
- [x] **Carpeta existente -> " (2)":** `salida`, `salida (2)`, `salida (3)` (xlsx y CSV); una carpeta ajena con
  `mis_notas.txt` no se toca; también en `split_file`.
- [x] **🔴 sin división -> error amigable:** "El resultado no entra en un solo archivo. ..." (sin traceback, no escribe
  nada), con 500 filas y 101 de tope; con 1,6M reales en xlsx (también "Solo guardar" + xlsx); con Google Sheets y
  poco límite de celdas. "Solo guardar" CSV de 1,6M sí escribe (1.600.000 filas, contadas en el disco).
  Varias hojas en CSV o con Google Sheets -> error amigable ("hojas" / "Google Sheets"), sin escribir nada.
- [x] **Reglas de CLAUDE.md** (`test_fase3_reglas.py`, 17 tests, más los 10 de la Fase 2):
  - límites solo en `core/limits.py` (el self-test de `app_entry.py` lee el JSON por defecto, permitido);
  - ni `read_csv` / `read_excel` / `to_csv` / `to_excel` / `ExcelWriter` fuera de `core/io`; `csv`, `xlsxwriter` y
    calamine solo se importan ahí; `csv_bytes` se define solo en `core/io` (el uso en `limits.py` es correcto);
  - openpyxl no se importa en `src/` y está solo en el extra `dev` de `pyproject.toml` (xlsxwriter en las
    dependencias);
  - nada de tkinter ni `excel_tool.ui` en `core/`;
  - `sys.frozen` / `_MEIPASS` solo en `platform_utils.py`; los resultados se guardan solo con `save_result`;
  - textos de la UI sin jerga, ahora también "split", "chunk", "spec" y "output" (el detector está probado).
  Más: los textos que se ven en la ventana Guardar renderizada, sin jerga.
- [x] **Suites:** 3 corridas de `xvfb-run -a scripts/dev.sh --tests` (891 passed, 2 skipped, ruff y format OK, "Todo bien") +
  `scripts/build_linux.sh` con SELF-TEST OK (206 MB / 73 MB).
- [x] **git status:** ningún archivo generado sin ignorar (`tests/fixtures/generados/`, `.venv/`, `build/`, `dist/` están ignorados). Quedan sin
  commitear los fuentes y tests de la fase, mis 4 archivos nuevos en `tests/integracion/` y el cambio en
  `docs/qa/fase-2.md` (checklist de Windows).
- [ ] **Prueba en Windows (ADR-009)** -> pendiente del PM (checklist en `docs/qa/fase-2.md`, sección E, pasos 28 a 42).

## Bugs
Ninguno de gravedad alta, media ni baja que rompa un criterio.

## Observaciones (no son bugs; para el líder)
- OBS-1 [back] baja. Guardar con tipos es bastante más lento que sin tipos: el CSV de 2M en 7 partes con `column_types` tarda 29 s
  (F3-03 midió 4,4 s sin tipos) y el xlsx 173 s (F3-06 midió 136 s sin tipos). Es la conversión de texto a número y fecha, y
  es lo que pide el diseño (la UI siempre pasa tipos). Con 1M de filas (ADR-026) son ~15 y ~85 s. La ventana de progreso
  se muestra; solo para tenerlo presente.
- OBS-2 [back] baja. Los CSV con tipos dejan las fechas con hora como `dd/mm/aaaa hh:mm:ss` (el `13:05` original sale `13:05:00`),
  igual que lo que devuelve la lectura de xlsx. Es consistente con el contrato de lectura, pero el CSV ya no
  es idéntico letra por letra al original en esa columna (un CSV sin tipos sí). Mis tests normalizan los segundos.
- OBS-3 [qa/contrato] baja. Una parte de un xlsx compuesta solo por filas vacías (por ejemplo, el valor "(vacío)" de una columna
  cuando esas filas no tienen nada más) se vuelve a leer con 0 filas, y las filas vacías al final de una hoja se descartan
  al leer (regla de lectura de la Fase 1). `save_result` cuenta bien las filas escritas (la suma coincide); es solo la relectura.
  Mis tests lo contemplan.
- OBS-4 [back] baja. `SplitSpec("sheets")` ignora `value_columns`: si se pide por valor en hojas, divide por filas (y con
  `rows_per_part=None` usa el tope). El contrato solo define `rows_per_part` para `sheets` y la UI no ofrece esa combinación;
  podría lanzar un error amigable en vez de ignorarlo en silencio.
- OBS-5 [back] baja. `by_value` distingue mayúsculas ("Sur" y "sur" son dos archivos; el segundo lleva " (2)"), ya anotado en F3-20a.
  El índice muestra el vacío como "(vacío)" y el archivo como `..._vacio.csv`.
- OBS-6. `split_file` solo trabaja con CSV (Excel por partes queda para la Fase 7, ya anotado). No se probó
  cancelar desde el botón Cancelar de la ventana de progreso al guardar (la cancelación se probó en el core con el
  token real y en `tests/ui/test_guardar.py` con dobles); la prueba manual de Windows (paso 37) lo cubre.
- OBS-7 [qa]. Los tests grandes de la fase suman ~100 s de la suite (el de 1,6M en UI, 400k en xlsx por partes, 1,6M en
  `split_file`). La suite pasó de ~3 min 40 s a ~5 min 55 s.

## Para probar a mano (PM, en Linux)
- `scripts/dev.sh`, "Ver información de un archivo", `tests/fixtures/generados/grandes/fase3_dividir_1600k.csv`: en
  General, Excel en rojo con "Dividir y guardar…" y Google Sheets en amarillo. Dividir -> "Un archivo por cada valor" de
  Región -> "CSV para Excel" -> Guardar; "Abrir carpeta".
- Lo mismo con `fase3_400k.csv` (Excel amarillo, sin botón), pestaña Datos con un filtro y "Exportar lo filtrado…" en
  .xlsx: abrirlo con LibreOffice (encabezado fijo, filtro, números y fechas reales).
- Probar con una ventana de 768 px de alto (por ejemplo `xrandr` o una VM) que la ventana Guardar se desplaza y que
  "Guardar…" queda a la vista.
- Copiar `limites.json` a `~/.local/share/ExcelJoinTool/` con `{"excel": {"comfortable_rows": 100000}}` y mirar cómo cambia el semáforo.

## Para probar en Windows (PM)
La checklist completa está en `docs/qa/fase-2.md` (sección **E. Guardar, dividir y límites (Fase 3)**, pasos 28 a 42;
la sección de informe pasó a ser la F, paso 43). Lo más importante, en orden:
1. Abrir las partes xlsx en **Excel real**: encabezado congelado, filtro, números y fechas reales, "Código" con ceros y textos con "=" como texto (pasos 29-31).
2. **Archivo abierto en Excel al guardar** y que no quede ningún `.tmp` (paso 34); **antivirus** reteniendo el `.tmp` (paso 36).
3. **Ruta larga** (paso 35) y **nombres raros** en el Explorador de Windows: `CON`, `nul`, `Norte/Sur`, `a:b` (paso 32).
4. **CSV para Excel con doble clic** en un Windows en castellano y CSV para Google Sheets importado en Google Sheets real (paso 38).
5. `limites.json` junto al `.exe` y en `%LOCALAPPDATA%\ExcelJoinTool` (paso 39).
6. Colores del semáforo (paso 28) y **desplazamiento de la ventana Guardar** en 1366x768 con escala 125 % y 150 % (paso 40).
7. Memoria <= 1,5 GB y tiempos con el CSV de 2M (paso 41); repetirlo todo con el `.exe` (paso 42).
