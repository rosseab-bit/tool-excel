# QA Fase 2 (F2-20): Explorador

## Resultado: APROBADO CON OBSERVACIONES
## Suite: pytest 564 passed / 0 failed / 2 skipped (4 corridas completas, todas iguales) · ruff OK · git status limpio de generados: SÍ

Corridas de `xvfb-run -a scripts/dev.sh --tests`: 222,8 s, 221,3 s, 220,5 s (pytest) y una cuarta dentro de
`scripts/build_linux.sh` (220,3 s) y una quinta con `-rs` (221,0 s). Ninguna falla intermitente. No apareció el
aviso `main thread is not in main loop` de F2-11 en ninguna de las corridas.

Los 2 skips: `tests/integracion/test_fixtures.py:185` (cubre el caso "faltan los grandes" y los grandes existen) y
`tests/integracion/test_fase2_explorador.py:282` (no hay fixture `.ods`, ver OBS-2).
Los fixtures grandes estaban generados; no se regeneró nada. `scripts/build_linux.sh`: SELF-TEST OK (el self-test
exige pyarrow real; `dist/ExcelJoinTool/_internal/pyarrow` tiene 28 `.so`), carpeta 206 MB, `.tar.gz` 73 MB;
el ejecutable corrido a mano con `--self-test` bajo xvfb también da OK, código 0.

## Tests nuevos (tests/integracion/)
- `test_fase2_explorador.py` (33 tests + 1 skip): core de punta a punta con "para explorar" en CSV y en xlsx
  (`read_table` -> `profile_table` -> `value_counts` -> `rows_with_value` -> `filter_rows` / `sort_rows` ->
  `get_page` -> `export_profile` / `export_distincts` -> releer). Los esperados salen de `perfil.json` y de
  recontar el CSV con el módulo `csv` de la biblioteca estándar (sin usar el producto). Varias hojas (3 hojas
  perfiladas una por una), `chico.xls` y `chico.xlsb`.
- `test_fase2_ui.py` (28 tests): UI con xvfb y el core real (sin perfiles falsos), `messagebox` y `filedialog`
  reemplazados por monkeypatch en `tests/integracion/conftest.py` (los carteles se registran y se verifican).
- `test_fase2_reglas.py` (10 tests): los grep de reglas de CLAUDE.md, como chequeo automático sobre el AST de `src/`.
- `medicion_fase2.py` (no es un test): medición de tiempo y, con `/usr/bin/time -v`, pico de memoria.
- `conftest.py` de integración: entorno de pantalla (mismo cuidado con el recolector de basura que `tests/ui/conftest.py`).

## Criterios de aceptación de la fase
- [x] **Core contra `perfil.json`** (CSV y xlsx): 200 filas, 13 columnas, 4 filas vacías, 8 repetidas; por columna:
  posición, letra, tipo, vacíos, distintos, ceros a la izquierda, espacios al borde y desglose de la mixta
  (`entero` 158 / `texto` 38). Variantes "Norte" / "NORTE" / "Norte " (`case_variants` = 3, un grupo en
  `variant_groups`, y la opción de juntar las unifica). Ninguna columna es "posible clave" (ID tiene 4 vacíos).
  La columna sin nombre sale como "Columna J" y la segunda "Total" como "Total (2)" (aviso incluido).
  Tests: `test_perfil_*`, `test_value_counts_*`.
- [x] **Distintos / filas con un valor / filtros / orden / páginas** contra un recuento independiente (módulo `csv`):
  Región (7 valores, porcentajes suman 100), dos columnas a la vez, buscar, incluir vacíos, orden alfabético;
  `rows_with_value` (exacto y juntando variantes); filtro "contiene" ("ÑANDÚ" encuentra "ñandú": 32 filas),
  "mayor" (109 filas), ambos a la vez (Y), vacío / no vacío, valor inválido -> `ExcelToolError` con mensaje;
  orden numérico descendente con vacíos al final; `get_page` con "N.º fila" correcto y páginas fuera de rango vacías.
- [x] **Informe e informe de distintos, releídos:** `export_profile` -> xlsx con hojas General, Columnas, Región y
  Avisos (valores iguales a los del perfil); `export_distincts` en xlsx y CSV (`;`), cantidades suman 200, el
  "Norte " con espacio se conserva.
- [x] **Varias hojas:** `varias_hojas.xlsx` (Clientes 12x4, Ventas 25x6, Resumen año 2x2): cada hoja perfilada con
  `sheets=list_sheets(...)`; en la UI, cambiar de hoja re-analiza, vacía los distintos y actualiza Datos.
- [x] **Otros formatos:** `chico.xls` y `chico.xlsb` dan el mismo perfil que el xlsx y se exploran en la UI
  (12 filas x 4 columnas). **`.ods`: no hay fixture**, no verificado (OBS-2).
- [x] **Desde la UI (xvfb, core real), CSV y xlsx:** se abre el explorador con "para explorar", se recorren las 4
  pestañas (General, Columnas, Valores distintos, Datos) con los datos correctos; calcular distintos de Región
  (solo a pedido: antes de calcular no hay resultado); **doble clic real** (eventos de ratón) sobre "Norte" lleva
  a Datos con 31 filas y el aviso "Región es «Norte»"; "Quitar filtro" vuelve a 200; filtro "contiene" (32 filas),
  "es mayor que" (109 filas), ambos combinados; orden por encabezado dos veces (de mayor a menor: filas 122, 197,
  160 y valores descendentes); juntar variantes (5 filas); exportar informe y lista (CSV) a `tmp_path` y releerlos
  (cartel "Listo"); valor de filtro inválido y carpeta de destino inexistente -> cartel amigable, sin traceback;
  archivo dañado (`danado.xlsx`) -> cartel amigable y la ventana se cierra; sin jerga en ningún texto visible.
- [x] **Cartel de memoria:** con `platform_utils.available_memory` reemplazado por 5 MB, el selector muestra
  "Archivo muy grande" ("necesita alrededor de ... y hay ... libres. Cerrá otros programas y probá de nuevo."),
  no carga el archivo, no llama a `read_table` y no acepta la selección. Con 8 GB libres, carga sin cartel.
- [x] **Cancelar el análisis del CSV de 2M desde la UI** (botón Cancelar de la ventana de progreso, que aparece
  sola): la ventana se cierra en 0,02-0,15 s, tanto leyendo (a 1 s) como perfilando (a 8 s). Sin carteles.
  (El hilo de fondo tarda 2,3 s en terminar al cancelar durante la lectura, 0,4 s durante el perfil: OBS-1.)
- [x] **Cerrar la app con un análisis largo en curso** (riesgo de F2-16; procesos aparte, CSV de 2M, a 1,5 s
  leyendo y a 10 s perfilando; cerrando el explorador y cerrando la ventana principal): 4 combinaciones, todas
  con código de salida 0 (sin core dump), sin nada en stderr (sin traceback), terminan en 2,0-12 s
  (el tiempo es el del análisis más el cierre; nunca se cuelga). Pasaron en las 5 corridas de la suite.
- [x] **Rendimiento** (esta máquina: 8 núcleos, 24 GB; pico = "Maximum resident set size" de `/usr/bin/time -v`):

  | Escenario | Tiempo | Pico de memoria |
  |---|---|---|
  | CSV 2M, leer + perfilar | 32,2 s (leer 3,3 + perfil 28,9); objetivo <= 60 s | **1.150 MB** |
  | ... + distintos de `ID` (2M valores únicos; tabla truncada a 100.000, `truncated=True`) | 33,4 s (distintos 1,0 s) | **1.150 MB** |
  | ... + filtro `Importe > 500000` (1.000.404 filas) y orden numérico | 44,6 s (filtro 4,5 s; orden 2,7 s; orden de las 2M filas 4,8 s; página 0,1 s) | **1.150 MB** |
  | Camino real de la UI (xvfb: explorador + distintos de ID + filtro + orden + exportar informe) | 46,7 s en total (perfil listo a los 32,8 s) | **1.161 MB** |
  | xlsx 500k, leer + perfilar | 16,9 s (leer 8,7 + perfil 8,2) | **765 MB** |

  Presupuesto de ADR-023 (<= 1,5 GB): se cumple con 340 MB de margen en todos los escenarios, y distintos, filtro y
  orden no suben el pico (lo marca el perfil). Reproducir:
  `/usr/bin/time -v .venv/bin/python tests/integracion/medicion_fase2.py perfil|distintos|filtro_orden|xlsx`
  y, bajo xvfb, `... medicion_fase2.py ui`.
- [x] **La grilla nunca carga más de una página:** `Grilla` con `page_size` 1000 (contrato) en el explorador;
  con páginas de 50 sobre las 200 filas muestra 50 por página (4 páginas, primera/siguiente/última, la última
  empieza en la fila 151); con los 2M, tras filtrar y ordenar, la grilla muestra 1000 filas. Los distintos se
  calculan solo a pedido (sin resultado hasta apretar "Calcular" / "Ver valores distintos").
- [x] **Textos sin jerga:** grep manual en `src/excel_tool/ui/` de join, merge, group by, distinct, split, null,
  DataFrame, encoding, dtype, traceback y afines solo en textos visibles: ninguno. Chequeo automático
  `test_textos_de_la_ui_sin_jerga` (AST, sin docstrings ni logs; probado que detecta "join" y "DataFrame") y
  asserts sobre los textos reales renderizados en las 4 pestañas. Únicas menciones: nombres de formato de
  archivo (".xlsx", ".csv") y la palabra "bytes" para tamaños menores a 1 KB (OBS-3).
- [x] **Reglas de CLAUDE.md** (grep + `test_fase2_reglas.py`, 10 tests):
  - sin tkinter en `core/` (solo la mención en un docstring de `core/progress.py`);
  - sin `read_csv` / `read_excel` / `to_csv` / `to_excel` / `ExcelWriter` fuera de `core/io`;
  - sin `sys.platform` / `os.name` / `platform.system` fuera de `platform_utils.py`;
  - sin 1048576, 20_000_000 ni 16384 en `src/`;
  - todo `open(` / `.open(` / `read_text` / `write_text` con `encoding=` (solo `path.open("rb")` binario en `core/io/csv.py`);
  - nada a stdout/stderr salvo `app_entry._print` (solo `--version`, `--self-test` y uso, contrato de la Fase 0);
  - `subprocess` solo en `platform_utils.py` (`xdg-open`, ADR-012);
  - sin `/tmp` fijo ni `_MEIPASS`.
- [x] **Mensajes de error amigables, sin tracebacks en pantalla:** valor de filtro inválido, destino inexistente y
  archivo dañado muestran un cartel de `showerror` con mensaje en castellano (probado).
- [x] **Suites:** 3 corridas de `xvfb-run -a scripts/dev.sh --tests` (564 passed, 2 skipped, ruff OK, "Todo bien");
  `scripts/build_linux.sh` con SELF-TEST OK y pyarrow real.
- [x] **git status:** ningún archivo generado sin ignorar. Quedan sin commitear solo los fuentes y tests de la
  fase (más mis 5 archivos nuevos en `tests/integracion/`); `tests/fixtures/generados/`, `.venv/`, `build/` y
  `dist/` están ignorados.
- [ ] **Prueba en Windows (ADR-009)** -> pendiente del PM, checklist más abajo.

## Bugs
Ninguno bloqueante ni de gravedad media o alta.

## Observaciones (no son bugs; para el líder)
- OBS-1 [back] baja. Al cancelar durante la lectura del CSV de 2M, el hilo de fondo tarda 2,3 s en terminar
  (3 corridas: 2,28, 2,31, 2,35 s); el contrato de cancelación dice <= 2 s. La UI no lo sufre (la ventana se
  cierra en < 0,2 s y el resultado se descarta, ADR-027). Durante el perfil el hilo termina en 0,4 s. Lo mide
  `test_cancelar_el_analisis_del_csv_de_2m_...[leyendo]`, cuyo tope es 5 s.
- OBS-2 [qa] baja. No hay fixture `.ods` (el generador no lo crea y el plan de la fase lo pedía "si hay
  fixtures"). Hay un skip fijo, `test_ods_sin_fixture`. Si el PM quiere cobertura de `.ods`, hay que pedir a
  backend que `make_fixtures.py` lo genere (calamine lo lee; escribirlo necesita otra herramienta, nivel 2).
- OBS-3 [front] baja. `ui/formato.py:61` muestra "N bytes" para archivos de menos de 1 KB; "bytes" es jerga
  mínima. Se podría decir "N caracteres" o mostrar siempre KB.
- OBS-4 (ya anotada en F2-03/F2-14). Filtrar y ordenar no se pueden cancelar (la ventana de progreso aparece
  si tardan más de 0,3 s, sin botón). En esta máquina: 3-5 s con 2M filas; en una PC más lenta podrían ser ~10 s.
- OBS-5 (ya anotada en F2-03). El orden alfabético pone "Ñ" y las vocales con tilde después de la Z. Un test mío
  (`test_orden_alfabetico_sin_importar_mayusculas`) lo deja asentado comparando con el orden de Python.
- OBS-6. La advertencia de F2-11 (`main thread is not in main loop`) no apareció en 5 corridas completas.
- OBS-7. Margen de `test_rendimiento_2m_filas` (detect_types 2M, aviso > 15 s, falla > 30 s): en la medición de
  QA el test completo dura 17,4 s (incluye la lectura de ~3 s). Pasó en las 5 corridas.
- OBS-8. `perfilar` factoriza dos veces cada columna (ya anotada en F2-04): el perfil de 2M tarda ~29 s, con la
  mitad del margen respecto de los 60 s del criterio. En una PC con la mitad de velocidad quedaría al límite.

## Para probar a mano (PM, en Linux)
- `scripts/dev.sh`, "Ver información de un archivo", elegir `tests/fixtures/generados/explorar/para_explorar.csv`:
  General (200 filas, 13 columnas, 4 filas vacías, 8 repetidas), Columnas (clic en "Región": "Norte" / "NORTE" /
  "Norte " como pocos valores), "Ver valores distintos de esta columna", doble clic en "Norte", Datos, probar
  filtrar y ordenar, "Guardar informe del archivo...".
- Lo mismo con `.../explorar/para_explorar.xlsx` y con `.../formatos/varias_hojas.xlsx` (cambiar de hoja).
- Con `tests/fixtures/generados/grandes/rendimiento_2M.csv`: tiene que mostrar el progreso, poder cancelarse y
  cerrarse en un instante; cerrar la ventana a mitad del análisis.

## Para probar en Windows (PM): checklist de las Fases 0 a 3 (la Fase 3 es la sección E; el detalle de la Fase 3 está en `docs/qa/fase-3.md`)
Lo que hay que hacer en una PC nueva con Windows (ideal: 8 GB de RAM, como las de los usuarios). Anotar al lado
de cada paso OK / FALLA / observación (con captura si falla). Referencias: `README.md`, `docs/fases/fase-0.md`,
`fase-1.md` y `fase-2.md` (Notas).

**A. Instalación y preparación (Fase 0)**
1. Instalar **Git for Windows** (https://git-scm.com).
2. Instalar **Python 3.12 de 64 bits** desde https://www.python.org/downloads/ marcando **"tcl/tk and IDLE"** y
   **"Add python.exe to PATH"**. Verificar en una consola: `py -3.12 -c "import struct, tkinter; print(struct.calcsize('P')*8)"`
   debe imprimir `64` (sin error de tkinter). Importante: sin Python de 64 bits no hay ruedas de pyarrow.
3. Clonar el repo y entrar: `git clone <URL del repo> tool-excel` y `cd tool-excel`.
4. `scripts\dev.bat -Help` -> muestra la ayuda.
5. `scripts\dev.bat -Setup` -> crea `.venv` e instala dependencias (incluida pyarrow). Repetirlo: la segunda vez
   dice "Dependencias al dia". Riesgo: `constraints.txt` salió de Linux; si falta algún paquete exacto en
   Windows, anotar cuál. Verificar que **no use el alias falso de Python de la Microsoft Store**.
6. `.venv\Scripts\python.exe scripts\make_fixtures.py` -> genera `tests\fixtures\generados\` (incluye el CSV de 2M
   y el xlsx de 500k; tarda varios minutos; con `--sin-grandes` los saltea, pero entonces no se pueden hacer los
   pasos 17 a 20).
7. `scripts\dev.bat -Tests` -> tiene que terminar con **"Todo bien"** (en Linux, tras la Fase 3: 891 passed, 2 skipped). Si falla
   algo, copiar el nombre del test y el mensaje. Los tests de pantalla (`pantalla`) corren con la ventana real.
   Anotar el tiempo total (en Linux: ~3 min 40 s hasta la Fase 2; ~6 min con la Fase 3).
8. `scripts\dev.bat -AppArgs --version` -> imprime "Excel Join Tool 0.1.0"; `scripts\dev.bat -AppArgs --self-test`
   -> `SELF-TEST OK` (ahora incluye pyarrow real y `core/io`).
9. Doble clic en `scripts\dev.bat` -> abre la ventana de la app (la consola espera y no se cierra sola ante un error).

**B. Explorador a mano (Fase 2)**
10. En la app: "Ver información de un archivo" abre el selector; elegir `tests\fixtures\generados\explorar\para_explorar.csv`.
    En el selector: separador "Punto y coma", tipo de texto "Recomendado (con acentos y ñ)", vista previa con
    tildes y ñ bien (ej.: "Ñandú", "Descripción"). Cerrar la app y volver a abrirla: el selector recuerda la
    última carpeta.
11. Aceptar. **General:** 200 filas, 13 columnas, 4 filas vacías, 8 repetidas, avisos sobre "Columna J", "Total (2)",
    "Mixta" y "Región". **Columnas:** 13 filas con tipos; clic en cada encabezado ordena; al elegir "Región" se ve el
    detalle y los valores más frecuentes. **Valores distintos:** elegir Región, "Calcular" -> 7 valores y aviso de
    "Norte / NORTE / Norte"; tildar "Juntar escrituras parecidas" y calcular de nuevo -> 5 valores; probar
    "Buscar un texto" y ordenar de la A a la Z; elegir dos columnas (Activo y Región).
12. **Doble clic** en "Norte" -> salta a Datos con 31 filas y el cartel "Estás viendo solo las filas donde Región
    es «Norte»"; "Quitar filtro" -> 200 filas. En Datos: "Agregar condición" (Descripción contiene "ñandú" ->
    32 filas; Importe es mayor que 5000 -> 109 filas), clic en el encabezado "Importe" para ordenar.
13. **Menú del clic derecho** sobre el encabezado de la grilla de Datos: ocultar y volver a mostrar columnas,
    "Mostrar todas las columnas". (Mirar a ojo: en Windows el menú de Tk puede quedarse abierto o no cerrarse con
    clic afuera; anotar.)
14. **Orden alfabético con ñ y tildes:** en Datos ordenar por "Descripción" (hay "Ñandú ..."): hoy "Ñ" y las
    vocales con tilde quedan después de la Z (conocido, F2-03); anotar si al PM le molesta, se decide en la Fase 9.
15. "Guardar informe del archivo..." -> guardar como `.xlsx` y abrirlo con **Excel real**: hojas General, Columnas
    y una por cada lista de distintos calculada; tildes y números bien. En Valores distintos, "Exportar esta lista..."
    como `.xlsx` y como `.csv`.
16. **CSV con doble clic (Fase 1):** el `.csv` exportado en el paso 15 se abre con doble clic en Excel con las
    columnas separadas, con tildes y ñ bien y la coma decimal (el "00123" del código puede verse como 123 en
    Excel, pero el archivo lo conserva). También abrir un **CSV exportado por Excel de Windows** (Guardar como ->
    "CSV (delimitado por comas)", que en un Windows en castellano usa `;` y cp1252): el selector tiene que
    detectar "Windows antiguo" y mostrar bien las tildes.
17. **Archivos de varias hojas y otros formatos:** `tests\fixtures\generados\formatos\varias_hojas.xlsx` (cambiar de
    hoja con la lista de arriba a la derecha: se vuelve a analizar), `chico.xls` y `chico.xlsb`.
18. **`chico.xlsb` en Excel real (pendiente de la Fase 0):** abrir `tests\fixtures\generados\formatos\chico.xlsb`
    con Excel: tiene que abrir sin pedir "reparar" y mostrar la hoja "Clientes" (12 filas, 4 columnas). Se generó con
    un escritor propio (ADR-017). Hacer lo mismo con `chico.xls`.
19. **Archivo abierto en Excel (pendiente de la Fase 1):** abrir un `.xlsx` en Excel y dejarlo abierto; en la app,
    "Ver información" sobre ese mismo archivo: tiene que decir **"Cerrá el archivo en Excel y volvé a intentar."**.
    Si dice "No se pudo abrir el archivo", es el caso de `OSError` distinto de `PermissionError` anotado en
    `docs/fases/fase-1.md`: avisar. Repetir con un `.csv` abierto en Excel.

**C. Rendimiento con 2M de filas (Fase 2)** (con el Administrador de tareas abierto, pestaña Procesos)
20. Abrir `tests\fixtures\generados\grandes\rendimiento_2M.csv` (178 MB, 2M de filas): hay ventana de progreso,
    sin congelarse; **tiempo hasta ver todo el perfil <= 60 s** (en Linux 33 s, 8 núcleos); **memoria de la app
    <= 1,5 GB** (en Linux 1,16 GB). Anotar ambos números. Probar distintos de "ID" (2M únicos; avisa "se calcularon
    solo los primeros 100.000"), filtro "Importe es mayor que 500000" y ordenar por Importe (3 a 5 s en Linux):
    la grilla muestra 1.000 filas por página y se mueve con "Siguiente".
21. Repetir el paso 20 con `rendimiento_500k.xlsx` (en Linux 17 s, 765 MB).
22. **Cancelar:** abrir otra vez el CSV de 2M y apretar "Cancelar" a los 2 s (leyendo) y a los 15 s (perfilando):
    la ventana de progreso y el explorador se cierran **en menos de 2 s** y el consumo de CPU y memoria del proceso
    baja en unos segundos.
23. **Cerrar la app durante un análisis (riesgo de F2-16):** abrir el CSV de 2M y, mientras analiza, (a) cerrar el
    explorador con la X y (b) en otra prueba, cerrar la ventana principal. Esperado: no se congela, no aparece el
    cartel de Windows "dejó de funcionar" ni un traceback, y el proceso desaparece del Administrador de tareas en
    unos segundos; el log (`%LOCALAPPDATA%\ExcelJoinTool\logs`) no tiene errores. Hacerlo también con el `.exe`.
24. **Cartel de memoria (opcional, mejor en una VM con 4 GB):** con poca memoria libre, abrir el CSV de 2M: tiene que
    salir "Archivo muy grande ... necesita alrededor de X GB y hay Y GB libres. Cerrá otros programas y probá de
    nuevo." (montos con coma) y **no** cargar el archivo. Si el cartel no sale con la PC cargada, revisar
    `platform_utils.available_memory` (usa `GlobalMemoryStatusEx`).

**D. Compilación del .exe (Fases 0 a 3)**
25. `scripts\build_windows.bat` (corre tests, compila, prueba el `.exe` y arma `dist\ExcelJoinTool-0.1.0-windows.zip`).
    Tiene que terminar con `SELF-TEST OK` y mostrar tamaño y SHA256. En Linux la carpeta pesa 206 MB y el
    `.tar.gz` 73 MB; anotar los de Windows. Si pregunta SmartScreen ("Windows protegió su PC": Más información ->
    Ejecutar de todas formas) el self-test del build queda esperando hasta aceptarlo.
26. **pyarrow en el .exe (F2-08):** `dir dist\ExcelJoinTool\_internal\pyarrow\*.dll` debe listar las DLL de pyarrow
    (`arrow.dll`, `arrow_python.dll`, etc.). Si el self-test falla por una DLL, revisar el filtro de binarios
    en `excel_tool.spec` (no sacar nada que enlace `arrow_python.dll`) y avisar. Opcional: repetir con
    `scripts\build_windows.bat -OneFile`.
27. Copiar el `.zip` a una **PC sin Python**, descomprimir y abrir `ExcelJoinTool.exe` con doble clic: se abre sin
    consola negra; repetir los pasos 10 a 12 y 20 con el `.exe`.

**E. Guardar, dividir y límites (Fase 3)**
Archivos: `make_fixtures.py` (paso 6) ya genera `grandes\fase3_400k.csv`, `grandes\fase3_dividir_1600k.csv` y
`fase3\nombres_raros.csv`. Lo ya cubierto por los pasos 15, 16 y 19 (Excel real, CSV con doble clic, archivo
abierto en Excel al LEER) no se repite: acá se prueba al GUARDAR. Referencias: `docs/fases/fase-3.md` (Notas).
28. **Semáforo en General:** abrir `fase3_400k.csv`: Excel en amarillo ("CONVIENE DIVIDIR") y Google Sheets en
    verde, sin botón "Dividir y guardar…". Abrir `fase3_dividir_1600k.csv`: Excel en rojo ("NO ENTRA EN UN SOLO
    ARCHIVO") con el botón "Dividir y guardar…" y Google Sheets en amarillo. **Colores de Windows:** el texto se
    lee bien sobre cada color (verde, amarillo, rojo), también con el modo oscuro y con "contraste alto"; anotar.
29. **Exportar lo filtrado, con Excel real:** `para_explorar.csv`, pestaña Datos, "Importe es mayor que 5000",
    "Exportar lo filtrado…", Excel, guardar en el Escritorio, abrir en Excel: 109 filas; **encabezado congelado**
    (al bajar queda la fila 1), **flechitas de filtro** en el encabezado; ID, Importe y Total son **números**
    (alineados a la derecha, la barra de estado suma); Fecha es **fecha real** (se puede filtrar por año/mes);
    "Código" conserva el cero ("02055") como texto; Excel no pide "reparar" ni muestra #NOMBRE?.
30. **Textos que empiezan con "=":** crear en el Bloc de notas `ID;Texto` / `1;=1+1` / `2;+5` / `3;-3` / `4;@x`
    (guardar como CSV UTF-8), abrirlo en la app, exportarlo como .xlsx y abrirlo en Excel real: las celdas se
    ven literalmente como `=1+1`, `+5`, etc. (texto, sin calcular ni avisar de fórmula).
31. **Dividir y guardar, por valor:** `fase3_dividir_1600k.csv` -> "Dividir y guardar…" -> "Un archivo por cada valor"
    de Región -> "CSV para Excel" -> Guardar. Sale "Se guardaron N archivos con 1.600.000 filas en total
    (coincide con el resultado ✅)" y **"Abrir carpeta" abre el Explorador de Windows en esa carpeta**. En la carpeta:
    las partes y `_indice.xlsx` (abrirlo en Excel: columnas Archivo, Hoja, Filas, Desde fila, Hasta fila y Valor;
    Filas como números). Repetir con `fase3_400k.csv` en **.xlsx**, por filas, y abrir **cada parte en Excel real**
    como en el paso 29 (encabezado congelado, filtro, números y fechas reales).
32. **Nombres raros en el Explorador de Windows:** `nombres_raros.csv`, un archivo por cada valor de Sucursal (en
    CSV y en .xlsx). Tienen que existir los 20 archivos sin error, con nombres como `nombres_raros_Norte_Sur.csv`,
    `nombres_raros_Norte_Sur (2).csv` (el de "Norte/Sur" y el de "Norte_Sur" no se pisan), `nombres_raros_CON.csv`,
    `nombres_raros_nul.csv`, `nombres_raros_a_b.csv`, `nombres_raros_aux.txt.csv`, `nombres_raros_vacio.csv`, y
    los dos valores larguísimos de "AAAA…" acortados; ninguno con espacio o punto al final. Abrir varios con doble
    clic. Ninguno puede llamarse solo `CON` o `NUL` (Windows los trata como dispositivos).
33. **Carpeta repetida:** guardar dos veces por partes con el mismo nombre en la misma carpeta: la segunda crea
    `<nombre> (2)` y la primera no cambia.
34. **Archivo abierto en Excel al GUARDAR:** guardar un resultado en un solo archivo sobre un .xlsx que está abierto
    en Excel (misma ruta, confirmando que se reemplaza): cartel amable ("Cerrá el archivo en Excel…"), el original
    sigue intacto, **no queda ningún `.tmp` en la carpeta**. Cerrar Excel y repetir: guarda bien. Anotar si el mensaje
    dice otra cosa (caso `OSError` distinto de `PermissionError`, ver `fase-1.md`).
35. **Ruta larga:** guardar por partes en una carpeta anidada de ~200 caracteres: las partes se acortan y aparece el
    aviso "Algunos nombres de archivo se acortaron…"; con una ruta de más de ~250 caracteres, mensaje amable de
    "ruta demasiado larga" (sin traceback). Probar con "rutas largas" habilitadas y deshabilitadas en Windows.
36. **Antivirus y `.tmp`:** guardar `rendimiento_500k.xlsx` convertido o el resultado de 400k en .xlsx con el
    antivirus activo (Defender u otro de la empresa): al terminar no queda ningún `.tmp` ni `.part`; si el antivirus
    retiene el temporal y falla el renombrado, tiene que salir un cartel amable (nunca un traceback) y no dejar
    archivos a medias. Anotar si el escaneo hace notablemente más lento el guardado.
37. **Cancelar al guardar:** `fase3_dividir_1600k.csv` por partes de 200.000 filas en .xlsx, apretar "Cancelar" a
    los 20 s: se cierra la ventana de progreso, no queda la carpeta de partes (ni vacía) ni archivos `.tmp`, y lo que
    ya existía en la carpeta no cambió. Lo mismo con "CSV para Excel".
38. **CSV de salida:** una parte "CSV para Excel" abre con doble clic en Excel de Windows en castellano con las
    columnas separadas, tildes y ñ bien y coma decimal. Una parte "CSV para Google Sheets" (`,` y punto decimal) en
    Excel de Windows en castellano se puede ver mal (esperado: anotar cómo); importarla en **Google Sheets real**
    (Archivo -> Importar): debe entrar completa, y los avisos "La parte N pesa X MB" aparecen si pasa de 100 MB.
39. **`limites.json`:** copiar junto al `.exe` un `limites.json` con `{"excel": {"comfortable_rows": 100000}}`,
    reiniciar y abrir `fase3_400k.csv`: el semáforo de Excel pasa a "Conviene dividir" con otro número de partes;
    con `{"excel": {"max_rows": 300000}}` pasa a rojo; con el archivo roto (texto que no es JSON) la app abre igual
    y usa los valores por defecto. Repetir con el de `%LOCALAPPDATA%\ExcelJoinTool\limites.json` (si están los
    dos, manda este último). Mirar que el log registre una línea cuando el archivo está roto.
40. **Ventana Guardar en pantalla chica:** en una notebook de 1366x768 con escala del 100 %, 125 % y 150 %
    (Configuración -> Pantalla): la ventana entra entera, los botones "Guardar…" y "Cerrar" y el cartel con
    "Abrir carpeta" quedan a la vista sin tapar la barra de tareas; el cuerpo se desplaza con la barra y con la
    rueda del mouse (también con el puntero sobre la tabla de partes). Anotar si algo queda cortado.
41. **Memoria y tiempo con 2M (presupuesto 1,5 GB):** `rendimiento_2M.csv` -> "Dividir y guardar…" -> por filas
    (7 partes) en "CSV para Excel" y en .xlsx: pico de memoria en el Administrador de tareas **<= 1,5 GB** (en Linux
    1,15 GB, igual que el perfil: guardar no lo sube) y tiempos (Linux: guardar CSV 29 s, .xlsx 173 s).
42. Repetir 28, 29, 31 y 37 con el `.exe` compilado (el paso 25 ya prueba que `xlsxwriter` entra en el ejecutable).

**F. Para informar**
43. Pasar al líder: resultado de cada paso (OK / FALLA / observación), los números de tiempo y memoria de los
    pasos 7, 20, 21, 25 y 41, y el log de `%LOCALAPPDATA%\ExcelJoinTool\logs` si algo falló.
