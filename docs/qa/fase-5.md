# QA Fase 5 (F5-20): Asistente para unir archivos

## Resultado: APROBADO (F5-A, F5-B y F5-C resueltos; queda la prueba en Windows del PM)
Re-verificación F5-20b después de los arreglos. Suite verde 3 veces (2 corridas + la del build), memoria estable en ~1,2 GB, build con SELF-TEST OK.

## Estado de los defectos
| Defecto | Estado | Verificación |
|---|---|---|
| F5-C (variables tk sin `master=`) | **Resuelto** | Los 10 tests rojos pasan; suite completa verde |
| F5-A (cantidad de partes sugerida vs. generada) | **Resuelto** | Excel "sin cortar un cliente": sugiere 9, vista previa 9 (`299992 ... 299987, 120064`); ninguna minúscula. Suma = 2.519.998 |
| F5-B (memoria) | **Resuelto** | Pico 1.193 a 1.203 MB en 3 corridas (tope 1.500) |

Se sacó el `xfail` de `tests/integracion/test_fase5_grande.py`: ahora exige (Excel y Sheets) partes en la vista previa = sugeridas o una más, que en Excel la menor sea > 1/10 de la mayor y que cada archivo de Sheets pese < 100 MB. Además mide, en `escenario_grande_ui.py`, el cambio de Excel a Google Sheets con el Guardar ya abierto (sin reiniciar el estado): vuelve a **5 partes** (las sugeridas) y la división de Excel (300.000) ya no queda.

### Observaciones nuevas (no bloquean)
- Excel: la división preelegida ahora es de 300.000 filas (el tope cómodo), no 280.000: da 9 partes con la última de 120.064 filas. Correcto.
- Google Sheets: el semáforo sugiere 504.000 filas por parte (5 partes parejas), pero Guardar preelige 624.999 (el tope por parte): también 5 partes, pero la última de 20.032 filas (4 x ~625.000 + 20.032). Entra y suma bien; es solo menos parejo. Si se quiere parejo, que Guardar use `suggested_rows_per_part` cuando no hay "sin cortar" (decisión del líder).

## Memoria antes / después (`medicion_fase5`, pico = ru_maxrss)
| | Antes (5 corridas) | Después (3 corridas) |
|---|---|---|
| Pico total | 1.443 a 1.674 MB | **1.195 / 1.193 / 1.203 MB** (`/usr/bin/time`: 1.223.532 / 1.221.244 / 1.231.816 KB) |
| Dispersión | ±115 MB | ±5 MB |
| Unido en memoria | 1.045 | 1.008 |
| Guardado Excel + Sheets | 1.121 | 1.105 |
| Resumen final sobre el resultado completo | 1.443 a 1.674 | 1.194 |
| Base (leer y comparar) | 648 | 601 |

## Suite
- `xvfb-run -a scripts/dev.sh --tests`, una por vez: corrida 1 **1336 passed, 2 skipped** (506 s); corrida 2 **1336 passed, 2 skipped** (504 s). Antes: 10 a 11 failed.
- Ruff `check` y `format`: OK.
- `scripts/build_linux.sh` completo (incluye la suite: 1336 passed, 559 s): **SELF-TEST OK**, `dist/ExcelJoinTool-0.1.0-linux.tar.gz` 76,6 MB, SHA256 `7e4f544fdab155ca0c58d0202287b520361dfb708c6553239d2bd27eb03498d1`.
- `git status`: ningún generado sin ignorar (solo código, tests y docs).

## Defectos encontrados en la primera verificación (todos resueltos, se dejan como historia)

### F5-C (ALTA para el cierre: deja la suite en rojo; para el usuario no se nota) - frontend
**Qué falla:** 10 tests de `tests/ui/test_unir_pasos_1_3.py` pasan solos (17 passed) pero fallan en la suite completa
(`test_paso2_validar_y_volcar`: "Marcá al menos una columna clave"; los de paso 3: `assert 1 == 2`, el asistente no avanza).
**Causa:** variables de tkinter creadas **sin `master=`** en `ui/unir/`:
`lista_casillas.py:27` (`tk.StringVar()`) y `:54` (`tk.BooleanVar(...)`), `pasos/paso2_clave.py:65` y `pasos/paso3_sumar.py:99`.
Sin `master` la variable se ata a la primera ventana de Tk que se creó en el proceso. En la suite, `tests/integracion/conftest.py` ya
crea un único Tk de sesión (ADR de Windows: un solo intérprete por sesión) y `tests/ui` crea el suyo, así que las casillas quedan
atadas al otro intérprete y no reflejan el clic. En la app real hay un solo Tk, por eso no se nota; pero es frágil y rompe las pruebas.
**Reproducir:** `xvfb-run -a .venv/bin/python -m pytest tests/integracion/test_fase4_ui.py tests/ui/test_unir_pasos_1_3.py -q`
(10 failed, 11 passed). Solo el segundo archivo: 17 passed.
**Arreglo:** pasar `master=self`/`master=marco` en esas 4 variables (como ya hacen `paso4_modo.py` y `paso6_resumen.py`). No toqué código de producto.

### F5-A (MEDIA) - backend (la regla) y frontend (lo que muestra); lo decide el líder
**Qué falla:** en el escenario grande para Excel, el semáforo dice **9 partes** (`check.suggested_parts = 9`, `suggested_rows_per_part = 280.000`)
pero la división que deja preelegida el paso Guardar ("sin cortar un mismo cliente", columna ID cliente) muestra **10 archivos**, y el
último tiene **37 filas** (`[279993, 279994, 279997, 279997, 279999, 279998, 279992, 279994, 279997, 37]`).
**Causa:** 9 x 280.000 = 2.520.000 apenas supera las 2.519.998 filas: sobran 2 filas. Al no poder cortar un cliente a la mitad, cada parte
termina unas filas antes de 280.000 y el resto se desborda a una parte nueva. No se pierde ni se duplica nada (la suma da 2.519.998 y ninguna
parte pasa de 300.000), pero la cantidad que se sugiere no es la que se genera y queda un archivo casi vacío.
**Pasos:** `.venv/bin/python -m pytest tests/integracion/test_fase5_grande.py -q -s` (1 min 20 s; xfail con el detalle), o a mano: unir
`clientes_grande.csv` + `ventas_grande.csv` en detalle -> paso 7 -> "Unir" -> "Guardar el resultado" -> Excel: la vista previa dice
"Se van a generar 10 archivos".
**Gravedad:** media-baja (no pierde datos). Arreglo posible: `suggested_rows_per_part` con margen cuando hay "sin cortar" (por ejemplo
ceil(filas/partes) + tamaño del grupo más grande), o preelegir "En partes de" sin agrupar.
**Observación relacionada (no es bug, decisión de diseño):** al abrir "Guardar el resultado" por segunda vez en el mismo asistente viene
preelegida la división del guardado anterior (la receta la recuerda). Después de guardar para Excel, abrirlo para Google Sheets muestra
10 partes de 280.000 filas (válidas: 2,24 M de celdas c/u) en vez de las 5 que sugiere el semáforo, porque "en partes de" ya está
"editado". Conviene que el usuario lo sepa o que al cambiar de destino se vuelva a sugerir.

### F5-B (MEDIA) - ver el análisis de memoria más abajo; reparto UI y core
El pico del recorrido completo pasa del tope de 1,5 GB (ADR-023) en algunas corridas (1.443 a 1.674 MB). El camino principal
(unir en detalle y guardar para Excel y Google Sheets) pica en **1.121 MB**, dentro del tope. Lo que lo pasa es encadenar en la misma sesión
"achicar" y un resumen final sobre el resultado completo.

### Observaciones menores (no son bugs)
- Las columnas con el mismo nombre: la primera conserva el nombre ("Fecha") y la que choca lleva el archivo **sin extensión**
  ("Fecha (cobranzas)"); PLAN §3.3 pone de ejemplo `Teléfono (crm.xlsx)`. El checklist ("sufijo con el nombre del archivo") se cumple.
- Los avisos de importes ilegibles ("3 celdas con contenido que no se pudieron leer") aparecen en el editor de resumen y **después de unir**;
  en "Antes de unir" (paso 7) todavía no, porque `diagnose_join` no suma. Es coherente con la estimación sin unir.
- El resumen final sobre un resultado enorme usa como muestra "los primeros 5.000 clientes unidos": en el escenario grande son 63.546 filas;
  el aviso dice la cantidad real de la muestra (no 5.000) y que "al unir se calcula con todo".
- Para Google Sheets el 🔴 del escenario grande viene también por el peso (106,7 MB con 6 columnas > 100 MB), no solo por las celdas.

## Los dos pedidos del líder (primera verificación)

### 1. BUG F5-A: ver arriba (qué falla, causa, pasos, gravedad y arreglo posible).

### 2. Por qué la memoria llegó a 1.632 MB (y a otras 1.443 / 1.497 / 1.546 / 1.674)
Medí el recorrido completo por la UI con `xvfb-run -a /usr/bin/time -v .venv/bin/python -m tests.integracion.medicion_fase5`
(pico = `ru_maxrss`; es un máximo histórico, no baja aunque se libere). **Cinco corridas del mismo recorrido: 1.632 / 1.443 / 1.497 /
1.546 / 1.674 MB** (`/usr/bin/time`: 1.670.844 / 1.477.928 / 1.533.384 / 1.582.924 / 1.714.332 KB). La variación de ±115 MB entre corridas
iguales viene del recolector de basura y del reparto de memoria de Python en cada corrida; no hay nada aleatorio en los datos.

**Pico por etapa (estable entre corridas, hasta la etapa 5):**

| Etapa | Pico (MB) | Qué la sube |
|---|---|---|
| leer clientes y ventas (con su perfil) y compararlos | 648 | las dos tablas leídas como texto |
| unir en detalle (2,5 M filas) | 1.045 | + ~400 MB: la tabla unida |
| guardar para Excel y para Google Sheets | 1.121 | + ~75 MB (partes y vista previa de partes) |
| achicar el resultado a 200.000 filas | 1.126 | nada |
| volver a detalle (paso 4: `diagnose_join`) | 1.247 | + 120 MB |
| tabla de muestra para el resumen final (`tabla_para_resumen`: diagnóstico + unión de 5.000 clientes) | 1.327 | + 80 MB |
| **unir con resumen final sobre el resultado completo** | **1.443 a 1.674** | + 120 a 350 MB, según la corrida |

**Core solo, sin UI** (script con `read_table` + `diagnose_join` + `run_join` con resumen final + `run_join` en detalle, todo vivo a la vez):
**952 a 993 MB** de pico; leer las dos tablas sin perfil: 376 MB. O sea que el core, por sí solo, hace todo el recorrido con ~1 GB.

**Qué parte es de la UI y cuál del core (mi lectura; no medí cada pieza por separado salvo lo indicado):**
- **UI, base permanente (~+270 MB sobre el core):** la UI deja cargadas las dos tablas y además un perfil por archivo (`leer_fuente` llama a
  `profile_table`). Medí el perfil aparte: **+235 MB de pico** (376 -> 610 MB) y 648 MB ya en el paso 3 (vs 376 en el core solo). El perfil
  se pide por la UI pero lo calcula `core/profile.py` (conteos de valores distintos sobre 2,5 M de filas): una parte es de cada lado.
- **UI, el resultado guardado en el estado:** `paso7.resultado` (el `JoinResult` de 2,5 M filas, ~400 MB) y `paso7.ventana_guardar`
  (`VentanaGuardar.df`, el mismo DataFrame, más sus partes) quedan vivos mientras el usuario sigue en el asistente. **Probé soltarlos**
  (experimento `LIBERAR=1` en `escenario_grande_ui.py`: `resultado`, `revision`, `ventana_guardar` y las ventanas, con `gc.collect()` antes
  del resumen final): el pico final **no bajó** (1.497 y 1.674 MB en dos corridas). Por eso no es la causa principal, aunque sí conviene
  soltarlos (son ~400 MB vivos sin necesidad entre pasos).
- **UI, recálculos repetidos:** cada cambio de paso vuelve a llamar a `diagnose_join` sobre las tablas completas (paso 3, 4, 5, 7 y de nuevo
  en `tabla_para_resumen`), cada uno con temporales de 100 a 300 MB. Es lo que explica las etapas "+120" y "+80", y deja la memoria
  fragmentada para la unión final.
- **UI, tabla para el resumen final:** la muestra en sí es chica (63.546 filas); lo que cuesta es el diagnóstico completo previo (+80 MB). No es
  el problema.
- **Core:** la unión con resumen final (`run_join` + `summarize` sobre 2,5 M filas) necesita ~+350 a 580 MB por sí sola (952 MB de pico en el
  script del core, con 376 de base). Es el **mismo trabajo** que la etapa de "unir en detalle", más una copia para agrupar.

**Conclusión para asignar:** ningún componente solo pasa el tope; lo pasa la suma de la base de la UI (~650 MB con perfiles) más el
pico transitorio de unir con resumen final (~+400 a 600 MB del core) con la memoria ya fragmentada por los diagnósticos previos.
- **frontend** (primero, es lo más barato): soltar `resultado`/`ventana_guardar`/`revision` al cambiar de paso o al volver atrás; no
  recalcular `diagnose_join` completo si el plan no cambió (cachear por plan); perfilar con muestra o a pedido las tablas enormes.
- **backend** (segundo): bajar el pico de `profile_table` en tablas de millones de filas (+235 MB) y el de `run_join` con resumen final
  (resumir por bloques sin materializar el detalle completo cuando el resumen final solo agrupa/cuenta).
Si el PM prefiere no tocar nada: el camino principal (detalle + guardar) pica en 1,1 GB; el riesgo está en usar "achicar" y resumen final
encadenados sobre el grande en una PC de 8 GB. Hay que medirlo también en Windows (paso 57 de `fase-2.md`).

## Tests nuevos y cambios (solo en mis carpetas)
- `tests/integracion/asistente_util.py` (no es un test): el `Conductor` maneja el asistente como un usuario, desde "Unir archivos" en la pantalla
  de inicio; solo reemplaza el selector de archivos, los diálogos del sistema y `open_folder`. En cada "Siguiente" verifica que la ventana y su pie
  entren en 768 px.
- `tests/integracion/test_fase5_asistente.py` (4): las 4 variantes del escenario por la UI, guardando con el paso Guardar y releyendo con openpyxl.
- `tests/integracion/test_fase5_checklist.py` (16): ítems de PLAN §6 (ver tabla abajo).
- `tests/integracion/test_fase5_reglas.py` (12): reglas sobre `ui/unir/`.
- `tests/integracion/test_fase5_grande.py` (1, xfail estricto de F5-A), `escenario_grande_ui.py` y `medicion_fase5.py` (la misma corrida como script, con el pico).
- `tests/integracion/conftest.py`: fixtures `esc` y `c` (escenario y conductor) compartidos.
- **Defecto de mis propios tests de la Fase 4, arreglado:** `_textos` de `test_fase4_ui.py` juntaba las letras con espacios ("j o i n"), así que el
  chequeo de jerga de esa pantalla **nunca podía fallar**. Lo corregí (ahora mira palabras de verdad) y sigue pasando: no había jerga.
- **Intermitente mío, arreglado:** `test_fase4_join.py::test_receta_guardada_y_cargada_...[xlsx]` comparaba byte a byte los `.xlsx` (incluido
  `_indice.xlsx`), que llevan la fecha y hora de creación: fallaba si los dos guardados caían en segundos distintos. Ahora compara las celdas.
- Los fixtures chicos extra (2 columnas clave, 8 archivos) los arma cada test con `tmp_path` y valores calculados a mano (determinísticos); no
  agregué nada a `make_fixtures.py` ni a `esperados/`.

## Criterios de aceptación de la fase (`docs/fases/fase-5.md`)
- [x] **4 variantes desde la UI, sin tocar código.** Detalle: 88 filas, filas por cliente = esperado, importes originales por recuento con `csv`.
  Resumen por cliente: 30 filas en el orden del principal, totales/compras/artículos/última compra = esperado para los 27 con compras, y los
  clientes 5, 17 y 28 con el **número 0** (ADR-037) y fecha vacía. Cliente + artículo: filas = pares + 3 sin compras, sumas por par con `Decimal`.
  Resumen final por Región: 5 filas = `resumen_final_por_region.json` (totales, clientes y total + excluido = suma de todo el CSV). Todas guardadas
  en xlsx y releídas; "Cancelar" pasa a "Terminar" al guardar.
- [x] **Escenario grande en detalle (UI).** Paso 4: 🔴 "~2.519.998 filas: no entra en una sola hoja de Excel" y "Al guardar te vamos a proponer
  dividirlo". Paso 5: 2.519.998 x 8 columnas baja a x 6 al desmarcar Nombre y Vendedor (20,2 M -> 15,1 M de celdas; Google Sheets deja de pasarse por
  celdas pero sigue 🔴 por peso). Paso 7: 🔴 Excel + "Achicar el resultado". Guardar para Excel: 🔴, "Todo en un solo archivo" deshabilitado, partes de
  <= 280.000 filas (CSV para Excel), suma 2.519.998 contra el total y recuento con `csv` de cada archivo. Google Sheets (partida de cero): 🔴, 5 partes
  sugeridas (5 M de celdas c/u) y cada archivo < 100 MB. "Abrir carpeta" llama a `open_folder` con la carpeta de las partes. **Con la salvedad de F5-A** (10
  archivos en vez de 9).
- [x] **Achicar el resultado:** el botón "Cambiar ventas_grande.csv a resumen por ID cliente -> 200.000 filas 🟢" se aplica con un clic y el semáforo pasa a 🟢.
- [x] **Resumen final con > 200.000 filas:** usa una muestra (63.546 filas del resultado de unir los primeros 5.000 clientes), avisa "El resultado es muy grande
  (2.519.998 filas)... al unir se calcula con todo" y el resultado completo (5 filas) da, por Región, `detalle_left.por_region` de `grandes.json`.
- [x] **Checklist §6 "Unir archivos" y "Resultados grandes":** ver tabla.
- [x] **Receta guardada y abierta con archivos del mes siguiente.** Se guarda desde el paso 7 (con resúmenes, orden de columnas y CSV en partes de 12),
  sin rutas absolutas ni `\`; copiada a otra carpeta con ventas "del mes siguiente" (importes x2) abre en "Archivos a sumar", mismas columnas y orden, el
  resultado da el doble y "Guardar" viene preelegido en CSV, 12 filas, 3 archivos. Con la receta sola: "Faltan archivos" y se completan con "Elegir…".
- [ ] 🪟 Prueba en Windows: la hace el PM (sección H nueva de `docs/qa/fase-2.md`, pasos 51 a 58).

## Checklist de PLAN §6 (lo automatizado y cómo)
| Ítem | Test |
|---|---|
| clientes + ventas en detalle, filas del resultado antes de guardar | `test_variante_detalle...` (88 filas en el paso 4 y el 7) |
| dos detalles: aviso y "Cambiar a resumen" en un clic | `test_dos_detalles_avisan...` (210 filas -> 88; el aviso desaparece) |
| ventas en resumen por cliente, Sumar Importe | `test_variante_resumen_por_cliente...` |
| cliente + artículo | `test_variante_cliente_y_articulo` |
| agrupar por mes | `test_agrupar_por_mes_de_fecha` ("Fecha (mes)", 2024-01...) |
| varios cálculos con nombres editados | `test_varios_calculos_con_nombres_editados...` |
| Listar valores sin repetidos | `test_listar_valores_sin_repetidos` |
| suma con inválidos: el aviso | `test_el_editor_avisa...` y el del resultado |
| resumen final por Región | `test_variante_resumen_final_por_region` |
| detalle y resumen en dos hojas | `test_detalle_y_resumen_en_dos_hojas...` (hojas Detalle, Resumen y "Sin coincidencia en ventas" con 98 y 99) |
| 3 archivos en resumen = filas de clientes | `test_receta...` y `test_ocho_archivos...` |
| clave con otro nombre ("Nro Cliente") | `test_dos_detalles...`, `test_ceros...` |
| clave emparejada sola (mayúsculas y espacios) | `test_clave_con_otro_nombre_y_emparejada_sola...` |
| elegir y reordenar columnas, clave una sola vez | `test_elegir_y_reordenar_columnas...` |
| columnas con el mismo nombre | `test_columnas_con_el_mismo_nombre...` |
| ceros a la izquierda (5 contra 27) | `test_ceros_a_la_izquierda...` |
| 2 columnas clave | `test_dos_columnas_clave_y_columna_cantidad...` |
| 5 o más archivos | `test_ocho_archivos_secundarios_con_solo_en_ambos` (8 archivos, uno "solo los que están en los dos": 20 filas) |
| receta con archivos del mes siguiente | `test_receta_guardada_abierta...` y `..._que_faltan_se_completan...` |
| hoja "Sin coincidencia" | dos hojas (arriba) |
| Escenario grande: aviso en paso 4 y al guardar | `test_fase5_grande.py` |
| indicador del paso 5 baja | `test_fase5_grande.py` |
| Excel 🔴 obliga y sugiere partes | `test_fase5_grande.py` (con F5-A) |
| Google Sheets 🔴 (no ofrece "varias hojas") | `test_fase5_grande.py` |
| Cambiar a resumen desde Guardar -> 🟢 | `test_fase5_grande.py` ("achicar") |
| Excel 🟡 a 400.000, "solo guardarlo", dividir por Región, `_indice.xlsx`, abrir en Excel y Sheets | ya cubiertos en Fase 3 (`docs/qa/fase-3.md`); abrir en Excel/Sheets es a mano |

## Columna "Cantidad" (F5-04)
`test_columna_cantidad_de_ventas_en_guardar_y_en_valores_distintos`: el detalle completo con "Cantidad" de ventas. En **Valores distintos** del resultado
("Ver información del resultado"): encabezados "Cantidad | **Cantidad de filas** | Porcentaje", 21 valores distintos en 88 filas, cuentas iguales al recuento
con `csv`, y combinada con Región suma 88. En **Guardar**: "Buscar columnas" sugiere "Cantidad: 21 archivos", se divide por esa columna y la suma de las
partes da 88 con las cuentas por valor del recuento independiente. Sin carteles de error. Funciona.

## Reglas (`ui/unir/`)
- Sin jerga, ahora con "join", "merge", "left", "inner", "key", "plan": cadenas con espacio en `ui/unir/` y en toda `ui/`; las de una palabra solo
  pueden ser `left`/`inner` (valores internos). Más el chequeo de los textos **visibles** en cada ventana del asistente (corre en todos los tests de
  integración).
- La UI no calcula: ni `groupby`, `merge`, `.sum(`, `.mean(`, `.agg(`, `value_counts`, `concat`, etc. en todo `ui/` (texto sin docstrings ni comentarios).
- Nada de widgets desde hilos: por AST, las 10+ tareas que se mandan a `run_in_background` (y los métodos propios que llaman, p. ej. `_hojas`)
  no llaman a `configure`, `set`, `insert`, `after`, `messagebox`, etc.; el detector se prueba con un caso malo. Además `ui/unir` no lee ni escribe tablas ni
  usa `sys.platform`, `subprocess` ni `print`.
- Todas las ventanas nuevas entran en 768 px: el asistente en cada paso (con 8 archivos y con el aviso de dos detalles), el editor de resumen del paso 4 y
  el del paso 6 (con "Listo" a la vista), el explorador del resultado y la ventana Guardar abierta desde el asistente.

## Para probar a mano (PM)
Sección H de `docs/qa/fase-2.md` (pasos 51 a 58): escenario de 3 archivos con el `.exe`, "Abrir carpeta", archivo abierto en Excel al guardar,
emojis 🔑 🧾 📊 1️⃣ ✅ ⚠️, asistente en 1366x768 con 125 %, receta de Linux en Windows, escenario grande (con el pico de memoria) y el `.exe` sin consola.

## Veredicto
**Aprobada** desde QA en Linux. Pendiente solo la prueba en Windows del PM (sección H de `docs/qa/fase-2.md`, pasos 51 a 58), incluida la medición de memoria del escenario grande.
