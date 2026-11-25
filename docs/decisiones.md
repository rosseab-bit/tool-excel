# Registro de decisiones

Formato: ADR-NNN · fecha · nivel (2 = líder, 3 = PM) · decisión · motivo · alternativas descartadas · impacto

Las decisiones nuevas se agregan al final. No se borran: si una cambia, se agrega una nueva que la reemplaza y se indica "reemplaza ADR-NNN".

- ADR-001 · 2026-11-25 · 3 · Formatos: CSV + Excel (.xlsx/.xlsm/.xlsb/.xls), entrada y salida, mezclables; lectura de .ods.
- ADR-002 · 2026-11-25 · 3 · Join = 1 principal + N secundarios por columna clave, eligiendo columnas por archivo.
- ADR-003 · 2026-11-25 · 3 · Por secundario: modo detalle o resumen, elegido siempre por el usuario (sin default).
- ADR-004 · 2026-11-25 · 3 · Resumen con agrupar-por y cálculos elegidos por el usuario; también sobre el resultado final. En v1.
- ADR-005 · 2026-11-25 · 3 · Explorador de un archivo en v1, construido antes que el join.
- ADR-006 · 2026-11-25 · 3 · Control de tamaño antes de guardar (Excel / Google Sheets / solo guardar) con recomendación de dividir; límites configurables.
- ADR-007 · 2026-11-25 · 3 · Equipo: líder técnico Opus 5.5 (effort medio, no toca código) + backend, frontend y QA en Sonnet 5.5. El PM decide los cambios de arquitectura.
- ADR-008 · 2026-11-25 · 3 · Plataformas: desarrollo en Linux; la app corre con Python en Linux y Windows; .exe compilado en Windows con PyInstaller. Scripts de entorno (dev) y compilación (build) para cada sistema, en la Fase 0.
- ADR-009 · 2026-11-25 · 3 · Todo local: sin GitHub Actions ni CI. Al repo se sube solo código. Archivos de prueba generados por script (no se suben); solo se suben resultados esperados chicos. Ejecutables fuera del repo. Pruebas en Windows a cargo del PM al cerrar las Fases 0, 2, 5 y 8.
- ADR-010 · 2026-11-25 · 3 · Los agentes no commitean; commitea el líder cuando QA aprueba. Nadie hace push salvo el PM.

## Preguntas abiertas
1. Tamaño típico de los archivos (pandas en memoria vs. streaming/DuckDB).
2. Destino más usado: Excel o Google Sheets (default de la primera vez).
3. ¿Firmar el .exe? (sin firma, Windows SmartScreen muestra advertencia).
4. Distro/versión de Linux de los usuarios finales, si alguno usa Linux.
5. PC o VM con Windows para pruebas y compilación.
6. ¿Entrega temprana solo con el explorador?
