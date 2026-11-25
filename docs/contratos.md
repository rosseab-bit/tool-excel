# Contratos core ↔ UI

Dataclasses y firmas públicas que cruzan de `src/excel_tool/core/` a `src/excel_tool/ui/`.

- **Lo mantiene el líder técnico.** Ningún agente lo edita: si hace falta un cambio, termina su tarea con estado BLOQUEADO y lo pide.
- Cambiar un contrato existente es nivel 2 si no cambia lo que ve el usuario, y nivel 3 si lo cambia.
- Punto de partida: los modelos de `PLAN.md` §2. El líder los completa y los cierra al planificar cada fase, antes de despachar al backend y al frontend.

## Estado

| Contrato | Fase | Estado |
|---|---|---|
| `app_entry.main` (`--version`, `--self-test`) | 0 | Por definir |
| `platform_utils` (carpetas de usuario, abrir carpeta, nombres seguros) | 0 | Por definir |
| `ReadOptions`, `CsvOptions`, `read_table`, `write_table` | 1 | Por definir |
| `ColumnType`, funciones de `core/types.py` | 1 | Por definir |
| `TableProfile`, `ColumnProfile`, `SheetSummary`, `value_counts`, `rows_with_value` | 2 | Borrador en PLAN.md §2 |
| Eventos de progreso y cancelación (`worker`) | 2 | Por definir |
| Errores del core hacia la UI (excepciones y mensajes) | 2 | Por definir |
| `SizeEstimate`, `SizeCheck`, `SplitSpec`, `OutputSpec`, `save_result` | 3 | Borrador en PLAN.md §2 |
| `AggregationSpec`, `JoinPlan`, `SourceSpec`, `SecondarySpec`, `KeyOptions`, `diagnose_join`, `run_join` | 4 | Borrador en PLAN.md §2 |

## Contratos

_(Vacío. El líder los agrega acá a medida que se cierran.)_
