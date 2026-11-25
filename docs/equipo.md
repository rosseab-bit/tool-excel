# Equipo y flujo de trabajo con Claude Code

Define quién hace qué, cómo se comunican los agentes y cuándo decide el PM. La configuración está en `.claude/`.

---

## 1. Roles

| Rol | Quién | Modelo | Toca código | Responsabilidad |
|---|---|---|---|---|
| **PM** | Vos | — | No | Prioridades, alcance, aprobar el plan de cada fase y el cierre. **Decisión final en todo lo que cambie la arquitectura.** Pruebas en Windows. `git push` |
| **Líder técnico / Arquitecto** | Sesión principal de Claude Code | **Opus 5.5, effort medio** | **No** (solo `docs/`, `PLAN.md`, `CLAUDE.md`) | Dividir cada fase en tareas, definir contratos entre back y front, despachar agentes, responder sus dudas técnicas, escalar al PM, hacer los commits cuando QA aprueba |
| **Backend** | Subagente | Sonnet 5.5 | `src/excel_tool/core/` y archivos base del paquete, `tests/core/`, scripts de entorno y compilación, `pyproject.toml`, `constraints.txt`, `excel_tool.spec`, `packaging/`, `README.md`, `.gitignore`, `.gitattributes` | Toda la lógica (io, tipos, perfil, join, resúmenes, límites, split, guardado), arranque de la app, scripts y empaquetado |
| **Frontend** | Subagente | Sonnet 5.5 | `src/excel_tool/ui/`, `src/excel_tool/assets/`, `tests/ui/` | Ventanas, asistentes, widgets, worker, textos de UI |
| **QA** | Subagente | Sonnet 5.5 | `tests/integracion/`, `tests/fixtures/`, `tests/conftest.py`, `scripts/make_fixtures.py`, `docs/qa/` | Fixtures, tests de integración, correr todo, revisar contra los criterios de aceptación, reportar bugs. **No arregla código de producto**: reporta |

**Principio:** el líder piensa y coordina, los agentes ejecutan. Opus solo se usa para lo que vale la pena: dividir el trabajo, contratos, preguntas difíciles y revisar los informes de QA. No lee archivos grandes de código: pide resúmenes a los agentes.

**Todo es local:** no hay GitHub Actions ni CI. Los agentes corren tests y scripts en esta máquina (Linux). Las pruebas en Windows las hace el PM y le pasa los resultados al líder.

---

## 2. Cómo se resuelven las preguntas (escalamiento)

Los subagentes **no pueden conversar con el líder a mitad de una tarea**. Cuando necesitan una decisión, **terminan con estado `BLOQUEADO`** y dejan la pregunta armada. El líder responde y **reanuda al mismo agente** (`SendMessage`), que conserva todo lo que ya había hecho.

| Nivel | Tipo de pregunta | Quién decide | Ejemplos |
|---|---|---|---|
| **1** | Detalle de implementación dentro de su carpeta, sin efecto en otros | **El agente** (lo anota en su informe) | Nombre de una función privada, cómo estructurar un test, qué widget de ttk usar |
| **2** | Técnica, con efecto en otro agente o en el contrato, pero **dentro de la arquitectura del plan** | **El líder (Opus)** | Agregar un campo a un dataclass, formato de un mensaje de progreso, cómo reportar errores del core a la UI, si un caso borde es bug o comportamiento esperado |
| **3** | **Cambia la arquitectura, el alcance o lo que ve el usuario** | **El PM** (el líder prepara la pregunta) | Cambiar pandas por DuckDB, agregar una dependencia nueva, cambiar la regla "1 principal + N secundarios", cambiar el flujo de un asistente, mover una función a otra fase, sacar algo del alcance, subir algo que no sea código al repo, cualquier cosa que contradiga el `PLAN.md` |

**Regla para nivel 3:** el líder **no improvisa**. Hace la pregunta al PM con `AskUserQuestion`:
- Contexto en 2 líneas (qué pasó y por qué hay que decidir).
- 2 a 4 opciones, con la **recomendada primero** y marcada "(Recomendado)".
- Impacto de cada opción: días, archivos o fases afectadas, riesgo.

Mientras tanto, se pausa solo lo que depende de esa decisión. El resto sigue.

Toda decisión de nivel 2 o 3 se registra en `docs/decisiones.md`.

---

## 3. Flujo de una fase

```
PM: "Arrancá la Fase N"
 │
 ▼
1. LÍDER planifica
   - Lee PLAN.md (solo la fase N y las secciones que cita) y docs/decisiones.md
   - Divide en tareas chicas (TaskCreate) con: agente, archivos, criterio de aceptación
   - Define o actualiza los contratos (dataclasses/firmas) en docs/contratos.md
   - Le muestra al PM el plan de la fase → espera OK  ◄── punto de control 1
 │
 ▼
2. QA prepara fixtures y tests de aceptación de la fase (si aplica)
 │
 ▼
3. BACKEND implementa el core (tests unitarios primero)
   FRONTEND implementa la UI contra el contrato (en paralelo si el contrato está cerrado;
   puede usar datos falsos hasta que el core esté listo)
   ── si alguno queda BLOQUEADO → el líder responde (nivel 2) o escala (nivel 3)
 │
 ▼
4. QA verifica (en local)
   - Corre pytest + ruff completos
   - Tests de integración de la fase
   - Revisa contra los criterios de aceptación y el checklist del PLAN.md
   - Revisa que no haya archivos generados sin ignorar (git status)
   - Informe: APROBADO / RECHAZADO con bugs numerados (archivo, pasos, esperado vs obtenido)
 │
 ├── RECHAZADO → el líder reparte los bugs a back/front → vuelve a 4
 ▼
5. LÍDER cierra
   - Commit(s) de la fase (git add solo de lo tocado, nunca archivos generados) con mensaje en castellano
   - Resumen al PM: qué se hizo, decisiones tomadas, desvíos, qué probar a mano
     (y, si la fase lo pide, qué probar en Windows)  ◄── punto de control 2
 │
 ▼
PM: prueba a mano (y en Windows si corresponde) → aprueba → git push → siguiente fase
```

**Reglas de coordinación:**
- **Paralelo solo si las carpetas no se pisan** (back en `core/`, front en `ui/`). Si dos agentes necesitan el mismo archivo, van en secuencia.
- **Los agentes no hacen commits.** Commitea el líder al final, cuando QA aprueba. Así se evitan choques de git entre agentes que trabajan a la vez.
- **Nadie hace push**: el push lo hace el PM.
- **El contrato manda.** Si el front necesita algo que el core no da, no lo inventa en la UI: queda `BLOQUEADO` y el líder decide.
- **Tareas chicas:** cada despacho a un agente debe poder terminarse en una sola corrida (idealmente menos de 1 hora de trabajo y un puñado de archivos).
- **Errores de Windows:** el PM le pega el error al líder, el líder lo clasifica y lo despacha como bug. El agente lo arregla sin poder probarlo en Windows: debe explicar en su informe por qué el cambio lo resuelve, y el PM lo vuelve a probar.

---

## 4. Formato de los mensajes

### 4.1 Del líder al agente (despacho)

```
Tarea: F4-03 — summarize() en core/aggregate.py
Contexto: PLAN.md §2 (AggregationSpec) y §3.2 paso 4. Decisión ADR-004.
Contrato: docs/contratos.md#aggregationspec (no cambiar firmas sin consultar).
Archivos que podés tocar: src/excel_tool/core/aggregate.py, tests/core/test_aggregate.py
Criterio de aceptación:
  - Todas las operaciones de AggFunc
  - Importes "1.234,56" se suman bien; celdas no numéricas se reportan
  - Coincide con tests/fixtures/esperados/resumen_cliente.json
Fuera de alcance: join.py, UI.
```

### 4.2 Del agente al líder (informe final, obligatorio)

```
## Estado: TERMINADO | BLOQUEADO | PARCIAL
## Hecho
- ...
## Archivos tocados
- ...
## Verificación
- `.venv/bin/python -m pytest tests/core -q` → 42 passed
- `.venv/bin/ruff check .` → OK
## Decisiones de nivel 1 que tomé
- ...
## Riesgos para Windows (si aplica)
- ...
## Preguntas (solo si BLOQUEADO)
1. Pregunta concreta
   - Opción A: ... (impacto)
   - Opción B: ... (impacto)
   - Recomiendo: A, porque ...
   - Nivel que creo que es: 2 | 3
## Desvíos del plan / riesgos detectados
- ...
```

### 4.3 Informe de QA

```
## Resultado: APROBADO | RECHAZADO
## Suite: pytest X passed / Y failed · ruff OK/FAIL · git status limpio de generados: SÍ/NO
## Criterios de aceptación de la fase
- [x] ...
- [ ] ... → BUG-12
## Bugs
- BUG-12 [back|front] [alta|media|baja] Título
  Pasos: ... · Esperado: ... · Obtenido: ... · Archivo/test: ...
## Para probar a mano (PM)
- ...
## Para probar en Windows (PM), si la fase lo pide
- ...
```

### 4.4 Registro de avance (para no perder trabajo si un agente se queda sin contexto)

Un agente que se corta a mitad de una tarea no llega a mandar su informe. Para que otro pueda continuar, cada tarea de código lleva un **registro de avance** (ADR-020).

- **Dónde:** `docs/progreso/<ID>.md` (ej. `docs/progreso/F1-02.md`). Mientras el hook no habilite esa carpeta a los agentes, el despacho indica una ruta alternativa en la carpeta temporal de la sesión.
- **Al empezar:** si el registro ya existe, es una continuación: leerlo primero y seguir desde "Próximo paso", sin rehacer lo hecho.
- **Cuándo actualizarlo:** al terminar cada subpaso (un módulo, un grupo de tests que pasa) y antes de cualquier comando largo.
- **Cortar a tiempo:** si el agente ya trabajó mucho y falta bastante, corta en un punto estable (lo hecho con tests pasando, registro al día) y termina con estado `PARCIAL`. El líder despacha un **agente nuevo** que continúa desde el registro.
- **Al aprobar la tarea:** el líder borra el registro (lo que importa queda en el informe y en git).

```
# <ID> — <título>
## Estado: EN CURSO | PARCIAL | TERMINADO
## Hecho (con archivos)
## Pendiente (en orden)
## Próximo paso (concreto: archivo, función, test)
## Decisiones de nivel 1 tomadas
## Cómo verificar lo hecho (comandos)
```

---

## 5. Configuración (`.claude/`)

| Archivo | Para qué |
|---|---|
| `.claude/settings.json` | La sesión principal arranca como `lider-tecnico`; solo el líder puede despachar agentes; permisos de comandos; prohíbe `git push` |
| `.claude/agents/lider-tecnico.md` | Opus 5.5, effort medio. Solo puede editar `docs/`, `PLAN.md`, `CLAUDE.md` |
| `.claude/agents/backend.md` | Sonnet 5.5. Core, arranque, scripts, empaquetado |
| `.claude/agents/frontend.md` | Sonnet 5.5. UI |
| `.claude/agents/qa.md` | Sonnet 5.5. Fixtures, integración, informes |
| `.claude/hooks/limitar_rutas.py` | Bloquea `Edit`/`Write` fuera de las rutas permitidas de cada agente |

> **Effort:** el `effort: medium` del líder está en su archivo de agente. Ese valor tiene la prioridad más baja: lo pisan `/effort`, `--effort`, la variable `CLAUDE_CODE_EFFORT_LEVEL` y `effortLevel` en tu `~/.claude/settings.json`. Si tenés alguno de esos configurado, sacalo o arrancá con `claude --effort medium`.

> **Hooks:** usan `python3` (desarrollo en Linux). Si algún día se trabaja con Claude Code desde Windows, cambiar `python3` por `python` en los `command` de los agentes.

---

## 6. Costos y ritmo

- **Opus cuesta más por token** que Sonnet. Por eso el líder trabaja con effort medio, no lee código y se apoya en los informes. El grueso de los tokens (código, tests, lecturas) va en Sonnet.
- Conviene **una sesión del líder por fase**. El estado queda en git, `docs/decisiones.md`, `docs/qa/` y la lista de tareas.
- **Contexto de los agentes:** cada tarea grande va a un **agente nuevo** (arranca con contexto vacío y es más barato que reanudar uno con mucho historial). Reanudar con `SendMessage` solo para correcciones chicas de la misma tarea. Los despachos citan secciones puntuales de `PLAN.md`; los agentes buscan con grep antes de abrir archivos enteros y corren los tests con `-q`. Ver el registro de avance (§4.4).
- Si los agentes de Sonnet se traban seguido en algo difícil (por ejemplo, el join multi-archivo o la estimación de tamaño), el líder puede proponerte (nivel 3) subir el effort de ese agente para esa tarea.

---

## 7. Alternativa: agent teams

Claude Code también tiene **agent teams**: varias sesiones coordinadas que se mandan mensajes entre sí. Permitiría que los agentes le pregunten al líder sin terminar su tarea. Es más complejo y consume más. **Recomendación:** empezar con subagentes y evaluar agent teams solo si el ida y vuelta de `BLOQUEADO` → reanudar se vuelve un cuello de botella.
