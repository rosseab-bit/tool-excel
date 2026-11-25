#!/usr/bin/env python3
"""Hook PreToolUse: limita qué archivos puede editar cada rol del equipo.

Uso (desde el frontmatter de cada agente):
    python3 "$CLAUDE_PROJECT_DIR/.claude/hooks/limitar_rutas.py" <rol>

- Edit / Write / NotebookEdit: bloquea rutas del repo fuera de las permitidas para el rol.
  Las rutas fuera del repo (scratchpad, memoria) no se controlan.
- Bash (solo backend, frontend y qa): bloquea comandos git que cambian el historial
  o descartan trabajo (commit, push, reset, ...). Commitea el líder.

Código de salida 2 = bloquear (el mensaje de stderr le llega al agente).
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path, PurePosixPath

# Prefijos permitidos (relativos a la raíz del repo, con "/"). Un prefijo que termina
# en "/" es una carpeta; si no, es un archivo exacto.
PERMITIDAS: dict[str, list[str]] = {
    "lider-tecnico": ["docs/", "PLAN.md", "CLAUDE.md"],
    "backend": [
        "src/excel_tool/",  # menos ui/ y assets/ (ver EXCLUIDAS)
        "tests/core/",
        "tests/__init__.py",
        "scripts/",  # menos make_fixtures.py (ver EXCLUIDAS)
        "pyproject.toml",
        "constraints.txt",
        "excel_tool.spec",
        "packaging/",
        "README.md",
        ".gitignore",
        ".gitattributes",
    ],
    "frontend": ["src/excel_tool/ui/", "src/excel_tool/assets/", "tests/ui/"],
    "qa": [
        "tests/integracion/",
        "tests/fixtures/",
        "tests/conftest.py",
        "scripts/make_fixtures.py",
        "docs/qa/",
    ],
}

EXCLUIDAS: dict[str, list[str]] = {
    "backend": ["src/excel_tool/ui/", "src/excel_tool/assets/", "scripts/make_fixtures.py"],
}

# Nadie crea CI (CLAUDE.md: todo local).
PROHIBIDAS_PARA_TODOS = [".github/"]

GIT_PROHIBIDO = re.compile(
    r"\bgit\s+(?:-C\s+\S+\s+)?"
    r"(commit|push|reset|rebase|merge|checkout|switch|restore|clean|stash|tag|branch\s+-[dD])\b"
)


def _coincide(rel: str, prefijo: str) -> bool:
    if prefijo.endswith("/"):
        return rel.startswith(prefijo)
    return rel == prefijo


def _bloquear(mensaje: str) -> None:
    print(mensaje, file=sys.stderr)
    sys.exit(2)


def main() -> None:
    rol = sys.argv[1] if len(sys.argv) > 1 else ""
    if rol not in PERMITIDAS:
        _bloquear(f"limitar_rutas.py: rol desconocido '{rol}'. Revisar el frontmatter del agente.")

    datos = json.load(sys.stdin)
    herramienta = datos.get("tool_name", "")
    entrada = datos.get("tool_input", {}) or {}

    if herramienta == "Bash":
        if rol == "lider-tecnico":
            return
        comando = entrada.get("command", "")
        if encontrado := GIT_PROHIBIDO.search(comando):
            _bloquear(
                f"Bloqueado: los agentes no ejecutan 'git {encontrado.group(1)}'. "
                "Commitea el líder cuando QA aprueba (CLAUDE.md). Dejá los cambios en el "
                "árbol de trabajo y mencionálos en el informe."
            )
        return

    ruta = entrada.get("file_path") or entrada.get("notebook_path")
    if not ruta:
        return

    raiz = Path(os.environ.get("CLAUDE_PROJECT_DIR") or datos.get("cwd") or ".").resolve()
    destino = Path(ruta)
    if not destino.is_absolute():
        destino = Path(datos.get("cwd") or raiz) / destino
    destino = destino.resolve()

    try:
        rel = PurePosixPath(destino.relative_to(raiz)).as_posix()
    except ValueError:
        return  # fuera del repo: scratchpad, memoria, etc.

    if any(_coincide(rel, p) for p in PROHIBIDAS_PARA_TODOS):
        _bloquear(f"Bloqueado: '{rel}'. El proyecto no usa CI ni GitHub Actions (CLAUDE.md).")

    excluida = any(_coincide(rel, p) for p in EXCLUIDAS.get(rol, []))
    permitida = any(_coincide(rel, p) for p in PERMITIDAS[rol])
    if permitida and not excluida:
        return

    _bloquear(
        f"Bloqueado: el rol '{rol}' no puede editar '{rel}'. "
        f"Rutas permitidas: {', '.join(PERMITIDAS[rol])}. "
        "Si necesitás cambiar ese archivo, terminá con estado BLOQUEADO y pedíselo al líder "
        "(docs/equipo.md §2)."
    )


if __name__ == "__main__":
    main()
