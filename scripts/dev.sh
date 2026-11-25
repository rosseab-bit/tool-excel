#!/usr/bin/env bash
# Entorno de desarrollo de Excel Join Tool (Linux).
set -euo pipefail

RAIZ="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$RAIZ"

SETUP=0; TESTS=0; RECREATE=0; NODEV=0; LOCK=0; APP_ARGS=()

ayuda() {
    cat <<'AYUDA'
Uso: scripts/dev.sh [opciones] [-- argumentos para la app]

  (sin opciones)  Prepara el entorno y abre la app
  --setup         Solo prepara el entorno
  --tests         Corre pytest y ruff (no reformatea nada)
  --recreate      Borra .venv y lo crea de cero
  --no-dev        Instala solo lo necesario para usar la app
                  (sin pytest, ruff ni pyinstaller)
  --lock          Regenera constraints.txt con las versiones instaladas
  -- <args>       Pasa argumentos a la app (ej.: -- --self-test)
  --help          Muestra esta ayuda
AYUDA
}

error() { echo "ERROR: $*" >&2; exit 1; }
info() { echo "==> $*"; }

while [ $# -gt 0 ]; do
    case "$1" in
        --setup) SETUP=1 ;;
        --tests) TESTS=1 ;;
        --recreate) RECREATE=1 ;;
        --no-dev) NODEV=1 ;;
        --lock) LOCK=1 ;;
        --help|-h) ayuda; exit 0 ;;
        --) shift; APP_ARGS=("$@"); break ;;
        *) echo "Opción desconocida: $1" >&2; ayuda >&2; exit 2 ;;
    esac
    shift
done

[ "$NODEV" = 1 ] && [ "$LOCK" = 1 ] && error "--lock necesita el entorno completo: no se puede combinar con --no-dev."
[ "$NODEV" = 1 ] && [ "$TESTS" = 1 ] && error "--tests necesita las herramientas de desarrollo: no se puede combinar con --no-dev."

# --- Buscar Python >= 3.12 ---------------------------------------------------
version_ok() {
    "$1" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)' >/dev/null 2>&1
}

PY=""
if [ -n "${PYTHON:-}" ]; then
    command -v "$PYTHON" >/dev/null 2>&1 || error "La variable PYTHON apunta a '$PYTHON', que no existe."
    PY="$PYTHON"
    version_ok "$PY" || error "'$PYTHON' no es Python 3.12 o más nuevo."
else
    for cand in python3.13 python3.12 python3; do
        if command -v "$cand" >/dev/null 2>&1 && version_ok "$cand"; then PY="$cand"; break; fi
    done
    [ -n "$PY" ] || error "No encontré Python 3.12 o más nuevo. Instalalo (por ejemplo: sudo apt install python3.12) o indicá uno con PYTHON=/ruta/a/python."
fi
info "Usando Python: $("$PY" -c 'import sys; print(sys.version.split()[0], sys.executable)')"

if ! "$PY" -c 'import tkinter' >/dev/null 2>&1; then
    error "Falta tkinter, que la app necesita para mostrar ventanas.
  Ubuntu/Debian: sudo apt install python3-tk
  Fedora:        sudo dnf install python3-tkinter"
fi

if ! "$PY" -c 'import venv, ensurepip' >/dev/null 2>&1; then
    MENOR="$("$PY" -c 'import sys; print("%d.%d" % sys.version_info[:2])')"
    error "Falta el soporte para crear entornos (venv).
  Ubuntu/Debian: sudo apt install python${MENOR}-venv
  Fedora:        sudo dnf install python3-libs"
fi

# --- Entorno virtual ----------------------------------------------------------
if [ "$RECREATE" = 1 ] && [ -d .venv ]; then
    info "Borrando .venv"
    rm -rf .venv
fi
VPY=".venv/bin/python"
if [ ! -x "$VPY" ] || ! version_ok "$VPY"; then
    [ -e .venv ] && { info "El .venv no sirve, lo recreo"; rm -rf .venv; }
    info "Creando .venv"
    "$PY" -m venv .venv || error "No se pudo crear el entorno virtual (.venv)."
fi

# --- Dependencias ----------------------------------------------------------------
if [ "$NODEV" = 1 ]; then MODO="app"; else MODO="dev"; fi
HASH_FILE=".venv/.deps-hash"

calcular_hash() {
    {
        echo "modo=$MODO"
        cat pyproject.toml
        if [ -f constraints.txt ]; then cat constraints.txt; else echo "sin-constraints"; fi
    } | sha256sum | cut -d' ' -f1
}

HASH="$(calcular_hash)"
if [ -f "$HASH_FILE" ] && [ "$(cat "$HASH_FILE")" = "$HASH" ]; then
    info "Dependencias al día"
else
    info "Instalando dependencias (modo: $MODO)"
    ARGS_C=()
    if [ -f constraints.txt ]; then
        ARGS_C=(-c constraints.txt)
    else
        echo "AVISO: no hay constraints.txt; instalo sin versiones fijas. Conviene correr scripts/dev.sh --lock." >&2
    fi
    if [ "$NODEV" = 1 ]; then OBJ="."; else OBJ=".[dev,build]"; fi
    rm -f "$HASH_FILE"
    "$VPY" -m pip install --disable-pip-version-check -e "$OBJ" "${ARGS_C[@]}" \
        || error "Falló la instalación de dependencias."
    echo "$HASH" > "$HASH_FILE"
fi

# --- Lock ------------------------------------------------------------------------
if [ "$LOCK" = 1 ]; then
    info "Generando constraints.txt"
    {
        echo "# Versiones fijas de las dependencias (se usan con: pip install -c constraints.txt)."
        echo "# Para regenerar este archivo: scripts/dev.sh --lock"
        "$VPY" -m pip freeze --exclude-editable --exclude excel-tool
    } > constraints.txt.tmp
    mv constraints.txt.tmp constraints.txt
    calcular_hash > "$HASH_FILE"
fi

# --- Tests --------------------------------------------------------------------------
if [ "$TESTS" = 1 ]; then
    info "Corriendo tests"
    if [ -n "${DISPLAY:-}" ]; then
        "$VPY" -m pytest -q
    elif command -v xvfb-run >/dev/null 2>&1; then
        info "Sin pantalla: uso xvfb-run"
        xvfb-run -a "$VPY" -m pytest -q
    else
        echo "AVISO: no hay pantalla ni xvfb-run; los tests de pantalla se van a saltar. (sudo apt install xvfb)" >&2
        "$VPY" -m pytest -q
    fi
    info "Revisando el código con ruff"
    .venv/bin/ruff check .
    .venv/bin/ruff format --check .
    info "Todo bien"
fi

# --- App ------------------------------------------------------------------------------
if [ ${#APP_ARGS[@]} -gt 0 ] || { [ "$SETUP" = 0 ] && [ "$TESTS" = 0 ] && [ "$LOCK" = 0 ]; }; then
    exec "$VPY" -m excel_tool "${APP_ARGS[@]}"
fi
