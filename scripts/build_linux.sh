#!/usr/bin/env bash
# Compila el ejecutable de Linux, lo prueba y lo empaqueta.
set -euo pipefail

RAIZ="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$RAIZ"

SKIP_TESTS=0; ONEFILE=0
error() { echo "ERROR: $*" >&2; exit 1; }
info() { echo "==> $*"; }

for arg in "$@"; do
    case "$arg" in
        --skip-tests) SKIP_TESTS=1 ;;
        --onefile) ONEFILE=1 ;;
        --help|-h)
            cat <<'AYUDA'
Uso: scripts/build_linux.sh [opciones]

  --skip-tests  No corre los tests antes de compilar
  --onefile     Genera un solo archivo en vez de una carpeta
  --help        Muestra esta ayuda
AYUDA
            exit 0 ;;
        *) error "Opción desconocida: $arg (usá --help)" ;;
    esac
done

info "1/6 Preparando el entorno"
scripts/dev.sh --setup

VPY=".venv/bin/python"
VERSION="$("$VPY" -c 'import excel_tool; print(excel_tool.__version__)')"

if [ "$SKIP_TESTS" = 1 ]; then
    info "2/6 Tests saltados (--skip-tests)"
else
    info "2/6 Corriendo tests"
    scripts/dev.sh --tests
fi

info "3/6 Limpiando build/ y dist/"
rm -rf build dist

info "4/6 Compilando con PyInstaller"
if [ "$ONEFILE" = 1 ]; then export EXCEL_TOOL_ONEFILE=1; else export EXCEL_TOOL_ONEFILE=0; fi
"$VPY" -m PyInstaller --noconfirm --clean excel_tool.spec || error "Falló la compilación."

if [ "$ONEFILE" = 1 ]; then
    EXE="dist/ExcelJoinTool"
    CONTENIDO="ExcelJoinTool"
else
    EXE="dist/ExcelJoinTool/ExcelJoinTool"
    CONTENIDO="ExcelJoinTool"
fi
[ -f "$EXE" ] || error "No se generó el ejecutable esperado: $EXE"

info "5/6 Probando el ejecutable (--self-test)"
if [ -n "${DISPLAY:-}" ]; then
    "$EXE" --self-test || error "El ejecutable no pasó el --self-test. Revisá el log en ~/.local/share/ExcelJoinTool/logs."
elif command -v xvfb-run >/dev/null 2>&1; then
    xvfb-run -a "$EXE" --self-test || error "El ejecutable no pasó el --self-test. Revisá el log en ~/.local/share/ExcelJoinTool/logs."
else
    error "No hay pantalla ni xvfb-run para probar el ejecutable (sudo apt install xvfb)."
fi

info "6/6 Empaquetando"
SALIDA="dist/ExcelJoinTool-${VERSION}-linux.tar.gz"
tar -czf "$SALIDA" -C dist "$CONTENIDO"
TAM="$(du -h "$SALIDA" | cut -f1)"
SHA="$(sha256sum "$SALIDA" | cut -d' ' -f1)"
echo
echo "Listo: $SALIDA"
echo "Tamaño: $TAM"
echo "SHA256: $SHA"
