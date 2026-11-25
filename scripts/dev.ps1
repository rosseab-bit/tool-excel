# Entorno de desarrollo de Excel Join Tool (Windows). Compatible con Windows PowerShell 5.1.
param(
    [switch]$Setup,
    [switch]$Tests,
    [switch]$Recreate,
    [switch]$NoDev,
    [switch]$Lock,
    [switch]$Help,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$AppArgs = @()
)

$ErrorActionPreference = 'Stop'
$Raiz = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $Raiz

function Info([string]$Mensaje) { Write-Host "==> $Mensaje" }

function Fallar([string]$Mensaje) {
    [Console]::Error.WriteLine("ERROR: $Mensaje")
    exit 1
}

function Mostrar-Ayuda {
    Write-Host @'
Uso: scripts\dev.bat [opciones]    (o: powershell -File scripts\dev.ps1 [opciones])

  (sin opciones)  Prepara el entorno y abre la app (tambien con doble clic)
  -Setup          Solo prepara el entorno
  -Tests          Corre pytest y ruff (no reformatea nada)
  -Recreate       Borra .venv y lo crea de cero
  -NoDev          Instala solo lo necesario para usar la app
                  (sin pytest, ruff ni pyinstaller)
  -Lock           Regenera constraints.txt con las versiones instaladas
  -AppArgs <args> Pasa argumentos a la app (ej.: -AppArgs --self-test)
  -Help           Muestra esta ayuda
'@
}

if ($Help) { Mostrar-Ayuda; exit 0 }

if ($NoDev -and $Lock) { Fallar "-Lock necesita el entorno completo: no se puede combinar con -NoDev." }
if ($NoDev -and $Tests) { Fallar "-Tests necesita las herramientas de desarrollo: no se puede combinar con -NoDev." }

# Ejecuta un programa nativo sin que los avisos por stderr corten el script.
# Devuelve el codigo de salida; la salida estandar queda en $script:Salida.
function Ejecutar-Silencioso([string]$Exe, [string[]]$Argumentos) {
    $previo = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $script:Salida = & $Exe @Argumentos 2>$null
        return $LASTEXITCODE
    } catch {
        $script:Salida = $null
        return 1
    } finally {
        $ErrorActionPreference = $previo
    }
}

# --- Buscar Python >= 3.12 -----------------------------------------------------
$CodigoVersion = 'import sys; print(sys.executable); sys.exit(0 if sys.version_info >= (3, 12) else 1)'

# Devuelve $true si el Python (exe + argumentos previos) funciona y es >= 3.12.
# Descarta el alias falso de la Microsoft Store: no ejecuta codigo y no imprime la ruta.
function Probar-Python([string]$Exe, [string[]]$Previos) {
    $cmd = Get-Command $Exe -ErrorAction SilentlyContinue
    if (-not $cmd) { return $false }
    $argumentos = @($Previos) + @('-c', $CodigoVersion)
    $codigo = Ejecutar-Silencioso $cmd.Source $argumentos
    if ($codigo -ne 0) { return $false }
    $ruta = [string]($script:Salida | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($ruta)) { return $false }
    if ($ruta -match '\\WindowsApps\\python(3(\.\d+)?)?\.exe$') { return $false }
    $script:PyRuta = $ruta
    return $true
}

$PyExe = $null
$PyPrevios = @()
$script:PyRuta = ''
if ($env:PYTHON) {
    if (-not (Get-Command $env:PYTHON -ErrorAction SilentlyContinue)) {
        Fallar "La variable PYTHON apunta a '$($env:PYTHON)', que no existe."
    }
    if (-not (Probar-Python $env:PYTHON @())) {
        Fallar "'$($env:PYTHON)' no es Python 3.12 o mas nuevo (o no se puede ejecutar)."
    }
    $PyExe = $env:PYTHON
} else {
    $candidatos = @(
        @{ Exe = 'py'; Previos = @('-3.13') },
        @{ Exe = 'py'; Previos = @('-3.12') },
        @{ Exe = 'py'; Previos = @('-3') },
        @{ Exe = 'python'; Previos = @() }
    )
    foreach ($c in $candidatos) {
        if (Probar-Python $c.Exe $c.Previos) {
            $PyExe = $c.Exe
            $PyPrevios = $c.Previos
            break
        }
    }
    if (-not $PyExe) {
        Fallar ("No encontre Python 3.12 o mas nuevo. Instalalo desde https://www.python.org/downloads/ " +
            "(marcando 'Add python.exe to PATH' y 'tcl/tk and IDLE') o indica uno con la variable PYTHON.")
    }
}
Info "Usando Python: $script:PyRuta"

$codigo = Ejecutar-Silencioso $PyExe (@($PyPrevios) + @('-c', 'import tkinter'))
if ($codigo -ne 0) {
    Fallar ("Falta tkinter, que la app necesita para mostrar ventanas. Reinstala Python desde " +
        "https://www.python.org/downloads/ marcando la opcion 'tcl/tk and IDLE'.")
}

$codigo = Ejecutar-Silencioso $PyExe (@($PyPrevios) + @('-c', 'import venv, ensurepip'))
if ($codigo -ne 0) {
    Fallar "Falta el soporte para crear entornos (venv). Reinstala Python desde python.org con la instalacion completa."
}

# --- Entorno virtual -------------------------------------------------------------
$VenvDir = Join-Path $Raiz '.venv'
$VPy = Join-Path $VenvDir 'Scripts\python.exe'

if ($Recreate -and (Test-Path -LiteralPath $VenvDir)) {
    Info "Borrando .venv"
    Remove-Item -LiteralPath $VenvDir -Recurse -Force
}

$venvSirve = $false
if (Test-Path -LiteralPath $VPy) {
    $venvSirve = ((Ejecutar-Silencioso $VPy @('-c', 'import sys; sys.exit(0 if sys.version_info >= (3, 12) else 1)')) -eq 0)
}
if (-not $venvSirve) {
    if (Test-Path -LiteralPath $VenvDir) {
        Info "El .venv no sirve, lo recreo"
        Remove-Item -LiteralPath $VenvDir -Recurse -Force
    }
    Info "Creando .venv"
    $previos = @($PyPrevios)
    & $PyExe @previos -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) { Fallar "No se pudo crear el entorno virtual (.venv)." }
}

# --- Dependencias --------------------------------------------------------------------
$Modo = if ($NoDev) { 'app' } else { 'dev' }
$HashFile = Join-Path $VenvDir '.deps-hash'
$Pyproject = Join-Path $Raiz 'pyproject.toml'
$Constraints = Join-Path $Raiz 'constraints.txt'

function Calcular-Hash {
    $texto = "modo=$Modo`n" + (Get-Content -LiteralPath $Pyproject -Raw -Encoding UTF8)
    if (Test-Path -LiteralPath $Constraints) {
        $texto += (Get-Content -LiteralPath $Constraints -Raw -Encoding UTF8)
    } else {
        $texto += "sin-constraints"
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($texto))
    return ([System.BitConverter]::ToString($bytes) -replace '-', '').ToLower()
}

$Hash = Calcular-Hash
$alDia = $false
if (Test-Path -LiteralPath $HashFile) {
    $guardado = (Get-Content -LiteralPath $HashFile -Raw -Encoding UTF8).Trim()
    $alDia = ($guardado -eq $Hash)
}

if ($alDia) {
    Info "Dependencias al dia"
} else {
    Info "Instalando dependencias (modo: $Modo)"
    $objetivo = if ($NoDev) { '.' } else { '.[dev,build]' }
    $argsPip = @('-m', 'pip', 'install', '--disable-pip-version-check', '-e', $objetivo)
    if (Test-Path -LiteralPath $Constraints) {
        $argsPip += @('-c', 'constraints.txt')
    } else {
        [Console]::Error.WriteLine("AVISO: no hay constraints.txt; instalo sin versiones fijas. Conviene correr scripts\dev.bat -Lock.")
    }
    if (Test-Path -LiteralPath $HashFile) { Remove-Item -LiteralPath $HashFile -Force }
    & $VPy @argsPip
    if ($LASTEXITCODE -ne 0) { Fallar "Fallo la instalacion de dependencias." }
    Set-Content -LiteralPath $HashFile -Value $Hash -Encoding ASCII
}

# --- Lock ---------------------------------------------------------------------------------
if ($Lock) {
    Info "Generando constraints.txt"
    $congeladas = & $VPy -m pip freeze --exclude-editable --exclude excel-tool
    if ($LASTEXITCODE -ne 0) { Fallar "No se pudo leer la lista de versiones instaladas." }
    $lineas = @(
        '# Versiones fijas de las dependencias (se usan con: pip install -c constraints.txt).',
        '# Para regenerar este archivo: scripts/dev.sh --lock  (o scripts\dev.bat -Lock)'
    ) + @($congeladas)
    $contenido = ($lineas -join "`n") + "`n"
    [System.IO.File]::WriteAllText($Constraints, $contenido, (New-Object System.Text.UTF8Encoding($false)))
    Set-Content -LiteralPath $HashFile -Value (Calcular-Hash) -Encoding ASCII
}

# --- Tests ----------------------------------------------------------------------------------
if ($Tests) {
    Info "Corriendo tests"
    & $VPy -m pytest -q
    if ($LASTEXITCODE -ne 0) { Fallar "Fallaron los tests." }
    Info "Revisando el codigo con ruff"
    & $VPy -m ruff check .
    if ($LASTEXITCODE -ne 0) { Fallar "ruff encontro problemas en el codigo." }
    & $VPy -m ruff format --check .
    if ($LASTEXITCODE -ne 0) { Fallar "ruff dice que hay archivos sin formatear." }
    Info "Todo bien"
}

# --- App ------------------------------------------------------------------------------------------
if (($AppArgs.Count -gt 0) -or (-not $Setup -and -not $Tests -and -not $Lock)) {
    & $VPy -m excel_tool @AppArgs
    exit $LASTEXITCODE
}
exit 0
