# Compila el .exe de Windows, lo prueba y lo empaqueta. Compatible con Windows PowerShell 5.1.
param(
    [switch]$SkipTests,
    [switch]$OneFile,
    [switch]$Help
)

$ErrorActionPreference = 'Stop'
$Raiz = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $Raiz

function Info([string]$Mensaje) { Write-Host "==> $Mensaje" }

function Fallar([string]$Mensaje) {
    [Console]::Error.WriteLine("ERROR: $Mensaje")
    exit 1
}

if ($Help) {
    Write-Host @'
Uso: scripts\build_windows.bat [opciones]

  -SkipTests  No corre los tests antes de compilar
  -OneFile    Genera un solo .exe en vez de una carpeta
  -Help       Muestra esta ayuda
'@
    exit 0
}

$Dev = Join-Path $PSScriptRoot 'dev.ps1'
function Llamar-Dev([string]$Opcion) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $Dev $Opcion
    if ($LASTEXITCODE -ne 0) { Fallar "dev.ps1 $Opcion fallo." }
}

Info "1/6 Preparando el entorno"
Llamar-Dev '-Setup'

$VPy = Join-Path $Raiz '.venv\Scripts\python.exe'
$Version = (& $VPy -c 'import excel_tool; print(excel_tool.__version__)')
if ($LASTEXITCODE -ne 0 -or -not $Version) { Fallar "No se pudo leer la version de la app." }
$Version = ([string]$Version).Trim()

if ($SkipTests) {
    Info "2/6 Tests saltados (-SkipTests)"
} else {
    Info "2/6 Corriendo tests"
    Llamar-Dev '-Tests'
}

Info "3/6 Limpiando build\ y dist\"
foreach ($carpeta in @('build', 'dist')) {
    if (Test-Path -LiteralPath $carpeta) { Remove-Item -LiteralPath $carpeta -Recurse -Force }
}

Info "4/6 Compilando con PyInstaller"
$env:EXCEL_TOOL_ONEFILE = if ($OneFile) { '1' } else { '0' }
& $VPy -m PyInstaller --noconfirm --clean excel_tool.spec
if ($LASTEXITCODE -ne 0) { Fallar "Fallo la compilacion." }

if ($OneFile) {
    $Exe = Join-Path $Raiz 'dist\ExcelJoinTool.exe'
    $Contenido = $Exe
} else {
    $Exe = Join-Path $Raiz 'dist\ExcelJoinTool\ExcelJoinTool.exe'
    $Contenido = Join-Path $Raiz 'dist\ExcelJoinTool'
}
if (-not (Test-Path -LiteralPath $Exe)) { Fallar "No se genero el ejecutable esperado: $Exe" }

Info "5/6 Probando el ejecutable (--self-test)"
# El .exe no tiene consola: el resultado es el codigo de salida.
$proceso = Start-Process -FilePath $Exe -ArgumentList '--self-test' -Wait -PassThru
if ($null -eq $proceso.ExitCode -or $proceso.ExitCode -ne 0) {
    Fallar "El ejecutable no paso el --self-test (codigo $($proceso.ExitCode)). Revisa el log en %LOCALAPPDATA%\ExcelJoinTool\logs."
}
Info "Self-test OK"

Info "6/6 Empaquetando"
$Salida = Join-Path $Raiz "dist\ExcelJoinTool-$Version-windows.zip"
Compress-Archive -LiteralPath $Contenido -DestinationPath $Salida -Force
$tamano = [math]::Round((Get-Item -LiteralPath $Salida).Length / 1MB, 1)
$sha = (Get-FileHash -LiteralPath $Salida -Algorithm SHA256).Hash.ToLower()
Write-Host ""
Write-Host "Listo: $Salida"
Write-Host "Tamano: $tamano MB"
Write-Host "SHA256: $sha"
exit 0
