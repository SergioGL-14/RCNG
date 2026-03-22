#requires -version 5.1

# Lanzador minimo.
# Busca el LazyWinAdmin*.ps1 mas reciente en la misma carpeta y lo ejecuta en
# el proceso actual. Se usa sobre todo como apoyo en desarrollo o revision.

# Resolver la carpeta del script con varios fallback por si cambia el contexto.
if ($PSCommandPath) {
    $ScriptDir = Split-Path -Parent $PSCommandPath
} elseif ($MyInvocation.MyCommand.Path) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    $ScriptDir = $PWD.Path
}

# Si por cualquier motivo la ruta queda vacia, usar el directorio actual.
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
    $ScriptDir = Get-Location | Select-Object -ExpandProperty Path
}

# Trabajar desde la carpeta del lanzador evita rutas relativas inconsistentes.
Set-Location -Path $ScriptDir

$EXE_PATTERN = "LazyWinAdmin*.ps1"

# Buscar el script principal mas reciente disponible.
Write-Host "Buscando LazyWinAdmin en: $ScriptDir" -ForegroundColor Cyan

$lazyScript = Get-ChildItem -Path $ScriptDir -Filter $EXE_PATTERN -ErrorAction SilentlyContinue | 
              Sort-Object LastWriteTime -Descending | 
              Select-Object -First 1

if (-not $lazyScript) {
    Write-Error "No se encontro ningun $EXE_PATTERN en $ScriptDir"
    Write-Host "Contenido del directorio:" -ForegroundColor Yellow
    Get-ChildItem -Path $ScriptDir | Format-Table Name, LastWriteTime -AutoSize
    throw "No se encontro ningun LazyWinAdmin*.ps1 en $ScriptDir"
}

Write-Host "Encontrado: $($lazyScript.FullName)" -ForegroundColor Green

# Ejecutar el script encontrado en el mismo proceso.
& $lazyScript.FullName
