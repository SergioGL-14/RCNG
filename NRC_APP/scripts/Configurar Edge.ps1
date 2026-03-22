<#
.SYNOPSIS
    Configura el Modo IE en Microsoft Edge

.DESCRIPTION
    1. Permitir que los sitios se carguen en modo IE (Compatibilidad IE)
    2. Mostrar botón de Modo IE en la barra de herramientas
    
.PARAMETER ComputerName
    Nombre del equipo remoto donde se aplicará la configuración. Si no se especifica, se aplica localmente.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName
)

# Función para pausar antes de salir
function Pause-Exit {
    param([int]$ExitCode = 0)
    Write-Host "`nPresione cualquier tecla para salir..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit $ExitCode
}

$ErrorActionPreference = "Stop"
$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
$isRemote = -not [string]::IsNullOrWhiteSpace($ComputerName)

# Si es ejecución remota, usar PsExec
if ($isRemote) {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  Configuración REMOTA Modo IE para Edge" -ForegroundColor Cyan
    Write-Host "  Equipo: $ComputerName" -ForegroundColor Yellow
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Verificar conectividad
    Write-Host "[0/3] Verificando conectividad con '$ComputerName'..." -ForegroundColor Yellow
    try {
        if (-not (Test-Connection -ComputerName $ComputerName -Count 2 -Quiet)) {
            Write-Host "    ✗ No se pudo conectar con el equipo" -ForegroundColor Red
            Pause-Exit 1
        }
        Write-Host "    ✓ Equipo accesible`n" -ForegroundColor Green
    } catch {
        Write-Host "    ✗ Error al verificar conectividad: $_" -ForegroundColor Red
        Pause-Exit 1
    }
    
    # Ruta a PsExec
    $psexecPath = Join-Path $PSScriptRoot "..\tools\PsExec.exe"
    
    if (-not (Test-Path $psexecPath)) {
        Write-Host "[ERROR] No se encontró PsExec.exe en: $psexecPath" -ForegroundColor Red
        Write-Host "Descargue PsExec de Sysinternals y colóquelo en la carpeta 'tools'" -ForegroundColor Yellow
        Pause-Exit 1
    }
} else {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  Configuración LOCAL Modo IE para Edge" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

# Cerrar Edge si está ejecutándose
if ($isRemote) {
    Write-Host "[1/3] Cerrando Microsoft Edge en '$ComputerName'..." -ForegroundColor Yellow
    try {
        $killResult = & $psexecPath -accepteula -s \\$ComputerName taskkill /F /IM msedge.exe 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "    ✓ Edge cerrado`n" -ForegroundColor Green
            Start-Sleep -Seconds 2
        } else {
            Write-Host "    • Edge no estaba ejecutándose o ya estaba cerrado`n" -ForegroundColor Gray
        }
    } catch {
        Write-Host "    • Edge no estaba ejecutándose`n" -ForegroundColor Gray
    }
} else {
    $edgeProcesses = Get-Process msedge -ErrorAction SilentlyContinue
    if ($edgeProcesses) {
        Write-Host "[1/3] Cerrando Microsoft Edge..." -ForegroundColor Yellow
        $edgeProcesses | Stop-Process -Force
        Start-Sleep -Seconds 3
        Write-Host "    ✓ Edge cerrado`n" -ForegroundColor Green
    } else {
        Write-Host "[1/3] Edge no está ejecutándose`n" -ForegroundColor Gray
    }
}

# Crear ruta de registro si no existe
if ($isRemote) {
    Write-Host "[2/3] Configurando registro en equipo remoto..." -ForegroundColor Green
    
    # Script que se ejecutará remotamente
    $remoteScript = @'
$ErrorActionPreference = 'Stop'
try {
    $regPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'
    
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    
    Set-ItemProperty -Path $regPath -Name 'InternetExplorerIntegrationLevel' -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $regPath -Name 'InternetExplorerIntegrationReloadInIEModeAllowed' -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $regPath -Name 'InternetExplorerModeToolbarButtonEnabled' -Value 1 -Type DWord -Force
    
    Write-Output 'CONFIGURACION_OK'
    exit 0
} catch {
    Write-Output "ERROR: $_"
    exit 1
}
'@
    
    # Generar nombre único para el archivo
    $randomName = "edge_config_" + (Get-Random -Minimum 1000 -Maximum 9999) + ".ps1"
    $remoteScriptPath = "C:\Windows\Temp\$randomName"
    $remoteScriptUNC = "\\$ComputerName\C`$\Windows\Temp\$randomName"
    
    try {
        # Verificar acceso al recurso compartido C$
        Write-Host "    Verificando acceso al equipo remoto..." -ForegroundColor Yellow
        if (-not (Test-Path "\\$ComputerName\C`$\Windows\Temp")) {
            Write-Host "    ✗ No se puede acceder a \\$ComputerName\C`$\Windows\Temp" -ForegroundColor Red
            Write-Host "    Verifique que tiene permisos de administrador en el equipo remoto" -ForegroundColor Yellow
            Pause-Exit 1
        }
        
        # Copiar script al equipo remoto
        Write-Host "    Copiando script al equipo remoto..." -ForegroundColor Yellow
        $remoteScript | Out-File -FilePath $remoteScriptUNC -Encoding UTF8 -Force
        
        # Verificar que se copió correctamente
        if (-not (Test-Path $remoteScriptUNC)) {
            Write-Host "    ✗ No se pudo copiar el script al equipo remoto" -ForegroundColor Red
            Pause-Exit 1
        }
        
        Write-Host "    ✓ Script copiado correctamente" -ForegroundColor Green
        
        # Ejecutar con PsExec usando la ruta LOCAL del equipo remoto
        Write-Host "    Ejecutando configuración remota..." -ForegroundColor Yellow
        
        $arguments = "\\$ComputerName -accepteula -s powershell.exe -ExecutionPolicy Bypass -NoProfile -File `"$remoteScriptPath`""
        
        Write-Host "    Debug - Comando: $psexecPath $arguments" -ForegroundColor Cyan
        
        $process = Start-Process -FilePath $psexecPath -ArgumentList $arguments -NoNewWindow -Wait -PassThru -RedirectStandardOutput "$env:TEMP\edge_output.txt" -RedirectStandardError "$env:TEMP\edge_error.txt"
        
        # Leer resultados
        $output = ""
        $errorOutput = ""
        
        if (Test-Path "$env:TEMP\edge_output.txt") {
            $output = Get-Content "$env:TEMP\edge_output.txt" -Raw -ErrorAction SilentlyContinue
            Remove-Item "$env:TEMP\edge_output.txt" -Force -ErrorAction SilentlyContinue
        }
        
        if (Test-Path "$env:TEMP\edge_error.txt") {
            $errorOutput = Get-Content "$env:TEMP\edge_error.txt" -Raw -ErrorAction SilentlyContinue
            Remove-Item "$env:TEMP\edge_error.txt" -Force -ErrorAction SilentlyContinue
        }
        
        Write-Host "    Debug - Código de salida: $($process.ExitCode)" -ForegroundColor Cyan
        Write-Host "    Debug - Salida: $output" -ForegroundColor Cyan
        if ($errorOutput) {
            Write-Host "    Debug - Error: $errorOutput" -ForegroundColor Cyan
        }
        
        if ($output -match "CONFIGURACION_OK" -or $process.ExitCode -eq 0) {
            Write-Host "    ✓ Modo IE habilitado" -ForegroundColor Green
            Write-Host "    ✓ Permitir recargar sitios en Modo IE: Activado" -ForegroundColor Green
            Write-Host "    ✓ Botón en toolbar: Activado`n" -ForegroundColor Green
        } else {
            Write-Host "    ✗ Error al configurar registro" -ForegroundColor Red
            if ($errorOutput) {
                Write-Host "    $errorOutput" -ForegroundColor Yellow
            }
            Pause-Exit 1
        }
        
    } catch {
        Write-Host "    ✗ Error ejecutando configuración: $_" -ForegroundColor Red
        Write-Host "    Detalle: $($_.Exception.Message)" -ForegroundColor Yellow
        Pause-Exit 1
    } finally {
        # Limpiar script remoto
        if (Test-Path $remoteScriptUNC) {
            Remove-Item $remoteScriptUNC -Force -ErrorAction SilentlyContinue
        }
    }
} else {
    if (!(Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }

    # Habilitar modo IE
    Write-Host "[2/3] Habilitando Modo Internet Explorer..." -ForegroundColor Green

    try {
        Set-ItemProperty -Path $registryPath -Name "InternetExplorerIntegrationLevel" -Value 1 -Type DWord -Force
        Set-ItemProperty -Path $registryPath -Name "InternetExplorerIntegrationReloadInIEModeAllowed" -Value 1 -Type DWord -Force
        Write-Host "    ✓ Modo IE habilitado" -ForegroundColor Green
        Write-Host "    ✓ Permitir recargar sitios en Modo IE: Activado`n" -ForegroundColor Green
    } catch {
        Write-Host "    ✗ Error: $_" -ForegroundColor Red
        Pause-Exit 1
    }

    # Mostrar botón en toolbar
    Write-Host "[3/3] Activando botón de Modo IE..." -ForegroundColor Green

    try {
        Set-ItemProperty -Path $registryPath -Name "InternetExplorerModeToolbarButtonEnabled" -Value 1 -Type DWord -Force
        Write-Host "    ✓ Botón activado`n" -ForegroundColor Green
    } catch {
        Write-Host "    ✗ Error: $_" -ForegroundColor Red
        Pause-Exit 1
    }
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  ✓ Configuración completada" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Cyan

if (-not $isRemote) {
    # Mostrar políticas aplicadas (solo en local)
    try {
        $props = Get-ItemProperty -Path $registryPath -ErrorAction Stop
        Write-Host "Políticas aplicadas:" -ForegroundColor Cyan
        Write-Host "  • Modo IE: " -NoNewline; Write-Host "Habilitado" -ForegroundColor Green
        Write-Host "  • Permitir recargar en Modo IE: " -NoNewline; Write-Host "Activado" -ForegroundColor Green
        Write-Host "  • Botón en toolbar: " -NoNewline; Write-Host "Activado`n" -ForegroundColor Green
    } catch {
        Write-Host "No se pudieron leer las políticas`n" -ForegroundColor Red
    }

    # Abrir Edge automáticamente
    Write-Host "Abriendo Microsoft Edge..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    Start-Process "msedge.exe" "edge://policy/"

    Write-Host "`n✓ Script finalizado" -ForegroundColor Green
    Write-Host "Verifica en edge://policy/ que las políticas se aplicaron correctamente`n" -ForegroundColor Gray
} else {
    Write-Host "`n✓ Configuración aplicada en '$ComputerName'" -ForegroundColor Green
    Write-Host "El usuario deberá reiniciar Edge para que los cambios surtan efecto`n" -ForegroundColor Yellow
    Pause-Exit 0
}