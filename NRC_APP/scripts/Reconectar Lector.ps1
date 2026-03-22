param (
    [string]$ComputerName
)

# Función para logging
function Write-Log {
    param ([string]$Message)
    $LogFile = "C:\Aplicacions Soporte\NRC_APP\logs\Reconectar_Lector.log"
    $LogDir = Split-Path $LogFile
    if (-not (Test-Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp - $Message`n" | Out-File -FilePath $LogFile -Append
    Write-Host $Message
}

# Verificar si se proporcionó el nombre del equipo
if (-not $ComputerName) {
    Write-Log "Error: No se proporcionó el nombre del equipo remoto."
    Write-Host "`nPresione cualquier tecla para salir..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Verificar conectividad al equipo remoto
if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
    Write-Log "Error: No se pudo alcanzar el equipo remoto '$ComputerName'."
    Write-Host "`nPresione cualquier tecla para salir..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Log "Detectando lector de SmartCard en '$ComputerName'..."

# Detectar TODOS los lectores de SmartCard remotamente usando WMI
$lectores = @(Get-WmiObject -Class Win32_PnPEntity -ComputerName $ComputerName | Where-Object {
    $_.ClassGuid -eq "{50dd5230-ba8a-11d1-bf5d-0000f805f530}" -or
    $_.Name -match "SmartCard|Tarjeta Inteligente|Lector"
})

if ($lectores.Count -eq 0) {
    Write-Log "No se encontró un lector de SmartCard. Iniciando búsqueda de cambios de hardware..."
    
    try {
        Invoke-WmiMethod -Path "root\cimv2:Win32_PnPEntity.DeviceID='HTREE\\ROOT\\0'" -Name ScanForHardwareChanges -ComputerName $ComputerName -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Búsqueda de hardware completada. Esperando 3 segundos..."
        Start-Sleep -Seconds 3
        
        $lectores = @(Get-WmiObject -Class Win32_PnPEntity -ComputerName $ComputerName | Where-Object {
            $_.ClassGuid -eq "{50dd5230-ba8a-11d1-bf5d-0000f805f530}" -or
            $_.Name -match "SmartCard|Tarjeta Inteligente|Lector"
        })
        
        if ($lectores.Count -eq 0) {
            Write-Log "Error: No se encontró un lector de SmartCard en '$ComputerName' después del escaneo de hardware."
            Write-Host "`nPresione cualquier tecla para salir..." -ForegroundColor Yellow
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit 1
        }
        
        Write-Log "Lectores detectados después del escaneo de hardware."
    } catch {
        Write-Log "Error durante el escaneo de hardware: $($_.Exception.Message)"
        Write-Log "Error: No se encontró un lector de SmartCard en '$ComputerName'."
        Write-Host "`nPresione cualquier tecla para salir..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
}

Write-Log "Se encontraron $($lectores.Count) dispositivo(s) de SmartCard:"
foreach ($lec in $lectores) {
    Write-Log "  - $($lec.Name) [Estado: $(if ($lec.ConfigManagerErrorCode -eq 0) { 'OK' } elseif ($lec.ConfigManagerErrorCode -eq 22) { 'DESHABILITADO' } else { 'ERROR ' + $lec.ConfigManagerErrorCode })]"
}

# Proceso de reconexión de TODOS los lectores
# Proceso de reconexión de TODOS los lectores
try {
    Write-Log ""
    Write-Log "Iniciando proceso de reconexión..."
    Write-Log ""
    
    $exitosos = 0
    $errores = 0
    
    # Recorrer cada lector
    foreach ($lector in $lectores) {
        $deviceID = $lector.DeviceID
        $deviceName = $lector.Name
        
        Write-Log "----------------------------------------"
        Write-Log "Procesando: $deviceName"
        
        # Verificar estado inicial
        $configManagerErrorCode = $lector.ConfigManagerErrorCode
        $isDisabled = ($configManagerErrorCode -eq 22)
        
        Write-Log "Estado inicial: $(if ($isDisabled) { 'DESHABILITADO' } else { 'HABILITADO' }) (ErrorCode: $configManagerErrorCode)"
        
        try {
            if (-not $isDisabled) {
                # Deshabilitar
                Write-Log "Deshabilitando..."
                $disableResult = $lector.Disable()
                
                if ($disableResult.ReturnValue -eq 0) {
                    Write-Log "Deshabilitado correctamente."
                } else {
                    Write-Log "Advertencia al deshabilitar (código: $($disableResult.ReturnValue))"
                }
                
                Start-Sleep -Seconds 2
            }
            
            # Re-obtener el dispositivo
            $lectorActualizado = Get-WmiObject -Class Win32_PnPEntity -ComputerName $ComputerName -ErrorAction SilentlyContinue | Where-Object {
                $_.DeviceID -eq $deviceID
            }
            
            if ($lectorActualizado) {
                # Habilitar
                Write-Log "Habilitando..."
                $enableResult = $lectorActualizado.Enable()
                
                if ($enableResult.ReturnValue -eq 0) {
                    Write-Log "Habilitado correctamente."
                    $exitosos++
                } else {
                    Write-Log "Advertencia al habilitar (código: $($enableResult.ReturnValue))"
                    $errores++
                }
            } else {
                Write-Log "Advertencia: No se pudo re-obtener el dispositivo."
                $errores++
            }
            
        } catch {
            Write-Log "Error procesando dispositivo: $($_.Exception.Message)"
            $errores++
        }
    }
    
    Write-Log "----------------------------------------"
    Write-Log ""
    
    # Ejecutar escaneo de hardware final para asegurar que todo está OK
    Write-Log "Ejecutando escaneo de hardware final..."
    try {
        Invoke-WmiMethod -Path "root\cimv2:Win32_PnPEntity.DeviceID='HTREE\\ROOT\\0'" -Name ScanForHardwareChanges -ComputerName $ComputerName -ErrorAction SilentlyContinue | Out-Null
        Start-Sleep -Seconds 3
        Write-Log "Escaneo completado."
    } catch {
        Write-Log "Advertencia en escaneo final: $($_.Exception.Message)"
    }
    
    Write-Log ""
    Write-Log "RESUMEN: $exitosos dispositivo(s) reconectado(s) exitosamente, $errores error(es)."
    Write-Log "Proceso de reconexión completado en '$ComputerName'."
    Write-Host ""
    
} catch {
    Write-Log "Error durante la reconexión: $($_.Exception.Message)"
    Write-Log "Continuando con la verificación del servicio..."
}

# Verificar e iniciar el servicio de SmartCard si está detenido
try {
    $service = Get-WmiObject -Class Win32_Service -ComputerName $ComputerName | Where-Object { $_.Name -eq "SCardSvr" }
    if ($service) {
        Write-Log "Estado del servicio SCardSvr: $($service.State)"
        if ($service.State -eq "Stopped") {
            Write-Log "Iniciando servicio SCardSvr..."
            $startResult = $service.StartService()
            if ($startResult.ReturnValue -eq 0) {
                Write-Log "Servicio SCardSvr iniciado correctamente."
            } else {
                Write-Log "Error al iniciar SCardSvr. Código: $($startResult.ReturnValue)"
            }
        } elseif ($service.State -eq "Running") {
            Write-Log "Servicio SCardSvr ya está ejecutándose."
        } else {
            Write-Log "Servicio SCardSvr en estado: $($service.State). No se intenta iniciar."
        }
    } else {
        Write-Log "Advertencia: Servicio SCardSvr no encontrado."
    }
} catch {
    Write-Log "Error al verificar/iniciar el servicio SCardSvr: $($_.Exception.Message)"
}

# Pausa final
Write-Host "`n" -NoNewline
Write-Host "Proceso completado. Presione cualquier tecla para salir..." -ForegroundColor Green
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")