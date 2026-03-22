param(
    [Parameter(Mandatory=$true)]
    [string]$ComputerName,
    
    [Parameter(Mandatory=$true)]
    [string]$KB
)

# Función para obtener el SID del usuario activo (versión original sin filtros)
function Get-ActiveUserSID {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    # Usamos WMI en vez de CIM para evitar problemas de WinRM
    $profiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $ComputerName -ErrorAction Stop

    # Filtramos únicamente los que Win32_UserProfile indica como 'Loaded = True'
    $loadedProfiles = $profiles | Where-Object {
        $_.Loaded -eq $true
    }

    if (-not $loadedProfiles) {
        return $null
    }

    # Tomar el primero que cumpla
    return $loadedProfiles[0].SID
}

# Función para interpretar códigos de error de wusa
function Get-WusaErrorMessage {
    param([int]$ExitCode)
    
    # Convertir a hexadecimal
    $hexCode = '0x{0:X8}' -f $ExitCode
    
    switch ($ExitCode) {
        0 { 
            return @{
                Success = $true
                Message = "La desinstalación se completó exitosamente."
                Color = "Green"
                Icon = "✓"
            }
        }
        3010 { 
            return @{
                Success = $true
                Message = "La desinstalación se completó exitosamente. Se requiere reiniciar el equipo."
                Color = "Yellow"
                Icon = "✓"
            }
        }
        2359303 { # 0x00240017
            return @{
                Success = $false
                Message = "La actualización KB$KB no está instalada en el equipo."
                Color = "Red"
                Icon = "✗"
            }
        }
        -2145124329 { # 0x80240017
            return @{
                Success = $false
                Message = "La actualización KB$KB no está instalada en el equipo."
                Color = "Red"
                Icon = "✗"
            }
        }
        2359302 { # 0x00240016
            return @{
                Success = $false
                Message = "Las dependencias necesarias no están instaladas o la actualización está protegida."
                Color = "Red"
                Icon = "✗"
            }
        }
        -2145124330 { # 0x80240016
            return @{
                Success = $false
                Message = "Las dependencias necesarias no están instaladas o la actualización está protegida."
                Color = "Red"
                Icon = "✗"
            }
        }
        -2145124312 { # 0x80240028
            return @{
                Success = $false
                Message = "Otro proceso de instalación/desinstalación está en ejecución. Espere e intente nuevamente."
                Color = "Red"
                Icon = "✗"
            }
        }
        -2145124318 { # 0x80240022
            return @{
                Success = $false
                Message = "El servicio de Windows Update no está ejecutándose o no responde."
                Color = "Red"
                Icon = "✗"
            }
        }
        -2147024891 { # 0x80070005 (Access Denied)
            return @{
                Success = $false
                Message = "Acceso denegado. Verifique permisos de administrador o el usuario canceló."
                Color = "Red"
                Icon = "✗"
            }
        }
        5 { # ERROR_ACCESS_DENIED
            return @{
                Success = $false
                Message = "Acceso denegado. El usuario no confirmó la desinstalación o permisos insuficientes."
                Color = "Yellow"
                Icon = "⚠"
            }
        }
        -2147024894 { # 0x80070002
            return @{
                Success = $false
                Message = "La actualización KB$KB no se encuentra en el sistema o los archivos están dañados."
                Color = "Red"
                Icon = "✗"
            }
        }
        -2147467259 { # 0x80004005 (Unspecified error)
            return @{
                Success = $false
                Message = "Error general no especificado. Puede ser un problema de red o configuración."
                Color = "Red"
                Icon = "✗"
            }
        }
        1223 { # ERROR_CANCELLED
            return @{
                Success = $false
                Message = "El usuario canceló la operación de desinstalación."
                Color = "Yellow"
                Icon = "⚠"
            }
        }
        1602 { # ERROR_INSTALL_USEREXIT
            return @{
                Success = $false
                Message = "El usuario canceló la instalación/desinstalación."
                Color = "Yellow"
                Icon = "⚠"
            }
        }
        1618 { # ERROR_INSTALL_ALREADY_RUNNING
            return @{
                Success = $false
                Message = "Ya hay otra instalación en curso. Espere a que finalice."
                Color = "Red"
                Icon = "✗"
            }
        }
        default { 
            return @{
                Success = $false
                Message = "Error desconocido ($hexCode). Consulte los logs del sistema para más detalles."
                Color = "Red"
                Icon = "✗"
            }
        }
    }
}

# Verificar conectividad
if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
    Write-Host "ERROR: No se pudo alcanzar el equipo remoto '$ComputerName'." -ForegroundColor Red
    Read-Host "Presione Enter para cerrar"
    exit 1
}

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "Desinstalación de KB$KB en equipo: $ComputerName" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# Ruta a PsExec (debe estar en la carpeta de la aplicación o en PATH)
$psexecPath = Join-Path $PSScriptRoot "..\tools\PsExec.exe"

if (-not (Test-Path $psexecPath)) {
    Write-Host "ERROR: No se encontró PsExec.exe en: $psexecPath" -ForegroundColor Red
    Write-Host "Descargue PsExec de Sysinternals y colóquelo en la carpeta 'tools'" -ForegroundColor Yellow
    Read-Host "Presione Enter para cerrar"
    exit 1
}

try {
    Write-Host "Detectando usuario activo en '$ComputerName'..." -ForegroundColor Yellow
    
    # Obtener el SID del usuario con sesión cargada
    $sidActivo = Get-ActiveUserSID -ComputerName $ComputerName
    
    if (-not $sidActivo) {
        Write-Host ""
        Write-Host "No se encontró ningún perfil cargado en '$ComputerName'." -ForegroundColor Red
        Write-Host "Asegúrese de que hay un usuario conectado al equipo." -ForegroundColor Yellow
        Read-Host "Presione Enter para cerrar"
        exit 1
    }
    
    Write-Host "Usuario activo detectado (SID: $sidActivo)" -ForegroundColor Green
    
    # Obtener información del usuario desde el SID para mostrar el nombre
    try {
        $userProfiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $ComputerName | Where-Object { $_.SID -eq $sidActivo }
        if ($userProfiles) {
            $localPath = $userProfiles.LocalPath
            $userName = Split-Path $localPath -Leaf
            Write-Host "Usuario: $userName" -ForegroundColor Cyan
        }
    } catch {
        # No es crítico, continuamos
    }
    
    Write-Host ""
    Write-Host "Ejecutando desinstalación en la sesión del usuario..." -ForegroundColor Yellow
    Write-Host "La ventana de desinstalación aparecerá en el escritorio del usuario." -ForegroundColor Yellow
    Write-Host ""
    
    # Ejecutar wusa en la sesión interactiva del usuario
    $arguments = "\\$ComputerName -accepteula -s -i wusa.exe /uninstall /kb:$KB /norestart"
    
    $process = Start-Process -FilePath $psexecPath -ArgumentList $arguments -NoNewWindow -Wait -PassThru
    
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "RESULTADO DE LA DESINSTALACIÓN" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Interpretar el código de salida
    $errorInfo = Get-WusaErrorMessage -ExitCode $process.ExitCode
    
    Write-Host "$($errorInfo.Icon) $($errorInfo.Message)" -ForegroundColor $errorInfo.Color
    Write-Host ""
    $hexCodeDisplay = '0x{0:X8}' -f $process.ExitCode
    Write-Host "Código de salida: $($process.ExitCode) ($hexCodeDisplay)" -ForegroundColor Gray
    
} catch {
    Write-Host ""
    Write-Host "ERROR CRÍTICO: No se pudo ejecutar el comando en el equipo remoto." -ForegroundColor Red
    Write-Host "Detalles: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Read-Host "Presione Enter para cerrar"