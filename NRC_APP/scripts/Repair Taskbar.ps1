#Requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'

$logDir = "C:\Temp"
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }

$log = Join-Path $logDir "AppxRepair_$(Get-Date -Format yyyyMMdd_HHmmss).log"
Start-Transcript -Path $log -Force | Out-Null

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','SKIP','STEP')]
        [string]$Level = 'INFO'
    )
    $ts  = Get-Date -Format 'HH:mm:ss'
    $pad = switch ($Level) {
        'INFO'  { '[INFO ]' }
        'OK'    { '[ OK  ]' }
        'WARN'  { '[WARN ]' }
        'ERROR' { '[ERROR]' }
        'SKIP'  { '[SKIP ]' }
        'STEP'  { '[STEP ]' }
    }
    Write-Host "$pad $ts  $Message"
}

Write-Log "================================================================" INFO
Write-Log "  AppX / Shell Repair  |  Equipo: $env:COMPUTERNAME" INFO
Write-Log "  Contexto: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" INFO
Write-Log "================================================================" INFO

try {
    # Obtener el SessionId del explorer activo
    Write-Log "----------------------------------------------------------------" INFO
    Write-Log "Identificando sesion de usuario activa..." STEP

    $activeUserFull  = (Get-WmiObject -Class Win32_ComputerSystem).UserName
    $activeSessionId = (Get-Process -Name explorer -ErrorAction SilentlyContinue |
        Where-Object { $_.SessionId -gt 0 } | Select-Object -First 1).SessionId

    $userSID          = $null
    $userLocalAppData = $null

    if ($activeUserFull) {
        try {
            $userSID     = ([System.Security.Principal.NTAccount]$activeUserFull).Translate(
                               [System.Security.Principal.SecurityIdentifier]).Value
            $userProfile = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$userSID" -ErrorAction Stop).ProfileImagePath
            $userLocalAppData = "$userProfile\AppData\Local"
            if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
                New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
            }
            Write-Log "Usuario : $activeUserFull" OK
            Write-Log "SID     : $userSID" OK
            Write-Log "Perfil  : $userProfile" OK
            Write-Log "Sesion  : $activeSessionId" OK
        } catch {
            Write-Log "No se pudo resolver el perfil del usuario: $($_.Exception.Message)" WARN
        }
    } else {
        Write-Log "No hay usuario interactivo detectado. Las operaciones de perfil se omitiran." WARN
    }

    # Cerrar explorer y los hosts UWP del shell antes de re-registrar paquetes.
    # Si alguno de estos procesos sigue activo, AppX devuelve 0x80073D02 (recurso en uso).
    Write-Log "----------------------------------------------------------------" INFO
    Write-Log "Cerrando shell y procesos UWP asociados..." STEP

    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    $uwpHosts = @('StartMenuExperienceHost','ShellExperienceHost','CortanaUI','SearchUI','SearchHost')
    foreach ($p in $uwpHosts) {
        Stop-Process -Name $p -Force -ErrorAction SilentlyContinue
    }

    # Espera 12 segundos a que los procesos terminen antes de continuar
    for ($i = 0; $i -lt 12; $i++) {
        if (-not (Get-Process -Name $uwpHosts -ErrorAction SilentlyContinue)) { break }
        Start-Sleep -Seconds 1
    }
    $stillRunning = $uwpHosts | Where-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue }
    if ($stillRunning) {
        Write-Log "Procesos UWP aun activos tras la espera: $($stillRunning -join ', ')" WARN
    } else {
        Write-Log "Shell cerrado correctamente." OK
    }

    # Reinicio WSearch
    Write-Log "----------------------------------------------------------------" INFO
    Write-Log "Reparando servicio WSearch e indice de busqueda..." STEP

    $wSearchSvc = Get-Service -Name WSearch -ErrorAction SilentlyContinue
    if ($wSearchSvc) {
        Stop-Service -Name WSearch -Force -ErrorAction SilentlyContinue
        Write-Log "Servicio WSearch detenido." OK

        $searchData = "C:\ProgramData\Microsoft\Search\Data\Applications\Windows"
        if (Test-Path $searchData) {
            Remove-Item "$searchData\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Indice de busqueda eliminado: $searchData" OK
        } else {
            Write-Log "Directorio de indice no encontrado, se omite: $searchData" INFO
        }

        Start-Service -Name WSearch -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        $svcStatus = (Get-Service -Name WSearch -ErrorAction SilentlyContinue).Status
        Write-Log "WSearch arrancado. Estado actual: $svcStatus" OK
    } else {
        Write-Log "El servicio WSearch no existe en este equipo." WARN
    }

    # Reset del paquete Settings
    Write-Log "----------------------------------------------------------------" INFO
    Write-Log "Reiniciando paquete Settings..." STEP
    Get-AppxPackage -Name "windows.immersivecontrolpanel" -ErrorAction SilentlyContinue | Reset-AppxPackage -ErrorAction SilentlyContinue
    Write-Log "Reset de Settings completado." OK

    # Re-registro paquetes
    Write-Log "----------------------------------------------------------------" INFO
    Write-Log "Re-registrando paquetes criticos del sistema (CBS, Core)..." STEP

    $criticalManifests = @(
        'C:\Windows\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\AppXManifest.xml',
        'C:\Windows\SystemApps\MicrosoftWindows.Client.Core_cw5n1h2txyewy\AppXManifest.xml'
    )
    foreach ($m in $criticalManifests) {
        $pkgLabel = Split-Path $m -Parent | Split-Path -Leaf
        if (-not (Test-Path $m)) {
            Write-Log "Manifiesto no encontrado, se omite: $pkgLabel" SKIP
            continue
        }
        $registered        = $false
        $skippedDueToInUse = $false
        for ($attempt = 1; $attempt -le 2 -and -not $registered; $attempt++) {
            try {
                Add-AppxPackage -DisableDevelopmentMode -Register $m -ErrorAction Stop
                $registered = $true
            } catch {
                $err = $_.Exception.Message
                if ($err -match '0x80073D02') {
                    Write-Log "En uso (0x80073D02), registro omitido: $pkgLabel" SKIP
                    $skippedDueToInUse = $true
                    break
                }
                Write-Log "Intento $attempt fallido para ${pkgLabel}: $err" WARN
                Start-Sleep -Seconds 2
                foreach ($p in $uwpHosts) { Stop-Process -Name $p -Force -ErrorAction SilentlyContinue }
            }
        }
        if ($registered) {
            Write-Log "Registrado OK: $pkgLabel" OK
        } elseif (-not $skippedDueToInUse) {
            Write-Log "No se pudo registrar tras $($attempt - 1) intento(s): $pkgLabel" ERROR
        }
    }

    Write-Log "----------------------------------------------------------------" INFO
    Write-Log "Re-registrando paquetes de shell de usuario..." STEP

    $targets = @(
        "Microsoft.Windows.ShellExperienceHost",
        "Microsoft.Windows.StartMenuExperienceHost",
        "windows.immersivecontrolpanel"
    )
    foreach ($name in $targets) {
        $pkg      = Get-AppxPackage -Name $name -ErrorAction SilentlyContinue
        $manifest = if ($pkg) { "$($pkg.InstallLocation)\AppXManifest.xml" } else { $null }

        if (-not $pkg) {
            Write-Log "Paquete no instalado, se omite: $name" SKIP
            continue
        }
        if (-not (Test-Path $manifest)) {
            Write-Log "Manifiesto inaccesible para ${name}: $manifest" WARN
            continue
        }

        $registered        = $false
        $skippedDueToInUse = $false
        for ($attempt = 1; $attempt -le 2 -and -not $registered; $attempt++) {
            try {
                Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction Stop
                $registered = $true
            } catch {
                $err = $_.Exception.Message
                if ($err -match '0x80073D02') {
                    Write-Log "En uso (0x80073D02), registro omitido: $name" SKIP
                    $skippedDueToInUse = $true
                    break
                }
                Write-Log "Intento $attempt fallido para ${name}: $err" WARN
                Start-Sleep -Seconds 2
                foreach ($p in $uwpHosts) { Stop-Process -Name $p -Force -ErrorAction SilentlyContinue }
            }
        }
        if ($registered) {
            Write-Log "Registrado OK: $name" OK
        } elseif (-not $skippedDueToInUse) {
            Write-Log "No se pudo registrar tras $($attempt - 1) intento(s): $name" ERROR
        }
    }

    # Regenerar Shell\
    Write-Log "----------------------------------------------------------------" INFO
    Write-Log "Limpiando cache de layout del shell..." STEP

    $shellPath = if ($userLocalAppData) { "$userLocalAppData\Microsoft\Windows\Shell" } else { $null }
    if (-not $shellPath) {
        Write-Log "Perfil de usuario no disponible, se omite la limpieza de cache." WARN
    } elseif (-not (Test-Path $shellPath)) {
        Write-Log "El directorio de cache no existe, nada que limpiar: $shellPath" INFO
    } else {
        Remove-Item $shellPath -Recurse -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
        if (-not (Test-Path $shellPath)) {
            Write-Log "Cache de shell eliminada: $shellPath" OK
        } else {
            Write-Log "Remove-Item no tuvo efecto, intentando forzar permisos con takeown/icacls..." WARN
            try {
                cmd /c "takeown /f `"$shellPath`" /r /d Y" | Out-Null
                cmd /c "icacls `"$shellPath`" /grant `"$activeUserFull`":(OI)(CI)F /t" | Out-Null
            } catch {
                Write-Log "Error al forzar permisos: $($_.Exception.Message)" ERROR
            }
            Start-Sleep -Milliseconds 200
            Remove-Item $shellPath -Recurse -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 300
            if (-not (Test-Path $shellPath)) {
                Write-Log "Cache de shell eliminada tras forzar permisos: $shellPath" OK
            } else {
                Write-Log "No se pudo eliminar la cache de shell: $shellPath" ERROR
            }
        }
    }

    # Limpar IrisService
    Write-Log "----------------------------------------------------------------" INFO
    Write-Log "Limpiando clave IrisService en el registro del usuario..." STEP

    if ($userSID) {
        $irisKey = "HKU:\$userSID\SOFTWARE\Microsoft\Windows\CurrentVersion\IrisService"
        if (Test-Path $irisKey) {
            Remove-Item -Path $irisKey -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Clave IrisService eliminada para SID: $userSID" OK
        } else {
            Write-Log "La clave IrisService no existe, no hay nada que limpiar." INFO
        }
    } else {
        Write-Log "SID de usuario no resuelto, se omite la limpieza de IrisService." WARN
    }
}
finally {
    Write-Log "================================================================" INFO
    Write-Log "Fase de finalizacion: relanzando shell de usuario..." STEP
    Start-Sleep -Seconds 5

    # Lanzar explorer.exe
    if ($activeSessionId) {
        try {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class SessionLauncher {
    [DllImport("wtsapi32.dll", SetLastError=true)]
    public static extern bool WTSQueryUserToken(uint sessionId, out IntPtr phToken);

    [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Auto)]
    public static extern bool CreateProcessAsUser(
        IntPtr hToken, string lpApplicationName, string lpCommandLine,
        IntPtr lpProcessAttributes, IntPtr lpThreadAttributes, bool bInheritHandles,
        uint dwCreationFlags, IntPtr lpEnvironment, string lpCurrentDirectory,
        ref STARTUPINFO lpStartupInfo, out PROCESS_INFORMATION lpProcessInformation);

    [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr h);

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
    public struct STARTUPINFO {
        public int cb, r1; public string lpDesktop, lpTitle;
        public uint dw1,dw2,dw3,dw4,dw5,dw6,dw7,dw8;
        public ushort w1,w2; public IntPtr p1,h1,h2,h3;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION {
        public IntPtr hProcess, hThread; public uint dwProcessId, dwThreadId;
    }
    public static bool LaunchInSession(uint sessionId, string cmd) {
        IntPtr token;
        if (!WTSQueryUserToken(sessionId, out token)) return false;
        try {
            var si = new STARTUPINFO();
            si.cb = System.Runtime.InteropServices.Marshal.SizeOf(si);
            si.lpDesktop = "winsta0\\\\default";
            PROCESS_INFORMATION pi;
            return CreateProcessAsUser(token, null, cmd,
                IntPtr.Zero, IntPtr.Zero, false, 0, IntPtr.Zero, null, ref si, out pi);
        } finally { CloseHandle(token); }
    }
}
"@ -ErrorAction Stop
            $ok = [SessionLauncher]::LaunchInSession([uint32]$activeSessionId, "explorer.exe")
            if ($ok) {
                Write-Log "Explorer relanzado en sesion $activeSessionId (usuario: $activeUserFull)." OK
            } else {
                $win32err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                Write-Log "CreateProcessAsUser fallo. Codigo Win32: $win32err" ERROR
            }
        } catch {
            Write-Log "Excepcion al relanzar explorer: $($_.Exception.Message)" ERROR
        }
    } else {
        Write-Log "SessionId no disponible, explorer no relanzado." WARN
    }

    Write-Log "================================================================" INFO
    Write-Log "Script finalizado. Log guardado en: $log" INFO
    Stop-Transcript | Out-Null
}