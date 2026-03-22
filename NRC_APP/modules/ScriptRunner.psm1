# ScriptRunner concentra la forma de lanzar scripts desde NRC_APP.
# Aquí se decide si un fichero se ejecuta en local, se lanza como batch o
# se copia al remoto para ejecutarlo con PsExec y contexto SYSTEM.

# Ejecuta un .ps1 local y, si corresponde, le pasa -ComputerName.
function Invoke-NRCPowerShellScript {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptPath,

        [string]$ComputerName = "",

        # Mantiene la ventana abierta al terminar para revisar la salida.
        [switch]$NoExit,

        # Lanza PowerShell sin mostrar la ventana.
        [switch]$Hidden
    )

    if (-not (Test-Path $ScriptPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No se encuentra el script:`n$ScriptPath",
            "Error - Script no encontrado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $flags = if ($NoExit) { "-NoExit" } else { "" }
    $style = if ($Hidden)  { "-WindowStyle Hidden" } else { "" }

    $argList = "$flags $style -ExecutionPolicy Bypass -NoProfile -File `"$ScriptPath`""

    if (-not [string]::IsNullOrWhiteSpace($ComputerName)) {
        $argList += " -ComputerName `"$ComputerName`""
    }

    Start-Process -FilePath "powershell.exe" -ArgumentList $argList
}

# Ejecuta un .bat/.cmd y deja la ventana abierta para revisar la salida.
function Invoke-NRCBatchScript {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptPath,

        [string]$ComputerName = ""
    )

    if (-not (Test-Path $ScriptPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No se encuentra el script:`n$ScriptPath",
            "Error - Script no encontrado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    # /k mantiene la ventana abierta para que el técnico pueda revisar el resultado.
    $cmdArgs = "/k call `"$ScriptPath`""
    if (-not [string]::IsNullOrWhiteSpace($ComputerName)) {
        $cmdArgs += " `"$ComputerName`""
    }

    Start-Process -FilePath "cmd.exe" -ArgumentList $cmdArgs
}

#------------------------------------------------------------------
# FUNCIÓN: Invoke-NRCScript
# Punto de entrada único: detecta la extensión y llama al launcher
# adecuado. Todos los scripts nuevos que se añadan "en caliente"
# pueden usarlo directamente.
#------------------------------------------------------------------
function Invoke-NRCScript {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptPath,

        [string]$ComputerName = "",

        [switch]$NoExit,
        [switch]$Hidden
    )

    $ext = [System.IO.Path]::GetExtension($ScriptPath).ToLower()

    switch ($ext) {
        '.ps1' {
            Invoke-NRCPowerShellScript -ScriptPath $ScriptPath `
                -ComputerName $ComputerName `
                -NoExit:$NoExit `
                -Hidden:$Hidden
        }
        '.bat' {
            Invoke-NRCBatchScript -ScriptPath $ScriptPath -ComputerName $ComputerName
        }
        '.vbs' {
            $vbsArgs = "`"$ScriptPath`""
            if (-not [string]::IsNullOrWhiteSpace($ComputerName)) {
                $vbsArgs += " `"$ComputerName`""
            }
            Start-Process -FilePath "wscript.exe" -ArgumentList $vbsArgs
        }
        default {
            # Intento genérico (ejecutables, etc.)
            if (-not [string]::IsNullOrWhiteSpace($ComputerName)) {
                Start-Process $ScriptPath -ArgumentList $ComputerName
            } else {
                Start-Process $ScriptPath
            }
        }
    }
}

#------------------------------------------------------------------
# FUNCIÓN: Test-NRCRemoteConnectivity
# Comprueba si el equipo responde a ping Y si el share C$ es
# accesible (necesario para copiar scripts vía PsExec).
# Devuelve $true/$false. Muestra MessageBox si falla.
#------------------------------------------------------------------
function Test-NRCRemoteConnectivity {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,

        [switch]$Silent   # Si se activa, no muestra MessageBox en caso de fallo
    )

    # 1) Ping
    try {
        if (-not (Test-Connection -ComputerName $ComputerName -Count 2 -Quiet)) {
            if (-not $Silent) {
                [System.Windows.Forms.MessageBox]::Show(
                    "El equipo '$ComputerName' no responde al ping.`nCompruebe que está encendido y accesible.",
                    "Sin Conectividad",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            }
            return $false
        }
    } catch {
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show(
                "Error al verificar conectividad con '$ComputerName':`n$_",
                "Error de Conexión",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
        return $false
    }

    return $true
}

#------------------------------------------------------------------
# FUNCIÓN: Invoke-NRCPsExecScript
# Copia un script .ps1 al equipo remoto vía admin share y lo
# ejecuta como SYSTEM usando PsExec. Ideal para scripts que
# necesitan #Requires -RunAsAdministrator o acceso local.
#------------------------------------------------------------------
function Invoke-NRCPsExecScript {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptPath,

        [Parameter(Mandatory=$true)]
        [string]$ComputerName,

        [string]$DisplayName = ''
    )

    $psexecExe = Join-Path $Global:ScriptRoot 'tools\PsExec.exe'
    if (-not (Test-Path $psexecExe)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No se encuentra PsExec.exe en tools\PsExec.exe",
            "PsExec no encontrado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    if (-not (Test-Path $ScriptPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No se encuentra el script:`n$ScriptPath",
            "Script no encontrado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $fileName      = [System.IO.Path]::GetFileName($ScriptPath)
    $safeName      = $fileName -replace '\s', '_'
    $remoteUncPath = '\\' + $ComputerName + '\C$\Windows\Temp\' + $safeName

    # Copiar script al equipo remoto
    try {
        Copy-Item -LiteralPath $ScriptPath -Destination $remoteUncPath -Force -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "No se pudo copiar el script a '$ComputerName':`n$($_.Exception.Message)",
            "Error al copiar",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $label          = if (-not [string]::IsNullOrWhiteSpace($DisplayName)) { $DisplayName } else { $fileName }
    $safeExe        = $psexecExe.Replace("'", "''")
    $remoteExecPath = 'C:\Windows\Temp\' + $safeName

    $invokeCmd = "Write-Host 'Script: $label | Equipo: $ComputerName' -ForegroundColor Cyan; " +
                 "& '$safeExe' \\$ComputerName -accepteula -s powershell.exe " +
                 "-NoProfile -ExecutionPolicy Bypass -NonInteractive -File '$remoteExecPath'; " +
                 "`$code=`$LASTEXITCODE; " +
                 "if(`$code -eq 0){Write-Host 'Completado correctamente.' -ForegroundColor Green}" +
                 "else{Write-Host ('Finalizado con codigo: ' + `$code) -ForegroundColor Red}; " +
                 "Read-Host 'Presione Enter para cerrar'"

    Start-Process powershell.exe -ArgumentList ("-NoExit -ExecutionPolicy Bypass -NoProfile -Command `"$invokeCmd`"")
    Add-Logs -text "$ComputerName - Script '$label' ejecutado via PsExec"
}

Export-ModuleMember -Function *


