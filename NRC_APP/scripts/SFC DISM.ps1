# Asegura C:\Logs
$LogPath = "C:\Logs"
if (!(Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory | Out-Null
}

$worker = Join-Path $env:TEMP 'PROBA_background_worker.ps1'
$workerScript = @'
# Worker: ejecuta DISM luego SFC; registra en C:\Logs
$LogPath = "C:\Logs"
if (!(Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory | Out-Null }

$dismPath  = Join-Path $env:SystemRoot 'System32\dism.exe'
$wrapLog   = Join-Path $LogPath 'Worker.log'

# ── DISM RestoreHealth ──────────────────────────────────────────────────────
$dismNativeLog = Join-Path $LogPath 'DISM_native.log'
Add-Content -Path $wrapLog -Value "=== DISM RestoreHealth started: $(Get-Date -Format o)"
try {
    $p = Start-Process -FilePath $dismPath `
        -ArgumentList '/Online','/Cleanup-Image','/RestoreHealth',"/LogPath:$dismNativeLog" `
        -WindowStyle Hidden -PassThru

    $lastLen = if (Test-Path $dismNativeLog) { (Get-Item $dismNativeLog).Length } else { 0 }
    while (-not $p.HasExited) {
        Start-Sleep -Seconds 15
        $len = if (Test-Path $dismNativeLog) { (Get-Item $dismNativeLog).Length } else { $lastLen }
        if ($len -gt $lastLen) {
            Add-Content -Path $wrapLog -Value "DISM RestoreHealth sigue trabajando... log crece (+$($len - $lastLen) bytes) $(Get-Date -Format o)"
            $lastLen = $len
        } else {
            Add-Content -Path $wrapLog -Value "DISM RestoreHealth sigue en ejecucion (sin cambios en log) $(Get-Date -Format o)"
        }
    }
    Add-Content -Path $wrapLog -Value "DISM RestoreHealth exit code: $($p.ExitCode)"
} catch {
    Add-Content -Path $wrapLog -Value "DISM RestoreHealth exception: $($_.Exception.Message)"
}
Add-Content -Path $wrapLog -Value "=== DISM RestoreHealth finished: $(Get-Date -Format o)`n"

# ── DISM StartComponentCleanup ──────────────────────────────────────────────
$dismSccNativeLog = Join-Path $LogPath 'DISM_SCC_native.log'
Add-Content -Path $wrapLog -Value "=== DISM StartComponentCleanup started: $(Get-Date -Format o)"
try {
    $pC = Start-Process -FilePath $dismPath `
        -ArgumentList '/Online','/Cleanup-Image','/StartComponentCleanup',"/LogPath:$dismSccNativeLog" `
        -WindowStyle Hidden -PassThru

    $lastLen = if (Test-Path $dismSccNativeLog) { (Get-Item $dismSccNativeLog).Length } else { 0 }
    while (-not $pC.HasExited) {
        Start-Sleep -Seconds 15
        $len = if (Test-Path $dismSccNativeLog) { (Get-Item $dismSccNativeLog).Length } else { $lastLen }
        if ($len -gt $lastLen) {
            Add-Content -Path $wrapLog -Value "DISM StartComponentCleanup sigue trabajando... log crece (+$($len - $lastLen) bytes) $(Get-Date -Format o)"
            $lastLen = $len
        } else {
            Add-Content -Path $wrapLog -Value "DISM StartComponentCleanup sigue en ejecucion (sin cambios en log) $(Get-Date -Format o)"
        }
    }
    Add-Content -Path $wrapLog -Value "DISM StartComponentCleanup exit code: $($pC.ExitCode)"
} catch {
    Add-Content -Path $wrapLog -Value "DISM StartComponentCleanup exception: $($_.Exception.Message)"
}
Add-Content -Path $wrapLog -Value "=== DISM StartComponentCleanup finished: $(Get-Date -Format o)`n"

# ── SFC ─────────────────────────────────────────────────────────────────────
$sfcNativeLog = Join-Path $LogPath 'SFC_native.log'
Add-Content -Path $wrapLog -Value "=== SFC started: $(Get-Date -Format o)"
try {
    $p2 = Start-Process -FilePath 'cmd.exe' `
        -ArgumentList '/c', "sfc /scannow >> `"$sfcNativeLog`" 2>&1" `
        -WindowStyle Hidden -Wait -PassThru
    Add-Content -Path $wrapLog -Value "SFC exit code: $($p2.ExitCode)"
} catch {
    Add-Content -Path $wrapLog -Value "SFC exception: $($_.Exception.Message)"
}
Add-Content -Path $wrapLog -Value "=== SFC finished: $(Get-Date -Format o)`n"
'@

Set-Content -Path $worker -Value $workerScript -Encoding UTF8

# Inicia worker
$psExe = Join-Path $PSHOME 'powershell.exe'
Start-Process -FilePath $psExe -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-WindowStyle','Hidden','-File',$worker -WindowStyle Hidden -WorkingDirectory $env:TEMP
Write-Host "Launcher iniciado; saliendo..."