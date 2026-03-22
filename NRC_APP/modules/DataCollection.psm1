<#
    Recogida remota de datos para la ventana principal.
    La consulta se ejecuta en un Job independiente para no bloquear la UI.
    Los resultados salen por partes y el timer de WinForms los va aplicando
    sobre el panel de salida y los indicadores de estado.
#>

# Reemplaza una línea ya pintada en el RTB sin reconstruir todo el texto.
function global:Update-StreamRTBLine {
    param(
        [string]$Key,
        [string]$NewText,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::Black
    )
    try {
        $rtb = $null
        if ($global:StreamUI -and $global:StreamUI.RTBOutput) { $rtb = $global:StreamUI.RTBOutput }
        if ($null -eq $rtb -or $rtb.IsDisposed) { return }
        if ($null -eq $global:StreamLineIndices) { return }
        if (-not $global:StreamLineIndices.ContainsKey($Key)) { return }

        $lineIdx = $global:StreamLineIndices[$Key]
        if ($lineIdx -lt 0) { return }

        $charStart = $rtb.GetFirstCharIndexFromLine($lineIdx)
        if ($charStart -lt 0) { return }

        $linesArray = $rtb.Lines
        if ($null -eq $linesArray -or $lineIdx -ge $linesArray.Length) { return }

        $oldText = $linesArray[$lineIdx]
        if ($null -eq $oldText) { return }

        $global:StreamUpdating = $true
        $rtb.Select($charStart, $oldText.Length)
        $rtb.SelectedText = $NewText
        $rtb.Select($charStart, $NewText.Length)
        $rtb.SelectionColor = $Color
        $rtb.SelectionStart  = $rtb.TextLength
        $rtb.SelectionLength = 0
        $rtb.ScrollToCaret()
        $global:StreamUpdating = $false
    } catch {
        $global:StreamUpdating = $false
    }
}

# Detiene timer y Job, y deja el estado listo para la siguiente consulta.
function global:Cleanup-StreamingCollection {
    if ($global:StreamTimer) {
        try { $global:StreamTimer.Stop(); $global:StreamTimer.Dispose() } catch {}
        $global:StreamTimer = $null
    }
    if ($global:StreamJob) {
        try { Stop-Job   $global:StreamJob -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job $global:StreamJob -Force -ErrorAction SilentlyContinue } catch {}
        $global:StreamJob = $null
    }
    $global:StreamRunning     = $false
    $global:StreamLineIndices = $null
    $global:StreamUpdating    = $false
    $global:StreamStart       = $null
    try {
        if ($global:StreamUI -and $global:StreamUI.ButtonCheck) {
            $global:StreamUI.ButtonCheck.Enabled = $true
        }
    } catch {}
    $global:StreamUI = $null
}

# Vuelca al RTB y a los labels lo que vaya llegando desde el Job.
function global:Apply-CollectionResults {
    param([hashtable]$R)

    $UI = $global:StreamUI
    if (-not $UI) { return }

    # Permisos
    if ($R.NoPermission) {
        try { if ($UI.LblPermission) { $UI.LblPermission.Text = 'FAIL'; $UI.LblPermission.ForeColor = 'red' } } catch {}
        return
    }
    try { if ($UI.LblPermission) { $UI.LblPermission.Text = 'OK'; $UI.LblPermission.ForeColor = 'green' } } catch {}

    # Nombre de equipo (DNS)
    if ($R.Hostname) {
        Update-StreamRTBLine -Key 'dns' -NewText "Nome de Equipo: $($R.Hostname)" -Color ([System.Drawing.Color]::Black)
    }

    # Sistema operativo
    if ($R.OSCaption) {
        $arch = if ($R.OSArch -match '64') { 'x64' } else { 'x86' }
        $osLine = "$($R.OSCaption) $arch"
        if ($R.OSCSD) { $osLine += "  $($R.OSCSD)" }
        if ($R.OSVersion) { $osLine += "  Versión:$($R.OSVersion)" }
        if ($R.OSDisplayVersion) { $osLine += " ($($R.OSDisplayVersion))" }
        Update-StreamRTBLine -Key 'os' -NewText "SO: $osLine" -Color ([System.Drawing.Color]::Black)
        try { if ($UI.LblOS) { $UI.LblOS.Text = 'ON'; $UI.LblOS.ForeColor = 'green' } } catch {}
    } else {
        try { if ($UI.LblOS) { $UI.LblOS.Text = 'N/A'; $UI.LblOS.ForeColor = 'gray' } } catch {}
    }

    # Tiempo de encendido
    try {
        if ($R.LastBootUpTime -and $UI.LblUptime) {
            $up = New-TimeSpan -Start $R.LastBootUpTime -End (Get-Date)
            $UI.LblUptime.Text      = "$($up.Days) Dias $($up.Hours) Horas $($up.Minutes) Minutos $($up.Seconds) Segundos"
            $UI.LblUptime.ForeColor = 'blue'
        }
    } catch {}

    # Logon server del equipo
    if ($R.LogonServerEquipo) {
        Update-StreamRTBLine -Key 'logon_eq' -NewText "Logon Server Equipo: $($R.LogonServerEquipo)" -Color ([System.Drawing.Color]::Black)
    } else {
        Update-StreamRTBLine -Key 'logon_eq' -NewText "Logon Server Equipo: No disponible" -Color ([System.Drawing.Color]::Gray)
    }

    # Logon server del usuario
    if ($R.LogonServerUsuario) {
        Update-StreamRTBLine -Key 'logon_usr' -NewText "Logon Server Usuario: $($R.LogonServerUsuario)" -Color ([System.Drawing.Color]::Black)
    }

    # Red
    if ($R.IPAddress) {
        Update-StreamRTBLine -Key 'net' -NewText "IP/Gateway/MAC/DNS: $($R.IPAddress) / $($R.Gateway) / $($R.MACAddress) / $($R.DNSServers)" -Color ([System.Drawing.Color]::Black)
    }

    # Hardware y disco
    if ($R.Model) {
        Update-StreamRTBLine -Key 'hw' -NewText "Modelo Equipo/CPU: $($R.Model) / $($R.CPUName) / $($R.CPUCores) núcleos" -Color ([System.Drawing.Color]::Black)
    }
    if ($null -ne $R.RAM) {
        Update-StreamRTBLine -Key 'ram' -NewText "RAM: $($R.RAM) GB" -Color ([System.Drawing.Color]::Black)
    }
    if ($R.SN) {
        Update-StreamRTBLine -Key 'sn' -NewText "S/N: $($R.SN)" -Color ([System.Drawing.Color]::Black)
    }
    if ($null -ne $R.DiskFree) {
        Update-StreamRTBLine -Key 'disk' -NewText "Espazo Libre $($R.DiskFree) GB ($($R.DiskTotal) GB en total)" -Color ([System.Drawing.Color]::Black)
    }

    # Sesión de usuario activa
    if ($R.ContainsKey('InteractiveUser')) {
        $u = $R.InteractiveUser
        if ([string]::IsNullOrWhiteSpace($u)) {
            $ud = 'Non hai ningún usuario logueado no equipo'
        } else {
            $ud = if ($u -match '^(.+)\\(.+)$') { $Matches[2] } else { $u }
            if ($R.IsRDP) { $ud += ' (Escritorio Remoto)' }
        }
        Update-StreamRTBLine -Key 'user' -NewText "Sesión Usuario: $ud" -Color ([System.Drawing.Color]::Black)
    }

    # Extensión telefónica
    if ($R.Extension -and $R.Extension -ne 'Sin extensión definida') {
        Update-StreamRTBLine -Key 'ext' -NewText "Extensión asignada: $($R.Extension)" -Color ([System.Drawing.Color]::Black)
    }

    # Labels VNC/RDP
    try {
        if ($UI.LblVNC) {
            if ($R.VNCPort) { $UI.LblVNC.Text='OPEN'; $UI.LblVNC.ForeColor='green' }
            else            { $UI.LblVNC.Text='CLOSED'; $UI.LblVNC.ForeColor='red' }
        }
    } catch {}
    try {
        if ($UI.LblRDP) {
            if ($R.RDPPort) { $UI.LblRDP.Text='OPEN'; $UI.LblRDP.ForeColor='green' }
            else            { $UI.LblRDP.Text='CLOSED'; $UI.LblRDP.ForeColor='red' }
        }
    } catch {}

    # Label WinRM
    try {
        if ($UI.LblWinRM -and $R.WinRMStatus) {
            $wr = $R.WinRMStatus
            $st = switch ($wr.StartMode) {
                'Auto'      { 'Auto' }
                'Automatic' { 'Auto' }
                'Manual'    { 'Manual' }
                'Disabled'  { 'Dis.' }
                default     { $wr.StartMode }
            }
            $estado = if ($wr.State) { $wr.State } elseif ($wr.Status) { $wr.Status } else { '' }
            if ($estado -eq 'Running') { $UI.LblWinRM.Text = "ON ($st)";  $UI.LblWinRM.ForeColor = 'green' }
            else                       { $UI.LblWinRM.Text = "OFF ($st)"; $UI.LblWinRM.ForeColor = 'red'   }
        } elseif ($UI.LblWinRM) {
            $UI.LblWinRM.Text = 'N/A'; $UI.LblWinRM.ForeColor = 'gray'
        }
    } catch {}
}

# Inicia la recogida de datos en un Job separado y un timer de 500ms que lo monitoriza.
# Escribe placeholders en el RTB antes de lanzar el Job para poder actualizarlos en sitio.
function global:Start-StreamingDataCollection {
    param(
        [string]   $ComputerNameParam,
        [hashtable]$UI = @{}
    )

    if ($global:StreamRunning) {
        [System.Windows.Forms.MessageBox]::Show(
            "Ya hay una recogida de datos en ejecución. Por favor, espere.",
            "Proceso en ejecución",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }

    $global:StreamUI        = $UI
    $global:StreamRunning   = $true
    $global:StreamStart     = Get-Date
    $global:StreamUpdating  = $false
    $global:StreamJob       = $null

    try { if ($UI.ButtonCheck) { $UI.ButtonCheck.Enabled = $false } } catch {}

    try { Add-Logs -text "🔄 Iniciando recogida de datos para: $ComputerNameParam (timeout: 40s)" } catch {}

    # Reset labels
    foreach ($lbl in @($UI.LblPermission, $UI.LblVNC, $UI.LblRDP, $UI.LblWinRM, $UI.LblOS)) {
        try { if ($lbl) { $lbl.Text = '...'; $lbl.ForeColor = [System.Drawing.Color]::DarkGray } } catch {}
    }
    try { if ($UI.LblUptime) { $UI.LblUptime.Text = ''; $UI.LblUptime.ForeColor = [System.Drawing.Color]::DarkGray } } catch {}

    $rtb = if ($UI.RTBOutput) { $UI.RTBOutput } else { $null }
    if (-not $rtb) {
        try { Add-Logs -text "ERROR: RTBOutput es null" } catch {}
        Cleanup-StreamingCollection
        return $false
    }

    # Escribe los placeholders iniciales en el RTB antes de lanzar el job
    $global:StreamUpdating = $true
    $sep = '#' * 41
    $rtb.SelectionStart  = $rtb.TextLength
    $rtb.SelectionLength = 0
    $rtb.SelectionColor  = [System.Drawing.Color]::Black
    $rtb.AppendText("$sep$ComputerNameParam$sep`n")

    $placeholders = @(
        @{ Key='dns';       Label='Nome de Equipo: --' }
        @{ Key='os';        Label='SO: --' }
        @{ Key='logon_eq';  Label='Logon Server Equipo: --' }
        @{ Key='logon_usr'; Label='Logon Server Usuario: --' }
        @{ Key='net';       Label='IP/Gateway/MAC/DNS: --' }
        @{ Key='hw';        Label='Modelo Equipo/CPU: --' }
        @{ Key='ram';       Label='RAM: --' }
        @{ Key='sn';        Label='S/N: --' }
        @{ Key='disk';      Label='Espazo Libre --' }
        @{ Key='user';      Label='Sesión Usuario: --' }
        @{ Key='blank';     Label='' }
        @{ Key='ext';       Label='Extensión asignada: --' }
    )

    $global:StreamLineIndices = @{}
    $rtb.SelectionColor = [System.Drawing.Color]::DarkGray
    foreach ($ph in $placeholders) {
        $global:StreamLineIndices[$ph.Key] = $rtb.GetLineFromCharIndex($rtb.TextLength)
        $rtb.AppendText("$($ph.Label)`n")
    }
    $rtb.ScrollToCaret()
    $global:StreamUpdating = $false
    [System.Windows.Forms.Application]::DoEvents()

    # El Job emite hashtables parciales via Write-Output; Receive-Job devuelve todos los emitidos,
    # coge siempre el ultimo que es el mas completo.
    $global:StreamJob = Start-Job -ScriptBlock {
        param($cn, $scriptRoot)

        $R = @{
            ComputerName = $cn
            NoPermission = $false
        }

        # Verificar acceso administrativo al equipo
        try {
            if (-not (Test-Path "\\$cn\c$")) {
                $R.NoPermission = $true
                Write-Output $R
                return
            }
        } catch {
            $R.NoPermission = $true
            Write-Output $R
            return
        }

        # Comprobar puertos VNC (5700) y RDP (3389)
        try {
            $t = New-Object System.Net.Sockets.TcpClient
            $c = $t.BeginConnect($cn, 5700, $null, $null)
            $R.VNCPort = ($c.AsyncWaitHandle.WaitOne(1000, $false) -and $t.Connected)
            try { $t.Close() } catch {}; $t.Dispose()
        } catch { $R.VNCPort = $false }

        try {
            $t = New-Object System.Net.Sockets.TcpClient
            $c = $t.BeginConnect($cn, 3389, $null, $null)
            $R.RDPPort = ($c.AsyncWaitHandle.WaitOne(1000, $false) -and $t.Connected)
            try { $t.Close() } catch {}; $t.Dispose()
        } catch { $R.RDPPort = $false }

        # Abrir sesión CIM única (DCOM) para todas las consultas WMI
        $cimSession = $null
        try {
            $cimOpt     = New-CimSessionOption -Protocol Dcom
            $cimSession = New-CimSession -ComputerName $cn -SessionOption $cimOpt -OperationTimeoutSec 12 -ErrorAction Stop
        } catch {
            $R.CIMError = $_.Exception.Message
        }

        if ($cimSession) {
            # Win32_ComputerSystem: modelo, RAM, dominio y usuario activo
            try {
                $cs = Get-CimInstance -CimSession $cimSession -ClassName Win32_ComputerSystem -OperationTimeoutSec 8 -ErrorAction Stop
                $R.Model  = $cs.Model
                $R.RAM    = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
                $R.Domain = $cs.Domain

                # Usuario: primero Win32_ComputerSystem.UserName, fallback al dueño de explorer.exe
                $iu = $null; $isRDP = $false
                try {
                    if ($cs.UserName) { $iu = $cs.UserName }
                    if (-not $iu) {
                        $exp = Get-CimInstance -CimSession $cimSession -ClassName Win32_Process -Filter "Name='explorer.exe'" -OperationTimeoutSec 5 -ErrorAction SilentlyContinue | Select-Object -First 1
                        if ($exp) {
                            $own = Invoke-CimMethod -InputObject $exp -MethodName GetOwner -OperationTimeoutSec 5 -ErrorAction SilentlyContinue
                            if ($own) { $iu = if ($own.Domain) { "$($own.Domain)\$($own.User)" } else { $own.User } }
                        }
                    }
                    if ($iu) {
                        $rdpProc = Get-CimInstance -CimSession $cimSession -ClassName Win32_Process -Filter "Name='mstsc.exe'" -OperationTimeoutSec 5 -ErrorAction SilentlyContinue
                        $isRDP = ($null -ne $rdpProc)
                    }
                } catch {}
                $R.InteractiveUser = $iu
                $R.IsRDP           = $isRDP
            } catch {}

            # Emitir parcial — usuario, modelo y RAM ya disponibles
            Write-Output $R.Clone()

            # Servicio WinRM
            try {
                $wr = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service -Filter "Name='WinRM'" -OperationTimeoutSec 8 -ErrorAction SilentlyContinue
                if ($wr) { $R.WinRMStatus = @{ State = $wr.State; StartMode = $wr.StartMode } }
            } catch {}

            # Win32_OperatingSystem: versión, arquitectura, uptime
            try {
                $os = Get-CimInstance -CimSession $cimSession -ClassName Win32_OperatingSystem -OperationTimeoutSec 8 -ErrorAction Stop
                $R.OSCaption      = $os.Caption
                $R.OSVersion      = $os.Version
                $R.OSArch         = $os.OSArchitecture
                $R.OSCSD          = $os.CSDVersion
                $R.OSOther        = $os.OtherTypeDescription
                $R.LastBootUpTime = $os.LastBootUpTime
            } catch {}

            # Procesador
            try {
                $proc = Get-CimInstance -CimSession $cimSession -ClassName Win32_Processor -OperationTimeoutSec 8 -ErrorAction Stop | Select-Object -First 1
                $R.CPUName  = $proc.Name
                $R.CPUCores = $proc.NumberOfCores
            } catch {}

            # Número de serie (BIOS)
            try {
                $bios = Get-CimInstance -CimSession $cimSession -ClassName Win32_BIOS -OperationTimeoutSec 8 -ErrorAction Stop
                $R.SN = $bios.SerialNumber
            } catch {}

            # Espacio en discos locales (tipo 3 = fixed)
            try {
                $disks = Get-CimInstance -CimSession $cimSession -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -OperationTimeoutSec 8 -ErrorAction Stop
                $R.DiskTotal = [math]::Round(($disks | Measure-Object -Property Size      -Sum).Sum / 1GB, 2)
                $R.DiskFree  = [math]::Round(($disks | Measure-Object -Property FreeSpace -Sum).Sum / 1GB, 2)
            } catch {}

            # Configuración de red (primer adaptador con IP activa)
            try {
                $net = Get-CimInstance -CimSession $cimSession -ClassName Win32_NetworkAdapterConfiguration -OperationTimeoutSec 8 -ErrorAction SilentlyContinue |
                       Where-Object { $_.IPEnabled } | Select-Object -First 1
                if ($net) {
                    $R.IPAddress  = ($net.IPAddress -join ' ')
                    $R.Gateway    = ($net.DefaultIPGateway -join ' ')
                    $R.MACAddress = $net.MACAddress
                    $R.DNSServers = ($net.DNSServerSearchOrder -join ' ')
                }
            } catch {}

            try { Remove-CimSession $cimSession -ErrorAction SilentlyContinue } catch {}

            # Emitir parcial — todos los datos CIM ya recogidos
            Write-Output $R.Clone()
        }

        # Resolver nombre DNS del equipo
        try {
            $dns = Resolve-DnsName -Name $cn -ErrorAction Stop
            if ($dns.NameHost)    { $R.Hostname = $dns.NameHost | Select-Object -First 1 }
            elseif ($dns.Name)    { $R.Hostname = $dns.Name     | Select-Object -First 1 }
            else                  { $R.Hostname = $cn }
        } catch { $R.Hostname = $cn }

        # DisplayVersion del registro (p.ej. '22H2')
        try {
            $regResult = reg query "\\$cn\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v DisplayVersion 2>$null |
                         Select-String -Pattern 'DisplayVersion'
            if ($regResult) {
                $R.OSDisplayVersion = ($regResult.ToString() -split '\s+')[-1]
            }
        } catch {}

        # Extensión telefónica desde SQLite
        try {
            $dllPath = Join-Path $scriptRoot 'libs\System.Data.SQLite.dll'
            $dbPath  = Join-Path $scriptRoot 'database\ComputerNames.sqlite'
            if ((Test-Path $dllPath) -and (Test-Path $dbPath)) {
                $alreadyLoaded = [System.AppDomain]::CurrentDomain.GetAssemblies() |
                    Where-Object { $_.GetName().Name -eq 'System.Data.SQLite' }
                if (-not $alreadyLoaded) { Add-Type -Path $dllPath -ErrorAction Stop }
                $extConn = New-Object System.Data.SQLite.SQLiteConnection "Data Source=$dbPath;Version=3;"
                $extConn.Open()
                $extCmd = $extConn.CreateCommand()
                $extCmd.CommandText = "SELECT extension FROM extensions WHERE equipo = @e"
                $null = $extCmd.Parameters.AddWithValue('@e', $cn.Trim().ToUpper())
                $extVal = $extCmd.ExecuteScalar()
                $extConn.Close(); $extConn.Dispose()
                if ($extVal) { $R.Extension = [string]$extVal }
            }
        } catch {}

        # Servidor de logon del equipo (Group Policy History)
        try {
            $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $cn)
            $subKey = $regKey.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\History')
            if ($subKey) {
                $R.LogonServerEquipo = $subKey.GetValue('DCName')
                $subKey.Close()
            }
            $regKey.Close()
        } catch {}
        # Fallback si no hay entrada en GP History
        if (-not $R.LogonServerEquipo -and $R.Domain) {
            $R.LogonServerEquipo = "\\$($R.Domain)"
        }

        # Servidor de logon del usuario activo (HKU\<SID>\Volatile Environment)
        try {
            if ($R.InteractiveUser) {
                $cleanUser = $R.InteractiveUser
                if ($cleanUser -match '^(.+)\\(.+)$') { $cleanUser = $Matches[2] }

                $urk = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('Users', $cn)
                foreach ($sid in $urk.GetSubKeyNames()) {
                    if ($sid -notmatch '^S-1-5-21-') { continue }
                    try {
                        $vk = $urk.OpenSubKey("$sid\Volatile Environment")
                        if ($vk) {
                            $un = $vk.GetValue('USERNAME')
                            if ($un -eq $cleanUser) {
                                $R.LogonServerUsuario = $vk.GetValue('LOGONSERVER')
                                $vk.Close()
                                break
                            }
                            $vk.Close()
                        }
                    } catch {}
                }
                $urk.Close()
            }
        } catch {}

        # Emitir resultado final completo
        Write-Output $R

    } -ArgumentList $ComputerNameParam, $Global:ScriptRoot

    # Timer de 500ms: cada tick comprueba el estado del Job y procesa los resultados al terminar
    $global:StreamTimer          = New-Object System.Windows.Forms.Timer
    $global:StreamTimer.Interval = 500
    $global:StreamTimer.Add_Tick({
        try {
            # Cortar si se supera el timeout de 40s
            $elapsed = (Get-Date) - $global:StreamStart
            if ($elapsed.TotalSeconds -gt 40) {
                try { Add-Logs -text "⏱️ Timeout: La recogida de datos superó los 40 segundos. Cancelando..." } catch {}
                try { Stop-Job $global:StreamJob -ErrorAction SilentlyContinue } catch {}

                # Recuperar el ultimo parcial emitido antes del timeout
                try {
                    $allOutput = @(Receive-Job $global:StreamJob -ErrorAction SilentlyContinue)
                    $lastResult = $null
                    foreach ($item in $allOutput) {
                        if ($item -is [hashtable]) { $lastResult = $item }
                    }
                    if ($lastResult) {
                        Apply-CollectionResults -R $lastResult
                        try { Add-Logs -text "⚠️ Timeout (40s) — se muestran datos parciales recuperados" } catch {}
                    } else {
                        try { Add-Logs -text "⚠️ Timeout (40s) — no se pudieron recuperar datos" } catch {}
                    }
                } catch {}

                Cleanup-StreamingCollection
                return
            }

            # Sigue en ejecucion, esperar al siguiente tick
            if (-not $global:StreamJob) { return }
            $state = $global:StreamJob.State
            if ($state -eq 'Running') { return }

            # Job terminado: procesar resultados
            $t2 = [math]::Round($elapsed.TotalSeconds, 1)

            if ($state -eq 'Completed') {
                try {
                    # Receive-Job devuelve todos los parciales emitidos; tomamos el ultimo
                    $allOutput = @(Receive-Job $global:StreamJob -ErrorAction Stop)
                    $lastResult = $null
                    foreach ($item in $allOutput) {
                        if ($item -is [hashtable]) { $lastResult = $item }
                    }
                    if ($lastResult) {
                        Apply-CollectionResults -R $lastResult
                        try { Add-Logs -text "✅ Recogida de datos completada (${t2}s)" } catch {}
                    }
                } catch {
                    try { Add-Logs -text "❌ Error al procesar resultados: $($_.Exception.Message)" } catch {}
                }
            } else {
                try { Add-Logs -text "❌ El Job terminó con estado: $state" } catch {}
            }

            Cleanup-StreamingCollection

        } catch {
            $global:StreamUpdating = $false
            try { Add-Logs -text "❌ Error en timer tick: $($_.Exception.Message)" } catch {}
            Cleanup-StreamingCollection
        }
    })
    $global:StreamTimer.Start()
    return $true
}



