
<#
.SYNOPSIS
    Lista impresoras USB y LPT en un equipo remoto o local
.DESCRIPTION
    Script que conecta via WMI/CIM a un equipo y lista todas las impresoras detectadas (USB y LPT)
    Genera un reporte HTML que se abre automáticamente en el navegador predeterminado
.PARAMETER ComputerName
    Nombre del equipo a consultar. Por defecto usa el equipo local.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName = $env:COMPUTERNAME
)

# Configuración de errores
$ErrorActionPreference = "SilentlyContinue"

# Función para obtener impresoras instaladas (no USB ni LPT: PDF, red, compartidas...)
function Get-InstalledNetworkPrinters {
    param(
        [string]$Computer
    )
    
    $printers = @()
    
    try {
        $wmiPrinters = Get-WmiObject -Class Win32_Printer -ComputerName $Computer -ErrorAction Stop
        foreach ($p in $wmiPrinters) {
            # Excluir impresoras USB y LPT — queremos sólo virtuales, red y compartidas
            if ($p.PortName -notmatch '^USB|^LPT') {
                $printers += [PSCustomObject]@{
                    Name       = $p.Name
                    Driver     = $p.DriverName
                    Port       = $p.PortName
                    Shared     = if ($p.Shared) { 'SI' } else { 'NO' }
                    ShareName  = $p.ShareName
                    Location   = $p.Location
                    Comment    = $p.Comment
                    Default    = if ($p.Default) { 'SI' } else { 'NO' }
                    Status     = $p.PrinterStatus
                }
            }
        }
    } catch {
        Write-Warning "Error obteniendo impresoras WMI: $_"
    }
    
    return $printers
}

# Helper: busca un S/N real (sin '&') para un ContainerID dado — equivale a BuscaContainerID() del .vbs
function Find-RealSerial {
    param($Reg, [string]$UsbPath, [string]$ContainerID)
    $result = ""
    try {
        $usbKey = $Reg.OpenSubKey($UsbPath)
        if ($null -eq $usbKey) { return $result }
        foreach ($vp in $usbKey.GetSubKeyNames()) {
            $vpKey = $Reg.OpenSubKey("$UsbPath\$vp")
            if ($null -eq $vpKey) { continue }
            foreach ($inst in $vpKey.GetSubKeyNames()) {
                if ($inst -notmatch '&') {   # sin '&' → S/N real del fabricante
                    $ik = $Reg.OpenSubKey("$UsbPath\$vp\$inst")
                    if ($null -ne $ik) {
                        $cid = $ik.GetValue("ContainerID")
                        if ($cid -eq $ContainerID) {
                            if ($result -ne "") { $result += " - " }
                            $result += $inst
                        }
                    }
                }
            }
        }
    } catch {}
    return $result
}

# Helper: convierte un GUID binario de 16 bytes (little-endian) a string con llaves
function Convert-BinaryGuid {
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Count -lt 16) { return "" }
    return ("{" +
        $Bytes[3].ToString("x2") + $Bytes[2].ToString("x2") + $Bytes[1].ToString("x2") + $Bytes[0].ToString("x2") + "-" +
        $Bytes[5].ToString("x2") + $Bytes[4].ToString("x2") + "-" +
        $Bytes[7].ToString("x2") + $Bytes[6].ToString("x2") + "-" +
        $Bytes[8].ToString("x2")  + $Bytes[9].ToString("x2")  + "-" +
        $Bytes[10].ToString("x2") + $Bytes[11].ToString("x2") + $Bytes[12].ToString("x2") +
        $Bytes[13].ToString("x2") + $Bytes[14].ToString("x2") + $Bytes[15].ToString("x2") + "}")
}

# Función para obtener impresoras USB — lógica fiel al .vbs original
# Algoritmo: USB enum → instancia = S/N candidato → ContainerID vincula a USBPRINT (modelo)
#            Si S/N tiene '&' (ID compuesto Windows) → Find-RealSerial busca S/N limpio
#            Instalación verificada comparando ContainerID binario de PnPData
function Get-USBPrinters {
    param([string]$Computer)
    $printers = @()
    try {
        $reg          = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
        $usbPath      = "SYSTEM\CurrentControlSet\Enum\USB"
        $usbPrintPath = "SYSTEM\CurrentControlSet\Enum\USBPRINT"
        $printersPath = "SYSTEM\CurrentControlSet\Control\Print\Printers"

        $usbKey = $reg.OpenSubKey($usbPath)
        if ($null -eq $usbKey) { $reg.Close(); return $printers }

        foreach ($vid_pid in $usbKey.GetSubKeyNames()) {
            $vpKey = $reg.OpenSubKey("$usbPath\$vid_pid")
            if ($null -eq $vpKey) { continue }

            foreach ($inst in $vpKey.GetSubKeyNames()) {
                $instKey = $reg.OpenSubKey("$usbPath\$vid_pid\$inst")
                if ($null -eq $instKey) { continue }

                $deviceDesc  = $instKey.GetValue("DeviceDesc")
                $containerID = $instKey.GetValue("ContainerID")

                # Solo impresoras — igual que el .vbs
                if ($deviceDesc -notmatch 'PRINT|LASER|DOT4') { continue }

                $serialNum  = $inst
                $vid_pid_17 = $vid_pid.Substring(0, [Math]::Min(17, $vid_pid.Length))
                $model      = ""
                $hardwareID = ""

                # Buscar modelo y HardwareID en USBPRINT por ContainerID
                $upKey = $reg.OpenSubKey($usbPrintPath)
                if ($null -ne $upKey) {
                    foreach ($mKey in $upKey.GetSubKeyNames()) {
                        $mSubKey = $reg.OpenSubKey("$usbPrintPath\$mKey")
                        if ($null -eq $mSubKey) { continue }
                        foreach ($mInst in $mSubKey.GetSubKeyNames()) {
                            $mInstKey = $reg.OpenSubKey("$usbPrintPath\$mKey\$mInst")
                            if ($null -eq $mInstKey) { continue }
                            $cid2 = $mInstKey.GetValue("ContainerID")
                            if ($cid2 -eq $containerID) {
                                if ($mKey -ne "UnknownPrinter" -or $model -eq "") {
                                    $hwRaw = $mInstKey.GetValue("HardwareID")
                                    if ($hwRaw) {
                                        $hwArr  = @($hwRaw)
                                        $lastHw = $hwArr[$hwArr.Count - 1]
                                        if ($hardwareID -ne "") { $hardwareID += " : " }
                                        $hardwareID += $lastHw
                                    }
                                    $dd = $mInstKey.GetValue("DeviceDesc")
                                    if ($dd -match "%;(.+)") { $dd = $matches[1].Trim() }
                                    if ($model -eq "" -or $mKey -ne "UnknownPrinter") {
                                        $model = if ($model -ne "") { "$model : $dd" } else { $dd }
                                    }
                                }
                            }
                        }
                    }
                }

                # Si el S/N tiene '&' es ID compuesto de Windows → buscar S/N real
                if ($serialNum -match '&') {
                    $realSN = Find-RealSerial -Reg $reg -UsbPath $usbPath -ContainerID $containerID
                    if ($realSN -ne "") { $serialNum = $realSN }
                }

                # Decodificaciones y modelos por VID/PID (igual que el .vbs)
                switch ($vid_pid_17) {
                    "VID_04B8&PID_0E15" {
                        $model = "EPSON TM-T20II"
                        if ($serialNum.Length -ge 10 -and $serialNum -notmatch '&') {
                            $sn = ""
                            for ($i = 0; $i -lt 8; $i += 2) {
                                $sn += [char][Convert]::ToInt32($serialNum.Substring($i, 2), 16)
                            }
                            $sn += $serialNum.Substring(8, [Math]::Min(6, $serialNum.Length - 8))
                            $serialNum = $sn
                        }
                    }
                    "VID_04B8&PID_0007" {
                        $model = "EPSON AL-M2000"
                        if ($serialNum.Length -gt 3) { $serialNum = $serialNum.Substring(3, [Math]::Min(10, $serialNum.Length - 3)) }
                    }
                    "VID_0A5F&PID_00A3" { $model = "Zebra ZDesigner LP 2824 Plus" }
                    "VID_04F9&PID_002B" { $model = "BROTHER HL-5250DN" }
                    "VID_04F9&PID_2039" { $model = "BROTHER TD-4000" }
                    "VID_03F0&PID_2B17" { $model = "HP LaserJet 1020" }
                    "VID_04B8&PID_0005" { $model = "EPSON EPL-6200" }
                    "VID_04F9&PID_007F" { $model = "BROTHER HL-5100DN" }
                    "VID_03F0&PID_1117" { $model = "HP LaserJet 1300n" }
                    "VID_03F0&PID_0317" { $model = "HP LaserJet 1200" }
                    "VID_03F0&PID_0C17" { $model = "HP LaserJet 1010" }
                }

                # Limpiar S/N de HP (quitar "00" inicial)
                if ($vid_pid_17 -like "VID_03F0*" -and $serialNum -like "00*") {
                    $serialNum = $serialNum.Substring(2)
                }

                # Comprobar si está instalada usando el GUID binario de PnPData
                $instalada = "NO";  $portName = "";  $shared = "NO";  $shareName = ""
                $pKey = $reg.OpenSubKey($printersPath)
                if ($null -ne $pKey) {
                    foreach ($prName in $pKey.GetSubKeyNames()) {
                        $pnpKey = $reg.OpenSubKey("$printersPath\$prName\PnPData")
                        if ($null -ne $pnpKey) {
                            $rawGuid = $pnpKey.GetValue("DeviceContainerId")
                            if ($rawGuid -is [byte[]] -and $rawGuid.Count -ge 16) {
                                $guidStr = Convert-BinaryGuid -Bytes $rawGuid
                                if ($containerID -eq $guidStr) {
                                    $instalada = "SI"
                                    $prSubKey = $reg.OpenSubKey("$printersPath\$prName")
                                    if ($null -ne $prSubKey) {
                                        $pn  = $prSubKey.GetValue("Port");       if ($pn)  { $portName  = $pn }
                                        $sn2 = $prSubKey.GetValue("Share Name"); if ($sn2) { $shared = "SI"; $shareName = $sn2 }
                                    }
                                }
                            }
                        }
                    }
                }

                # Comprobar si está conectada ahora (WMI PnPEntity)
                $conectada = "NO"
                try {
                    $pnp = Get-WmiObject -Class Win32_PnPEntity -ComputerName $Computer `
                        -Filter "DeviceID LIKE '%$vid_pid%'" -ErrorAction SilentlyContinue
                    if ($pnp) { $conectada = "SI" }
                } catch {}

                # Normalizar S/N numérico cero
                if ($serialNum -match '^\d+$' -and [long]$serialNum -eq 0) { $serialNum = "0" }

                # Filtro igual al .vbs: descartar entradas sin información relevante
                $snBad = ($serialNum -match '&') -or ($serialNum.Trim() -eq "") -or ($serialNum -eq "0")
                if ($hardwareID.Trim() -eq "" -and $instalada -eq "NO" -and $conectada -eq "NO" -and $snBad) { continue }

                $printers += [PSCustomObject]@{
                    Name         = if ($model) { $model } else { $vid_pid_17 }
                    Port         = $portName
                    SerialNumber = if ($serialNum -match '&' -or $serialNum.Trim() -eq "") { "—" } else { $serialNum }
                    Installed    = $instalada
                    Connected    = $conectada
                    Shared       = $shared
                    ShareName    = $shareName
                    HardwareID   = $hardwareID
                    VID_PID      = $vid_pid_17
                }
            }
        }
        $reg.Close()
    } catch {
        Write-Warning "Error obteniendo impresoras USB: $_"
    }
    return $printers
}

# Función para obtener impresoras LPT
function Get-LPTPrinters {
    param(
        [string]$Computer
    )
    
    $printers = @()
    
    try {
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
        $lptPath = "SYSTEM\CurrentControlSet\Enum\LPTenum"
        
        $lptKey = $reg.OpenSubKey($lptPath)
        if ($null -ne $lptKey) {
            foreach ($subKey in $lptKey.GetSubKeyNames()) {
                $deviceKey = $reg.OpenSubKey("$lptPath\$subKey")
                if ($null -ne $deviceKey) {
                    $deviceDesc = $deviceKey.GetValue("DeviceDesc")
                    $hwid = $deviceKey.GetValue("HardwareID")
                    
                    if ($deviceDesc) {
                        $printer = @{
                            Model = $deviceDesc
                            HardwareID = ($hwid -join " ; ")
                        }
                        $printers += $printer
                    }
                }
            }
        }
        
        $reg.Close()
    } catch {
        Write-Warning "Error accediendo a impresoras LPT: $_"
    }
    
    return $printers
}

# Mostrar resultados formateados en la ventana de consola
function Show-PrinterReport {
    param(
        [string]$Computer,
        [array]$InstalledPrinters,
        [array]$USBPrinters,
        [array]$LPTPrinters,
        [string]$ErrorMessage
    )
    
    $sep = "=" * 72

    Write-Host ""
    Write-Host $sep -ForegroundColor Cyan
    Write-Host "  IMPRESORAS DETECTADAS EN: $($Computer.ToUpper())" -ForegroundColor White
    Write-Host $sep -ForegroundColor Cyan

    if ($ErrorMessage) {
        Write-Host ""
        Write-Host "  Error de conexion: $ErrorMessage" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Posibles causas:" -ForegroundColor Yellow
        Write-Host "    - El equipo esta apagado o no es accesible en la red" -ForegroundColor Yellow
        Write-Host "    - Firewall bloqueando la conexion WMI/RPC" -ForegroundColor Yellow
        Write-Host "    - Sin permisos de administrador remoto" -ForegroundColor Yellow
        Write-Host "    - Servicio WMI detenido en el equipo remoto" -ForegroundColor Yellow
        Write-Host ""
        Write-Host $sep -ForegroundColor DarkCyan
        return
    }

    Write-Host ""
    Write-Host "  Conexion establecida correctamente con $Computer" -ForegroundColor Green

    # ── Impresoras instaladas (red, virtuales, compartidas) ──────────────────
    Write-Host ""
    Write-Host "  IMPRESORAS INSTALADAS  (red, virtuales y compartidas)" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 68) -ForegroundColor DarkCyan
    if ($InstalledPrinters -and $InstalledPrinters.Count -gt 0) {
        $InstalledPrinters | Format-Table -AutoSize -Property `
            @{N='Nombre';           E={$_.Name}},
            @{N='Driver';           E={$_.Driver}},
            @{N='Puerto';           E={$_.Port}},
            @{N='Compartida';       E={$_.Shared}},
            @{N='Nombre Compartido';E={$_.ShareName}},
            @{N='Predeterminada';   E={$_.Default}}
    } else {
        Write-Host "  No se detectaron impresoras de red, virtuales o compartidas." -ForegroundColor Gray
        Write-Host ""
    }

    # ── Impresoras USB ────────────────────────────────────────────────────────
    Write-Host "  IMPRESORAS CON CONEXION USB" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 68) -ForegroundColor DarkCyan
    if ($USBPrinters -and $USBPrinters.Count -gt 0) {
        $USBPrinters | Format-Table -AutoSize -Property `
            @{N='Modelo';           E={$_.Name}},
            @{N='Puerto';           E={$_.Port}},
            @{N='Num. Serie';       E={$_.SerialNumber}},
            @{N='Instalada';        E={$_.Installed}},
            @{N='Conectada';        E={$_.Connected}},
            @{N='Compartida';       E={$_.Shared}},
            @{N='Nombre Compartido';E={$_.ShareName}}
    } else {
        Write-Host "  No se detectaron impresoras USB en este equipo." -ForegroundColor Gray
        Write-Host ""
    }

    # ── Impresoras LPT ────────────────────────────────────────────────────────
    Write-Host "  IMPRESORAS CON CONEXION LPT" -ForegroundColor Cyan
    Write-Host ("  " + "-" * 68) -ForegroundColor DarkCyan
    if ($LPTPrinters -and $LPTPrinters.Count -gt 0) {
        foreach ($p in $LPTPrinters) {
            Write-Host "  $($p.Model)" -ForegroundColor White
            if ($p.HardwareID) { Write-Host "    HardwareID: $($p.HardwareID)" -ForegroundColor Gray }
        }
        Write-Host ""
    } else {
        Write-Host "  No se detectaron impresoras LPT en este equipo." -ForegroundColor Gray
        Write-Host ""
    }

    Write-Host $sep -ForegroundColor DarkCyan
    Write-Host "  Generado: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')  |  Equipo: $Computer" -ForegroundColor Gray
    Write-Host $sep -ForegroundColor DarkCyan
    Write-Host ""
}

# SCRIPT PRINCIPAL

# Ajustar tamanyo de la ventana de consola para que la tabla no se trunque
try {
    $Host.UI.RawUI.WindowTitle = "Impresoras | $ComputerName"
    $buf = $Host.UI.RawUI.BufferSize
    if ($buf.Width -lt 160) {
        $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(160, 3000)
    }
    $win = $Host.UI.RawUI.WindowSize
    if ($win.Width -lt 140) {
        $Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size(140, [Math]::Max($win.Height, 40))
    }
} catch {}

Write-Host ""
Write-Host "  Verificando conectividad con $ComputerName..." -ForegroundColor Yellow

$pingResult = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue

if (-not $pingResult) {
    Show-PrinterReport -Computer $ComputerName -InstalledPrinters @() -USBPrinters @() -LPTPrinters @() `
        -ErrorMessage "No se puede conectar al equipo $ComputerName. El equipo no responde al ping."
} else {
    Write-Host "  Ping exitoso" -ForegroundColor Green

    $usbPrinters       = @()
    $lptPrinters       = @()
    $installedPrinters = @()
    $errorMsg          = $null

    try {
        Write-Host "  Obteniendo impresoras instaladas (red/virtuales/compartidas)..." -ForegroundColor Yellow
        $installedPrinters = Get-InstalledNetworkPrinters -Computer $ComputerName
        Write-Host "  $($installedPrinters.Count) impresoras instaladas encontradas" -ForegroundColor Green

        Write-Host "  Obteniendo impresoras USB..." -ForegroundColor Yellow
        $usbPrinters = Get-USBPrinters -Computer $ComputerName
        Write-Host "  $($usbPrinters.Count) impresoras USB encontradas" -ForegroundColor Green

        Write-Host "  Obteniendo impresoras LPT..." -ForegroundColor Yellow
        $lptPrinters = Get-LPTPrinters -Computer $ComputerName
        Write-Host "  $($lptPrinters.Count) impresoras LPT encontradas" -ForegroundColor Green

    } catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        $errorMsg = "Error al obtener la informacion: $($_.Exception.Message)"
    }

    Show-PrinterReport -Computer $ComputerName `
        -InstalledPrinters $installedPrinters `
        -USBPrinters $usbPrinters `
        -LPTPrinters $lptPrinters `
        -ErrorMessage $errorMsg
}
