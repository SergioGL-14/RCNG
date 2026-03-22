#==================================================================
# RCNG / NRC_APP
# Script principal de la aplicación. Aquí se monta la interfaz, se inicializan
# los datos locales y se coordinan módulos, menús y subaplicaciones.
#==================================================================
# Resolver la carpeta real de ejecución, tanto en modo .ps1 como en modo .exe.
$Global:ScriptRoot = if ($MyInvocation.MyCommand.CommandType -eq 'ExternalScript') {
    Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
} else {
    Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
}

# En el ejecutable compilado se silencian salidas residuales a consola.
if ($env:PS2EXE) {
    [Console]::SetOut([System.IO.TextWriter]::Null)
    [Console]::SetError([System.IO.TextWriter]::Null)
}
# Cargar los ensamblados base usados por WinForms y acceso a datos.
[void][Reflection.Assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
[void][Reflection.Assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
[void][Reflection.Assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
# Estado global de la recogida de datos asíncrona.
$global:StreamJob         = $null   # Job activo de recoleccion
$global:StreamTimer       = $null   # Timer WinForms de 500ms
$global:StreamStart       = $null   # Marca de tiempo de inicio (para timeout)
$global:StreamRunning     = $false  # True mientras el Job esta en marcha
$global:StreamUI          = $null   # Referencias a controles UI
$global:StreamLineIndices = $null   # Mapa key->lineaRTB para actualizacion in-situ
$global:StreamUpdating    = $false  # Evita re-entrada en el evento TextChanged

# Cargar el gestor de sincronización local/servidor si está disponible.
try {
	$sharedDataMgrPath = Join-Path $Global:ScriptRoot 'modules\SharedDataManager.psm1'
	if (Test-Path $sharedDataMgrPath) {
		Import-Module $sharedDataMgrPath -Force -DisableNameChecking -ErrorAction Stop
	} else {
		Write-Output " SharedDataManager.psm1 no encontrado en: $sharedDataMgrPath"
	}
} catch { Write-Output " Error cargando SharedDataManager.psm1: $($_.Exception.Message)" }

# Inicializar la BD de equipos para búsquedas rápidas y autocompletado.
# Si el servidor compartido responde, se usa la copia activa resuelta por
# SharedDataManager; si no, se trabaja con la local.
$global:ComputerDBConn = $null
$global:ComputerDBAvailable = $false

function Initialize-ComputerDB {
	param(
		[string]$ScriptRootParam = $Global:ScriptRoot,
		[string]$DllRelative = 'libs\System.Data.SQLite.dll'
	)

	# Cerrar conexión anterior si existe (necesario al refrescar)
	if ($global:ComputerDBConn -ne $null) {
		try { $global:ComputerDBConn.Close(); $global:ComputerDBConn.Dispose() } catch {}
		$global:ComputerDBConn = $null
		$global:ComputerDBAvailable = $false
	}

	try {
		$dllPath = Join-Path $ScriptRootParam $DllRelative
		if (-not (Test-Path $dllPath)) { throw "SQLite DLL not found: $dllPath" }
		Add-Type -Path $dllPath -ErrorAction Stop

		# Resolver ruta: servidor (UNC) o local como fallback
		if (Get-Command 'Get-ActiveComputerDBPath' -ErrorAction SilentlyContinue) {
			$dbFile = Get-ActiveComputerDBPath
		} else {
			$dbFile = Join-Path $ScriptRootParam 'database\ComputerNames.sqlite'
		}

		if (-not (Test-Path $dbFile -ErrorAction SilentlyContinue)) { throw "Database not found: $dbFile" }

		# ParseViaFramework debe ser propiedad del objeto, NO keyword del connection string
		$connString = "Data Source=$dbFile;Version=3;Journal Mode=DELETE;BusyTimeout=5000;"
		$conn = New-Object System.Data.SQLite.SQLiteConnection -ArgumentList $connString
		$conn.ParseViaFramework = $true
		$conn.Open()

		$global:ComputerDBConn = $conn
		$global:ComputerDBAvailable = $true
		Write-Output "✅ ComputerNames DB loaded: $dbFile"
		return $global:ComputerDBConn
	} catch {
		$global:ComputerDBConn = $null
		$global:ComputerDBAvailable = $false
		Write-Output "⚠️ No se pudo inicializar ComputerNames DB: $($_.Exception.Message)"
		return $null
	}
}

function Invoke-ComputerDBQuery {
	param(
		[string]$Query,
		[Hashtable]$Parameters = @{},
		[switch]$Scalar
	)
	if (-not $global:ComputerDBAvailable) { Initialize-ComputerDB | Out-Null }
	if (-not $global:ComputerDBConn) { throw "Computer DB not available" }

	$cmd = $global:ComputerDBConn.CreateCommand()
	$cmd.CommandText = $Query
	foreach ($k in $Parameters.Keys) {
		$null = $cmd.Parameters.AddWithValue($k, $Parameters[$k])
	}

	if ($Scalar) { return $cmd.ExecuteScalar() }

	$dt = New-Object System.Data.DataTable
	(New-Object System.Data.SQLite.SQLiteDataAdapter($cmd)).Fill($dt) | Out-Null

	$out = @()
	foreach ($row in $dt.Rows) {
		$obj = @{}
		foreach ($col in $dt.Columns) { $obj[$col.ColumnName] = $row[$col] }
		$out += [PSCustomObject]$obj
	}
	return $out
}

function Get-ComputerByFilterDB {
	param(
		[string]$Filter,
		[int]$Limit = 200
	)
	if (-not $global:ComputerDBAvailable) { Initialize-ComputerDB | Out-Null }
	if (-not $global:ComputerDBConn) { return @() }

	$sql = @"
	SELECT ou, equipo, orig_line FROM computers
	WHERE UPPER(equipo) LIKE '%' || UPPER(@f) || '%'
	   OR UPPER(ou) LIKE '%' || UPPER(@f) || '%'
	LIMIT @limit;
"@
	return Invoke-ComputerDBQuery -Query $sql -Parameters @{ '@f' = $Filter; '@limit' = [int]$Limit }
}

# Inicializar la DB de forma inmediata para que esté disponible en el proceso principal
try { Initialize-ComputerDB | Out-Null } catch {}

# Importar modulo DHCP
try {
	$dhcpModulePath = Join-Path $Global:ScriptRoot 'modules\DHCP.psm1'
	if (Test-Path $dhcpModulePath) { Import-Module $dhcpModulePath -Force -ErrorAction SilentlyContinue }
} catch {}

#==================================================================
# BLOQUE: Recoleción de datos — módulo DataCollection.psm1
#==================================================================
try {
    $dcModulePath = Join-Path $Global:ScriptRoot 'modules\DataCollection.psm1'
    if (Test-Path $dcModulePath) {
        Import-Module $dcModulePath -Force -DisableNameChecking -ErrorAction Stop
    } else {
        Write-Output " DataCollection.psm1 no encontrado en: $dcModulePath"
    }
} catch { Write-Output " Error cargando DataCollection.psm1: $($_.Exception.Message)" }

#==================================================================
# BLOQUE: Función principal de entrada de la aplicación
#==================================================================
function Main {
	Param ([String]$Commandline)
	if(Show-MainForm_pff -eq "OK"){
	}
	$script:ExitCode = 0
}
#==================================================================
# BLOQUE: Funciones globales y auxiliares del entorno NRC
#==================================================================
#------------------------------------------------------------------
# SUBBLOQUE: Obtener nombre del equipo desde la interfaz
#------------------------------------------------------------------
function global:Get-ComputerTxtBox {
	$global:ComputerName = $textbox_computername.Text
}
#------------------------------------------------------------------
# SUBBLOQUE: Añadir contenido a la RichTextBox principal
#------------------------------------------------------------------
function global:Add-RichTextBox {
	[CmdletBinding()]
	param ($text)
	# ignore if the control hasn't been created (e.g. running in test mode)
	if ($null -eq $global:richtextbox_output) { return }
	$global:richtextbox_output.Text += "$text`n`n"
}

function Add-RichTextBoxCheck {
	[CmdletBinding()]
	param ($text, [System.Drawing.Color]$color = [System.Drawing.Color]::Black)
	$richtextbox_output.SelectionStart = $richtextbox_output.TextLength
	$richtextbox_output.SelectionLength = 0
	$richtextbox_output.SelectionColor = $color
	$richtextbox_output.AppendText("$text`n")
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener fecha actual con formato sortable
#------------------------------------------------------------------
function Get-datesortable {
	$global:datesortable = Get-Date -Format "yyyyMMdd-HH':'mm':'ss"
	return $global:datesortable
}
#------------------------------------------------------------------
# SUBBLOQUE: Añadir contenido al registro de logs
#------------------------------------------------------------------
function global:Add-Logs {
	[CmdletBinding()]
	param ($text)
	# if the UI element hasn't been built just ignore
	if ($null -eq $global:richtextbox_Logs) { return }
	Get-datesortable
	$global:richtextbox_Logs.Text += "[$global:datesortable] - $text`r"
}
Set-Alias alogs Add-Logs -Description "Add content to the RichTextBoxLogs"
Set-Alias Add-Log Add-Logs -Description "Add content to the RichTextBoxLogs"
#------------------------------------------------------------------
# SUBBLOQUE: Limpiar cajas de texto
#------------------------------------------------------------------
function Clear-RichTextBox {
	$richtextbox_output.Text = ""
}

function Clear-Logs {
	$richtextbox_logs.Text = ""
}
#------------------------------------------------------------------
# SUBBLOQUE: Copiar texto al portapapeles
#------------------------------------------------------------------
function Add-ClipBoard ($text){
	Add-Type -AssemblyName System.Windows.Forms
	$tb = New-Object System.Windows.Forms.TextBox
	$tb.Multiline = $true
	$tb.Text = $text
	$tb.SelectAll()
	$tb.Copy()	
}
#------------------------------------------------------------------
# SUBBLOQUE: Comprobación de puertos TCP remotos
#------------------------------------------------------------------
function Test-TcpPort ($ComputerName,[int]$port = 80) {
	$socket = new-object Net.Sockets.TcpClient
	$socket.Connect($ComputerName, $port)
	if ($socket.Connected) {
		$status = "Open"
		$socket.Close()
	} else {
		$status = "Closed / Filtered"
	}
	$socket = $null
	Add-RichTextBox "ComputerName:$ComputerName`nPort:$port`nStatus:$status"
}
#------------------------------------------------------------------
# SUBBLOQUE: Activar o desactivar Escritorio Remoto
#------------------------------------------------------------------
function Set-RDPEnable ($ComputerName = '.') {
	$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName)
	$regKey = $regKey.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Terminal Server" ,$True)
	$regkey.SetValue("fDenyTSConnections",0)
	$regKey.flush()
	$regKey.Close()
}

function Set-RDPDisable ($ComputerName = '.') {
	$regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $ComputerName)
	$regKey = $regKey.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Terminal Server" ,$True)
	$regkey.SetValue("fDenyTSConnections",1)
	$regKey.flush()
	$regKey.Close()
}
#------------------------------------------------------------------
# SUBBLOQUE: Lanzar procesos externos con opciones
#------------------------------------------------------------------
function Start-Proc {
	param (
		[string]$exe = $(Throw "An executable must be specified"),
		[string]$arguments,
		[switch]$hidden,
		[switch]$waitforexit
	)
	$startinfo = new-object System.Diagnostics.ProcessStartInfo 
	$startinfo.FileName = $exe
	$startinfo.Arguments = $arguments
	if ($hidden){
		$startinfo.WindowStyle = "Hidden"
		$startinfo.CreateNoWindow = $TRUE
	}
	$process = [System.Diagnostics.Process]::Start($startinfo)
	if ($waitforexit) {$process.WaitForExit()}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener tiempo de actividad (uptime)
#------------------------------------------------------------------
function Get-Uptime {
	param($ComputerName = "localhost")
	$wmi = Get-WmiObject -class Win32_OperatingSystem -computer $ComputerName
	$LBTime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
	[TimeSpan]$uptime = New-TimeSpan $LBTime $(get-date)
	Write-Output $uptime
}
#------------------------------------------------------------------
# SUBBLOQUE: Ejecutar actualización de directivas (GPUpdate)
#------------------------------------------------------------------
function Invoke-GPUpdate {
	param($ComputerName = ".")
	$targetOSInfo = Get-WmiObject -ComputerName $ComputerName -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
	if ($null -eq $targetOSInfo) {
		return "Unable to connect to $ComputerName"
	} else {
		if ($targetOSInfo.version -ge 5.1) {
			Invoke-WmiMethod -ComputerName $ComputerName -Path win32_process -Name create -ArgumentList "gpupdate /target:Computer /force /wait:0"
		} else {
			Invoke-WmiMethod -ComputerName $ComputerName -Path win32_process -Name create –ArgumentList "secedit /refreshpolicy machine_policy /enforce"
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener información del archivo de paginación (Pagefile)
#------------------------------------------------------------------
function Get-PageFile {
	[Cmdletbinding()]
	Param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	Process {
		if($ComputerName -match "(.*)(\\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\\$)$",'$1'
		}
		if(Test-Connection $ComputerName -Count 1 -Quiet) {
			try {
				$PagingFiles = Get-WmiObject Win32_PageFile -ComputerName $ComputerName -ErrorAction SilentlyContinue
				if($PagingFiles) {
					foreach($PageFile in $PagingFiles) {
						$myobj = @{
							ComputerName = $ComputerName
							Name = $PageFile.Name
							SizeGB = [int]($PageFile.FileSize / 1GB)
							InitialSize = $PageFile.InitialSize
							MaximumSize = $PageFile.MaximumSize
						}
						$obj = New-Object PSObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.Computer.PageFile')
						$obj
					}
				} else {
					$Pagefile = Get-ChildItem \\$ComputerName\c$\pagefile.sys -Force -ErrorAction SilentlyContinue 
					if($PageFile) {
						$myobj = @{
							ComputerName = $ComputerName
							Name = $PageFile.Name
							SizeGB = [int]($Pagefile.Length / 1GB)
							InitialSize = "System Managed"
							MaximumSize = "System Managed"
						}
						$obj = New-Object PSObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.Computer.PageFile')
						$obj
					} else {
						Write-Host "[Get-PageFile] :: No se ha encontrado el archivo de Paginación."
					}
				}
			} catch {
				Write-Verbose "[Get-PageFile] :: [$ComputerName] Failed with Error: $($lastError[0])"
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener configuración de Pagefile (Win32_PageFileSetting)
#------------------------------------------------------------------
function Get-PageFileSetting {
	[Cmdletbinding()]
	Param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	Process {
		Write-Verbose "[Get-PageFileSetting] :: Process Start"
		if($ComputerName -match "(.*)(\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\$)$",'$1'
		}
		if(Test-Host $ComputerName -TCPPort 135) {
			try {
				Write-Verbose "[Get-PageFileSetting] :: Collecting Paging File Info"
				$PagingFiles = Get-WmiObject Win32_PageFileSetting -ComputerName $ComputerName -EnableAllPrivileges
				if($PagingFiles) {
					foreach($PageFile in $PagingFiles) {
						$PageFile
					}
				} else {
					return "Configuración de PageFile no encontrada. Probablemente esté gestionado automáticamente por el sistema"
				}
			} catch {
				Write-Verbose "[Get-PageFileSetting] :: [$ComputerName] Failed with Error: $($lastError[0])"
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener contenido del archivo HOSTS remoto
#------------------------------------------------------------------
function Get-HostsFile {
	[cmdletbinding(DefaultParameterSetName = 'Default', ConfirmImpact = 'low')]
	Param(
		[Parameter(ValueFromPipeline = $True)]
		[string[]]$Computer
	)
	Begin {
		$psBoundParameters.GetEnumerator() | ForEach-Object {
			Write-Verbose "Parameter: $_"
		}
		if (!$PSBoundParameters['Computer']) {
			Write-Verbose "No computer name given, using local computername"
			[string[]]$Computer = $Env:Computername
		}
		$report = @()
	}
	Process {
		ForEach ($c in $Computer) {
			if (Test-Connection -ComputerName $c -Quiet -Count 1) {
				if (Test-Path "\\$c\C$\Windows\system32\drivers\etc\hosts") {
					$hostsPath = "\\$c\C$\Windows\system32\drivers\etc\hosts"
				} elseif (Test-Path "\\$c\C$\WinNT\system32\drivers\etc\hosts") {
					$hostsPath = "\\$c\C$\WinNT\system32\drivers\etc\hosts"
				} else {
					$report += [PSCustomObject]@{
						Computer = $c
						IPV4 = "NA"
						IPV6 = "NA"
						Hostname = "NA"
						Notes = "Unable to locate host file"
					}
					continue
				}

				Switch -regex -file ($hostsPath) {
					"^\d" {
						$new = $_.Split() | Where-Object {$_ -ne ""}
						$report += [PSCustomObject]@{
							Computer = $c
							IPV4 = $new[0]
							Hostname = $new[1]
							Notes = if ($new.Count -gt 2) { $new[2] } else { "NA" }
						}
					}
					Default {
						if (!($_ -match "^\s*$" -or $_.StartsWith("#"))) {
							$new = $_.Split() | Where-Object {$_ -ne ""}
							$report += [PSCustomObject]@{
								Computer = $c
								IPV6 = $new[0]
								Hostname = $new[1]
								Notes = if ($new.Count -gt 2) { $new[2] } else { "NA" }
							}
						}
					}
				}
			} else {
				$report += [PSCustomObject]@{
					Computer = $c
					IPV4 = "NA"
					IPV6 = "NA"
					Hostname = "NA"
					Notes = "Unable to locate Computer"
				}
			}
		}
	}
	End {
		Write-Output $report
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener información de particiones de disco
#------------------------------------------------------------------
function Get-DiskPartition {
	[Cmdletbinding()]
	Param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	Process {
		Write-Verbose "[Get-DiskPartition] :: Process Start"
		if($ComputerName -match "(.*)(\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\$)$",'$1'
		}
		if(Test-Host $ComputerName -TCPPort 135) {
			try {
				$Partitions = Get-WmiObject Win32_DiskPartition -ComputerName $ComputerName
				foreach($Partition in $Partitions) {
					$myobj = @{
						BlockSize = $Partition.BlockSize
						BootPartition = $Partition.BootPartition
						ComputerName = $ComputerName
						Description = $Partition.Name
						PrimaryPartition = $Partition.PrimaryPartition
						Index = $Partition.Index
						SizeMB = ($Partition.Size/1mb).ToString("n2", [CultureInfo]::InvariantCulture)
						Type = $Partition.Type
						IsAligned = $Partition.StartingOffset % 65536 -eq 0
					}
					$obj = New-Object PSObject -Property $myobj
					$obj.PSTypeNames.Clear()
					$obj.PSTypeNames.Add('BSonPosh.DiskPartition')
					$obj
				}
			} catch {
				Write-Verbose "[Get-DiskPartition] :: [$ComputerName] Failed with Error: $($lastError[0])"
			}
		}
		Write-Verbose "[Get-DiskPartition] :: Process End"
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener espacio en disco (Win32_Volume o LogicalDisk)
#------------------------------------------------------------------
function Get-DiskSpace {
    param (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    try {
        $discos = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $ComputerName -Filter "DriveType=3"

        if (-not $discos) {
            return "No se encontraron unidades físicas en $ComputerName."
        }

        $resultado = foreach ($d in $discos) {
            $totalGB = [math]::Round($d.Size / 1GB, 2)
            $libreGB = [math]::Round($d.FreeSpace / 1GB, 2)
            $porcentaje = if ($d.Size -ne 0) {
                [math]::Round(($d.FreeSpace / $d.Size) * 100, 2)
            } else {
                0
            }

            [PSCustomObject]@{
                Unidad       = $d.DeviceID
                Nombre       = $d.VolumeName
                Sistema      = $d.FileSystem
                Total_GB     = "$totalGB GB"
                Libre_GB     = "$libreGB GB"
                Libre_Porc   = "$porcentaje %"
            }
        }

        return ($resultado | Format-Table -AutoSize | Out-String).Trim()
    }
    catch {
        return "Error al obtener el espacio en disco de $ComputerName : $_"
    }
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener información del procesador (Win32_Processor)
#------------------------------------------------------------------
function Get-Processor {
	[Cmdletbinding()]
	Param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	Process {
		if($ComputerName -match "(.*)(\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\$)$",'$1'
		}
		if(Test-Host -ComputerName $ComputerName -TCPPort 135) {
			try {
				$CPUS = Get-WmiObject Win32_Processor -ComputerName $ComputerName -ea STOP
				foreach($CPU in $CPUs) {
					$myobj = @{
						ComputerName = $ComputerName
						Name = $CPU.Name
						Manufacturer = $CPU.Manufacturer
						Speed = $CPU.MaxClockSpeed
						Cores = $CPU.NumberOfCores
						L2Cache = $CPU.L2CacheSize
						Stepping = $CPU.Stepping
					}
					$obj = New-Object PSObject -Property $myobj
					$obj.PSTypeNames.Clear()
					$obj.PSTypeNames.Add('BSonPosh.Computer.Processor')
					$obj
				}
			} catch {
				Write-Host "Host [$ComputerName] Failed with Error: $($lastError[0])" -ForegroundColor Red
			}
		} else {
			Write-Host "Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener dirección IP y configuración de red
#------------------------------------------------------------------
function Get-IP {
	[Cmdletbinding()]
	Param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	Process {
		$NICs = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $ComputerName
		foreach($Nic in $NICs) {
			$myobj = @{
				Name = $Nic.Description
				MacAddress = $Nic.MACAddress
				IP4 = $Nic.IPAddress | Where-Object { $_ -match "\d+\.\d+\.\d+\.\d+" }
				IP6 = $Nic.IPAddress | Where-Object { $_ -match "\:\:" }
				IP4Subnet = $Nic.IPSubnet | Where-Object { $_ -match "\d+\.\d+\.\d+\.\d+" }
				DefaultGWY = $Nic.DefaultIPGateway | Select-Object -First 1
				DNSServer = $Nic.DNSServerSearchOrder
				WINSPrimary = $Nic.WINSPrimaryServer
				WINSSecondary = $Nic.WINSSecondaryServer
			}
			$obj = New-Object PSObject -Property $myobj
			$obj.PSTypeNames.Clear()
			$obj.PSTypeNames.Add('BSonPosh.IPInfo')
			$obj
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener software instalado desde el registro
#------------------------------------------------------------------
function Get-InstalledSoftware {
	[Cmdletbinding()]
	Param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	Process {
		if ($ComputerName -match "(.*)(\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\$)$", '$1'
		}
		$HKLM = 2147483650 # 0x80000002
		$uninstallPaths = @(
			"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
			"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
		)
		try {
			$reg = [wmiclass]"\\$ComputerName\root\default:StdRegProv"
			$seen = @{}
			foreach ($regPath in $uninstallPaths) {
				$enumResult = $reg.EnumKey($HKLM, $regPath)
				if ($enumResult.ReturnValue -ne 0 -or $null -eq $enumResult.sNames) { continue }
				foreach ($key in $enumResult.sNames) {
					$subPath = "$regPath\$key"
					$dispName = ($reg.GetStringValue($HKLM, $subPath, "DisplayName")).sValue
					if ([string]::IsNullOrWhiteSpace($dispName)) { continue }
					if ($seen.ContainsKey($dispName)) { continue }
					$seen[$dispName] = $true
					[PSCustomObject]@{
						Name    = $dispName
						Version = ($reg.GetStringValue($HKLM, $subPath, "DisplayVersion")).sValue
						Vendor  = ($reg.GetStringValue($HKLM, $subPath, "Publisher")).sValue
					}
				}
			}
		} catch {
			Write-Verbose "[Get-InstalledSoftware] Error en $ComputerName : $_"
			throw
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Resolver nombre de Hive del registro a objeto .NET
#------------------------------------------------------------------
function Get-RegistryHive {
	param($HiveName)
	Switch -regex ($HiveName) {
		"^(HKCR|ClassesRoot|HKEY_CLASSES_ROOT)$" { [Microsoft.Win32.RegistryHive]"ClassesRoot"; continue }
		"^(HKCU|CurrentUser|HKEY_CURRENT_USER)$" { [Microsoft.Win32.RegistryHive]"CurrentUser"; continue }
		"^(HKLM|LocalMachine|HKEY_LOCAL_MACHINE)$" { [Microsoft.Win32.RegistryHive]"LocalMachine"; continue }
		"^(HKU|Users|HKEY_USERS)$" { [Microsoft.Win32.RegistryHive]"Users"; continue }
		"^(HKCC|CurrentConfig|HKEY_CURRENT_CONFIG)$" { [Microsoft.Win32.RegistryHive]"CurrentConfig"; continue }
		"^(HKPD|PerformanceData|HKEY_PERFORMANCE_DATA)$" { [Microsoft.Win32.RegistryHive]"PerformanceData"; continue }
		Default { 1; continue }
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener una clave del registro (local o remoto)
#------------------------------------------------------------------
function Get-RegistryKey {
	[Cmdletbinding()]
	Param(
		[Parameter(mandatory=$true)]
		[string]$Path,
		[Alias("Server")]
		[Parameter(ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:ComputerName,
		[switch]$Recurse,
		[Alias("RW")]
		[switch]$ReadWrite
	)
	Begin {
		$PathParts = $Path -split "\\|/",0,"RegexMatch"
		$Hive = $PathParts[0]
		$KeyPath = $PathParts[1..$PathParts.count] -join "\\"
	}
	Process {
		$RegHive = Get-RegistryHive $Hive
		if($RegHive -eq 1) {
			Write-Host "Invalid Path: $Path, Registry Hive [$Hive] is invalid!" -ForegroundColor Red
		} else {
			$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive,$ComputerName)
			try {
				$Key = $BaseKey.OpenSubKey($KeyPath, $ReadWrite)
				if ($Key) {
					$Key = $Key | Add-Member -NotePropertyName ComputerName -NotePropertyValue $ComputerName -PassThru
					$Key = $Key | Add-Member -NotePropertyName Hive -NotePropertyValue $RegHive -PassThru
					$Key = $Key | Add-Member -NotePropertyName Path -NotePropertyValue $KeyPath -PassThru
					$Key.PSTypeNames.Clear()
					$Key.PSTypeNames.Add('BSonPosh.Registry.Key')
					$Key
				}
			} catch {
				Write-Verbose "[Get-RegistryKey] :: ERROR :: Unable to Open Key:$KeyPath on $ComputerName"
			}
			if ($Recurse -and $Key) {
				$SubKeyNames = $Key.GetSubKeyNames()
				foreach($Name in $SubKeyNames) {
					try {
						$SubKey = $Key.OpenSubKey($Name)
						if($SubKey.GetSubKeyNames()) {
							Get-RegistryKey -ComputerName $ComputerName -Path $SubKey.Name -Recurse
						} else {
							Get-RegistryKey -ComputerName $ComputerName -Path $SubKey.Name
						}
					} catch {
						Write-Verbose "[Get-RegistryKey] :: ERROR :: Unable to Open SubKey:$Name in $($Key.Name)"
					}
				}
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener un valor del registro
#------------------------------------------------------------------
function Get-RegistryValue {
	[Cmdletbinding()]
	Param(
		[Parameter(mandatory=$true)]
		[string]$Path,
		[Parameter()]
		[string]$Name,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:ComputerName,
		[Parameter()]
		[switch]$Recurse,
		[Parameter()]
		[switch]$Default
	)
	Process {
		if ($Recurse) {
			$Keys = Get-RegistryKey -Path $Path -ComputerName $ComputerName -Recurse
			foreach ($Key in $Keys) {
				if ($Name) {
					try {
						$myobj = @{
							ComputerName = $ComputerName
							Name = $Name
							Value = $Key.GetValue($Name)
							Type = $Key.GetValueKind($Name)
							Path = $Key
						}
						$obj = New-Object PSCustomObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
						$obj
					} catch {
						Write-Verbose "[Get-RegistryValue] :: ERROR :: Unable to Get Value for:$Name in $($Key.Name)"
					}
				} elseif ($Default) {
					try {
						$myobj = @{
							ComputerName = $ComputerName
							Name = "(Default)"
							Value = if ($Key.GetValue("")) { $Key.GetValue("") } else { "EMPTY" }
							Type = if ($Key.GetValue("")) { $Key.GetValueKind("") } else { "N/A" }
							Path = $Key
						}
						$obj = New-Object PSCustomObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
						$obj
					} catch {
						Write-Verbose "[Get-RegistryValue] :: ERROR :: Unable to Get Value for:(Default) in $($Key.Name)"
					}
				} else {
					foreach ($ValueName in $Key.GetValueNames()) {
						try {
							$myobj = @{
								ComputerName = $ComputerName
								Name = if ($ValueName -match "^$") { "(Default)" } else { $ValueName }
								Value = $Key.GetValue($ValueName)
								Type = $Key.GetValueKind($ValueName)
								Path = $Key
							}
							$obj = New-Object PSCustomObject -Property $myobj
							$obj.PSTypeNames.Clear()
							$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
							$obj
						} catch {
							Write-Verbose "[Get-RegistryValue] :: ERROR :: Unable to Get Value for:$ValueName in $($Key.Name)"
						}
					}
				}
			}
		} else {
			$Key = Get-RegistryKey -Path $Path -ComputerName $ComputerName
			if ($Name) {
				try {
					$myobj = @{
						ComputerName = $ComputerName
						Name = $Name
						Value = $Key.GetValue($Name)
						Type = $Key.GetValueKind($Name)
						Path = $Key
					}
					$obj = New-Object PSCustomObject -Property $myobj
					$obj.PSTypeNames.Clear()
					$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
					$obj
				} catch {
					Write-Verbose "[Get-RegistryValue] :: ERROR :: Unable to Get Value for:$Name in $($Key.Name)"
				}
			} elseif ($Default) {
				try {
					$myobj = @{
						ComputerName = $ComputerName
						Name = "(Default)"
						Value = if ($Key.GetValue("")) { $Key.GetValue("") } else { "EMPTY" }
						Type = if ($Key.GetValue("")) { $Key.GetValueKind("") } else { "N/A" }
						Path = $Key
					}
					$obj = New-Object PSCustomObject -Property $myobj
					$obj.PSTypeNames.Clear()
					$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
					$obj
				} catch {
					Write-Verbose "[Get-RegistryValue] :: ERROR :: Unable to Get Value for:(Default) in $($Key.Name)"
				}
			} else {
				foreach ($ValueName in $Key.GetValueNames()) {
					try {
						$myobj = @{
							ComputerName = $ComputerName
							Name = if ($ValueName -match "^$") { "(Default)" } else { $ValueName }
							Value = $Key.GetValue($ValueName)
							Type = $Key.GetValueKind($ValueName)
							Path = $Key
						}
						$obj = New-Object PSCustomObject -Property $myobj
						$obj.PSTypeNames.Clear()
						$obj.PSTypeNames.Add('BSonPosh.Registry.Value')
						$obj
					} catch {
						Write-Verbose "[Get-RegistryValue] :: ERROR :: Unable to Get Value for:$ValueName in $($Key.Name)"
					}
				}
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Crear nueva clave en el registro
#------------------------------------------------------------------
function New-RegistryKey {
	[Cmdletbinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(mandatory=$true)]
		[string]$Path,
		[Parameter(mandatory=$true)]
		[string]$Name,
		[Alias("Server")]
		[Parameter(ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:ComputerName
	)
	Begin {
		$ReadWrite = [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree
		$PathParts = $Path -split "\\|/",0,"RegexMatch"
		$Hive = $PathParts[0]
		$KeyPath = $PathParts[1..$PathParts.count] -join "\\"
	}
	Process {
		$RegHive = Get-RegistryHive $Hive
		if ($RegHive -eq 1) {
			Write-Host "Invalid Path: $Path, Registry Hive [$Hive] is invalid!" -ForegroundColor Red
		} else {
			$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive,$ComputerName)
			$Key = $BaseKey.OpenSubKey($KeyPath,$True)
			if ($PSCmdlet.ShouldProcess($ComputerName,"Creating Key [$Name] under $Path")) {
				$Key.CreateSubKey($Name,$ReadWrite)
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Crear valor de clave en el registro
#------------------------------------------------------------------
function New-RegistryValue {
	[Cmdletbinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(mandatory=$true)]
		[string]$Path,
		[Parameter(mandatory=$true)]
		[string]$Name,
		[Parameter()]
		[string]$Value,
		[Parameter()]
		[string]$Type,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:ComputerName
	)
	Begin {
		Switch ($Type) {
			"Unknown"      { $ValueType = [Microsoft.Win32.RegistryValueKind]::Unknown; continue }
			"String"       { $ValueType = [Microsoft.Win32.RegistryValueKind]::String; continue }
			"ExpandString" { $ValueType = [Microsoft.Win32.RegistryValueKind]::ExpandString; continue }
			"Binary"       { $ValueType = [Microsoft.Win32.RegistryValueKind]::Binary; continue }
			"DWord"        { $ValueType = [Microsoft.Win32.RegistryValueKind]::DWord; continue }
			"MultiString"  { $ValueType = [Microsoft.Win32.RegistryValueKind]::MultiString; continue }
			"QWord"        { $ValueType = [Microsoft.Win32.RegistryValueKind]::QWord; continue }
			Default         { $ValueType = [Microsoft.Win32.RegistryValueKind]::String; continue }
		}
	}
	Process {
		if (Test-RegistryValue -Path $Path -Name $Name -ComputerName $ComputerName) {
			"Registry value already exist"
		} else {
			$Key = Get-RegistryKey -Path $Path -ComputerName $ComputerName -ReadWrite
			if ($PSCmdlet.ShouldProcess($ComputerName, "Creating Value [$Name] under $Path with value [$Value]")) {
				if ($Value) {
					$Key.SetValue($Name, $Value, $ValueType)
				} else {
					$Key.SetValue($Name, $ValueType)
				}
				Get-RegistryValue -Path $Path -Name $Name -ComputerName $ComputerName
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Eliminar clave de registro
#------------------------------------------------------------------
function Remove-RegistryKey {
	[Cmdletbinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(mandatory=$true)]
		[string]$Path,
		[Parameter(mandatory=$true)]
		[string]$Name,
		[Alias("Server")]
		[Parameter(ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:ComputerName,
		[Parameter()]
		[switch]$Recurse
	)
	Begin {
		$PathParts = $Path -split "\\|/",0,"RegexMatch"
		$Hive = $PathParts[0]
		$KeyPath = $PathParts[1..$PathParts.count] -join "\\"
	}
	Process {
		if (Test-RegistryKey -Path "$Path\$Name" -ComputerName $ComputerName) {
			$RegHive = Get-RegistryHive $Hive
			if ($RegHive -eq 1) {
				Write-Host "Invalid Path: $Path, Registry Hive [$Hive] is invalid!" -ForegroundColor Red
			} else {
				$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
				$Key = $BaseKey.OpenSubKey($KeyPath, $True)
				if ($PSCmdlet.ShouldProcess($ComputerName, "Deleting Key [$Name]")) {
					if ($Recurse) {
						$Key.DeleteSubKeyTree($Name)
					} else {
						$Key.DeleteSubKey($Name)
					}
				}
			}
		} else {
			"Key [$Path\$Name] does not exist"
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Eliminar valor de registro
#------------------------------------------------------------------
function Remove-RegistryValue {
	[Cmdletbinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(mandatory=$true)]
		[string]$Path,
		[Parameter(mandatory=$true)]
		[string]$Name,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:ComputerName
	)
	Process {
		if (Test-RegistryValue -Path $Path -Name $Name -ComputerName $ComputerName) {
			$Key = Get-RegistryKey -Path $Path -ComputerName $ComputerName -ReadWrite
			if ($PSCmdlet.ShouldProcess($ComputerName, "Deleting Value [$Name] under $Path")) {
				$Key.DeleteValue($Name)
			}
		} else {
			"Registry Value is already gone"
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Buscar claves y valores en el registro
#------------------------------------------------------------------
function Search-Registry {
	[Cmdletbinding(DefaultParameterSetName="ByFilter")]
	Param(
		[Parameter(ParameterSetName="ByFilter",Position=0)]
		[string]$Filter= ".*",
		[Parameter(ParameterSetName="ByName",Position=0)]
		[string]$Name,
		[Parameter(ParameterSetName="ByValue",Position=0)]
		[string]$Value,
		[Parameter()]
		[string]$Path,
		[Parameter()]
		[string]$Hive = "LocalMachine",
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME,
		[Parameter()]
		[switch]$KeyOnly
	)
	Begin {
		$RegHive = Get-RegistryHive $Hive
	}
	Process {
		switch ($PSCmdlet.ParameterSetName) {
			"ByFilter" {
				if ($Path -and (Test-RegistryKey "$RegHive\$Path")) {
					$Keys = Get-RegistryKey -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse
				} else {
					$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
					$Keys = foreach ($SubKeyName in $BaseKey.GetSubKeyNames()) {
						try {
							$SubKey = $BaseKey.OpenSubKey($SubKeyName, $true)
							Get-RegistryKey -Path $SubKey.Name -ComputerName $ComputerName -Recurse
						} catch {}
					}
				}
				if ($KeyOnly) {
					$Keys | Where-Object { $_.Name -match "$Filter" }
				} else {
					$Keys | Where-Object { $_.Name -match "$Filter" }
					Get-RegistryValue -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | Where-Object { $_.Name -match "$Filter" }
				}
			}
			"ByName" {
				$NameFilter = "^.*\\{0}$" -f $Name
				if ($Path -and (Test-RegistryKey "$RegHive\$Path")) {
					$Keys = Get-RegistryKey -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse
				} else {
					$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
					$Keys = foreach ($SubKeyName in $BaseKey.GetSubKeyNames()) {
						try {
							$SubKey = $BaseKey.OpenSubKey($SubKeyName, $true)
							Get-RegistryKey -Path $SubKey.Name -ComputerName $ComputerName -Recurse
						} catch {}
					}
				}
				if ($KeyOnly) {
					$Keys | Where-Object { $_.Name -match $NameFilter }
				} else {
					$Keys | Where-Object { $_.Name -match $NameFilter }
					Get-RegistryValue -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | Where-Object { $_.Name -eq $Name }
				}
			}
			"ByValue" {
				if ($Path -and (Test-RegistryKey "$RegHive\$Path")) {
					Get-RegistryValue -Path "$RegHive\$Path" -ComputerName $ComputerName -Recurse | Where-Object { $_.Value -eq $Value }
				} else {
					$BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
					foreach ($SubKeyName in $BaseKey.GetSubKeyNames()) {
						try {
							$SubKey = $BaseKey.OpenSubKey($SubKeyName, $true)
							Get-RegistryValue -Path $SubKey.Name -ComputerName $ComputerName -Recurse | Where-Object { $_.Value -eq $Value }
						} catch {}
					}
				}
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Modificar valor en el registro
#------------------------------------------------------------------
function Set-RegistryValue {
	[Cmdletbinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(mandatory=$true)]
		[string]$Path,
		[Parameter(mandatory=$true)]
		[string]$Name,
		[Parameter()]
		[string]$Value,
		[Parameter()]
		[string]$Type,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:ComputerName
	)
	Begin {
		Switch ($Type) {
			"Unknown"      { $ValueType = [Microsoft.Win32.RegistryValueKind]::Unknown; continue }
			"String"       { $ValueType = [Microsoft.Win32.RegistryValueKind]::String; continue }
			"ExpandString" { $ValueType = [Microsoft.Win32.RegistryValueKind]::ExpandString; continue }
			"Binary"       { $ValueType = [Microsoft.Win32.RegistryValueKind]::Binary; continue }
			"DWord"        { $ValueType = [Microsoft.Win32.RegistryValueKind]::DWord; continue }
			"MultiString"  { $ValueType = [Microsoft.Win32.RegistryValueKind]::MultiString; continue }
			"QWord"        { $ValueType = [Microsoft.Win32.RegistryValueKind]::QWord; continue }
			Default         { $ValueType = [Microsoft.Win32.RegistryValueKind]::String; continue }
		}
	}
	Process {
		$Key = Get-RegistryKey -Path $Path -ComputerName $ComputerName -ReadWrite
		if ($PSCmdlet.ShouldProcess($ComputerName, "Creating Value [$Name] under $Path with value [$Value]")) {
			if ($Value) {
				$Key.SetValue($Name, $Value, $ValueType)
			} else {
				$Key.SetValue($Name, $ValueType)
			}
			Get-RegistryValue -Path $Path -Name $Name -ComputerName $ComputerName
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Comprobar existencia de clave y valor de registro
#------------------------------------------------------------------
function Test-RegistryKey {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]$Path,

        [Alias("dnsHostName")]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [string]$ComputerName = $Env:COMPUTERNAME
    )
    Process {
        $PathParts = $Path -split "\\|/", 0, "RegexMatch"
        $Hive = $PathParts[0]
        if ($PathParts.Count -gt 1) {
            $KeyPath = $PathParts[1..($PathParts.Count - 1)] -join "\"
        }
        else {
            $KeyPath = ""
        }

        $RegHive = Get-RegistryHive $Hive
        if ($RegHive -eq 1) {
            return $false
        }
        try {
            $BaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegHive, $ComputerName)
            $Key = $BaseKey.OpenSubKey($KeyPath)
            return [bool]$Key
        }
        catch {
            return $false
        }
    }
}
function Test-RegistryValue {
	[Cmdletbinding()]
	Param(
		[Parameter(mandatory=$true)]
		[string]$Path,
		[Parameter(mandatory=$true)]
		[string]$Name,
		[Parameter()]
		[string]$Value,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	Process {
		$Key = Get-RegistryKey -Path $Path -ComputerName $ComputerName
		if ($null -eq $Key) { return $false }
		try {
			if ($Value) {
				return ($Key.GetValue($Name) -eq $Value)
			} else {
				return ($null -ne $Key.GetValue($Name))
			}
		} catch {
			return $false
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener relaciones de disco, puntos de montaje y unidades mapeadas
#------------------------------------------------------------------
function Get-DiskRelationship {
	param ([string]$computername = "localhost")
	Get-WmiObject -Class Win32_DiskDrive -ComputerName $computername | ForEach-Object {
		"`n $($_.Name) $($_.Model)"
		$query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" + $_.DeviceID + "'} WHERE ResultClass=Win32_DiskPartition"
		Get-WmiObject -Query $query -ComputerName $computername | ForEach-Object {
			"Name             : $($_.Name)"
			"Description      : $($_.Description)"
			"PrimaryPartition : $($_.PrimaryPartition)"
			$query2 = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" + $_.DeviceID + "'} WHERE ResultClass=Win32_LogicalDisk"
			Get-WmiObject -Query $query2 -ComputerName $computername | Format-List Name,
			@{Name="Disk Size (GB)"; Expression={"{0:F3}" -f ($_.Size/1GB)}},
			@{Name="Free Space (GB)"; Expression={"{0:F3}" -f ($_.FreeSpace/1GB)}}
		}
	}
}
function Get-MountPoint {
	param ([string]$computername = "localhost")
	Get-WmiObject -Class Win32_MountPoint -ComputerName $computername |
	Where-Object {$_.Directory -like 'Win32_Directory.Name="*"'} |
	ForEach-Object {
		$vol = $_.Volume
		Get-WmiObject -Class Win32_Volume -ComputerName $computername | Where-Object {$_.__RELPATH -eq $vol} |
		Select-Object @{Name="Folder"; Expression={$_.Caption}},
		@{Name="Size (GB)"; Expression={"{0:F3}" -f ($_.Capacity / 1GB)}},
		@{Name="Free (GB)"; Expression={"{0:F3}" -f ($_.FreeSpace / 1GB)}},
		@{Name="%Free"; Expression={"{0:F2}" -f (($_.FreeSpace/$_.Capacity)*100)}} |
		Format-Table -AutoSize
	}
}
function Get-MappedDrive {
	param ([string]$computername = "localhost")
	Get-WmiObject -Class Win32_MappedLogicalDisk -ComputerName $computername |
	Format-List DeviceId, VolumeName, SessionID, Size, FreeSpace, ProviderName
}
#------------------------------------------------------------------
# SUBBLOQUE: Comprobar conectividad a host (ping o puerto TCP)
#------------------------------------------------------------------
function Test-Host {
	[CmdletBinding()]
	param(
		[Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, Mandatory=$true)]
		[string]$ComputerName,
		[Parameter()]
		[int]$TCPPort = 80,
		[Parameter()]
		[int]$Timeout = 3000,
		[Parameter()]
		[string]$Property
	)
	begin {
		function PingServer {
			param($MyHost)
			try {
				(Get-WmiObject win32_pingstatus -f "address='$MyHost'").StatusCode -eq 0
			} catch {
				$false
			}
		}
	}
	process {
		if ($ComputerName -match "(.*)(\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\$)$",'$1'
		}
		if ($TCPPort) {
			if ($Property) {
				if (Test-Port $_.$Property -TCP $TCPPort -Timeout $Timeout) { if ($_){$_} else {$ComputerName} }
			} else {
				if (Test-Port $ComputerName -TCP $TCPPort -Timeout $Timeout) { if ($_){$_} else {$ComputerName} }
			}
		} else {
			if ($Property) {
				if (PingServer $_.$Property) { if ($_){$_} else {$ComputerName} }
			} else {
				if (PingServer $ComputerName) { $ComputerName }
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Comprobar apertura de puerto en host
#------------------------------------------------------------------
function Test-Port {
	[Cmdletbinding()]
	param(
		[Parameter()]
		[int]$TCPPort = 135,
		[Parameter()]
		[int]$Timeout = 3000,
		[Alias("dnsHostName")]
		[Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
		[string]$ComputerName = $env:COMPUTERNAME
	)
	process {
		try {
			$tcpclient = New-Object System.Net.Sockets.TcpClient
			$iar = $tcpclient.BeginConnect($ComputerName, $TCPPort, $null, $null)
			$wait = $iar.AsyncWaitHandle.WaitOne($Timeout, $false)
			if (!$wait) {
				$tcpclient.Close()
				return $false
			}
			$tcpclient.EndConnect($iar) | Out-Null
			$tcpclient.Close()
			return $true
		} catch {
			return $false
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener configuración de módulos de memoria
#------------------------------------------------------------------
function Get-MemoryConfiguration {
	[Cmdletbinding()]
	param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	process {
		if ($ComputerName -match "(.*)(\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\$)$",'$1'
		}
		if (Test-Host $ComputerName -TCPPort 135) {
			try {
				Get-WmiObject Win32_PhysicalMemory -ComputerName $ComputerName -ea Stop | ForEach-Object {
					[PSCustomObject]@{
						ComputerName = $ComputerName
						Description  = $_.Tag
						Slot         = $_.DeviceLocator
						Speed        = $_.Speed
						SizeGB       = $_.Capacity / 1GB
					} | Add-Member -MemberType NoteProperty -Name PSTypeName -Value 'BSonPosh.MemoryConfiguration' -PassThru
				}
			} catch {
				Write-Host " Host [$ComputerName] Failed with Error: $($_.Exception.Message)" -ForegroundColor Red
			}
		} else {
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener información de estado de conexiones de red
#------------------------------------------------------------------
function Get-NetStat {
	[Cmdletbinding(DefaultParameterSetName = "All")]
	param(
		[Parameter()]
		[string]$ProcessName,

		[Parameter()]
		[ValidateSet("LISTENING", "ESTABLISHED", "CLOSE_WAIT", "TIME_WAIT")]
		[string]$State,

		[Parameter(ParameterSetName = "Interval")]
		[int]$Interval,

		[Parameter()]
		[int]$Sleep = 1,

		[Parameter(ParameterSetName = "Loop")]
		[switch]$Loop
	)
	# Subfunción para parsear resultados de netstat
	function Convert-Netstat {
		param($NetStat)

		[regex]$RegEx = '\s+(?<Protocol>\S+)\s+(?<LocalAddress>\S+)\s+(?<RemoteAddress>\S+)\s+(?<State>\S+)\s+(?<PID>\S+)'

		foreach ($line in $NetStat) {
			if ($line -match $RegEx) {
				$procName = (Get-Process -Id $matches.PID -ErrorAction SilentlyContinue).Name
				$connection = [PSCustomObject]@{
					Protocol      = $matches.Protocol
					LocalAddress  = ($matches.LocalAddress -split ":")[0]
					LocalPort     = ($matches.LocalAddress -split ":")[1]
					RemoteAddress = ($matches.RemoteAddress -split ":")[0]
					RemotePort    = ($matches.RemoteAddress -split ":")[1]
					State         = $matches.State
					ProcessID     = $matches.PID
					ProcessName   = $procName
				}
				$connection.PSTypeNames.Insert(0, 'BSonPosh.NetStatInfo')

				if ($ProcessName) {
					if ($connection.ProcessName -eq $ProcessName) { $connection }
				} elseif ($State) {
					if ($connection.State -eq $State) { $connection }
				} else {
					$connection
				}
			}
		}
	}

	switch ($PSCmdlet.ParameterSetName) {
		"All" {
			$results = netstat -ano | Where-Object { $_ -match "^(TCP|UDP)\s+\d" }
			Convert-Netstat $results
		}
		"Interval" {
			for ($i = 1; $i -le $Interval; $i++) {
				Start-Sleep -Seconds $Sleep
				$results = netstat -ano | Where-Object { $_ -match "^(TCP|UDP)\s+\d" }
				Convert-Netstat $results | Out-String
			}
		}
		"Loop" {
			Write-Host "`nProtocol LocalAddress  LocalPort RemoteAddress  RemotePort State       ProcessName   PID"
			Write-Host "-------- ------------  --------- -------------  ---------- -----       -----------   ---" -ForegroundColor White

			$oldPos = $Host.UI.RawUI.CursorPosition
			[console]::TreatControlCAsInput = $true
			$Connections = @{}

			while ($true) {
				$results = netstat -ano | Where-Object { $_ -match "^(TCP|UDP)\s+\d" }
				$parsed = Convert-Netstat $results

				foreach ($conn in $parsed) {
					$key = $conn.LocalPort
					$msg = "{0,-9}{1,-14}{2,-10}{3,-15}{4,-11}{5,-12}{6,-14}{7,-10}" -f $conn.Protocol, $conn.LocalAddress, $conn.LocalPort, $conn.RemoteAddress, $conn.RemotePort, $conn.State, $conn.ProcessName, $conn.ProcessID

					if ($Connections[$key] -eq $conn.ProcessID) {
						Write-Host $msg
					} else {
						$Connections[$key] = $conn.ProcessID
						Write-Host $msg -ForegroundColor Yellow
					}
				}

				if ($Host.UI.RawUI.KeyAvailable -and (3 -eq [int]$Host.UI.RawUI.ReadKey("AllowCtrlC,IncludeKeyUp,NoEcho").Character)) {
					Write-Host "`nExiting..." -ForegroundColor Yellow
					[console]::TreatControlCAsInput = $false
					break
				}
				$Host.UI.RawUI.CursorPosition = $oldPos
				Start-Sleep -Seconds $Sleep
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener información de interfaces de red
#------------------------------------------------------------------
function Get-NICInfo {
	[Cmdletbinding()]
	param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	process {
		if ($ComputerName -match "(.*)(\\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\\$)$",'$1'
		}
		if (Test-Host -ComputerName $ComputerName -TCPPort 135) {
			try {
				$NICS = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName
				foreach ($NIC in $NICS) {
					$Query = "Select Name,NetConnectionID FROM Win32_NetworkAdapter WHERE Index='$($NIC.Index)'"
					$NetConnectionID = Get-WmiObject -Query $Query -ComputerName $ComputerName
					[PSCustomObject]@{
						ComputerName = $ComputerName
						Name         = $NetConnectionID.Name
						NetID        = $NetConnectionID.NetConnectionID
						MacAddress   = $NIC.MacAddress
						IP           = $NIC.IPAddress | Where-Object { $_ -match "\d+\.\d+\.\d+\.\d+" }
						Subnet       = $NIC.IPSubnet  | Where-Object { $_ -match "\d+\.\d+\.\d+\.\d+" }
						Enabled      = $NIC.IPEnabled
						Index        = $NIC.Index
					} | Add-Member -MemberType NoteProperty -Name PSTypeName -Value 'BSonPosh.NICInfo' -PassThru
				}
			} catch {
				Write-Host " Host [$ComputerName] Failed with Error: $($_.Exception.Message)" -ForegroundColor Red
			}
		} else {
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener información de placa base
#------------------------------------------------------------------
function Get-MotherBoard {
	[Cmdletbinding()]
	param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	process {
		if ($ComputerName -match "(.*)(\\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\\$)$",'$1'
		}
		if (Test-Host -ComputerName $ComputerName -TCPPort 135) {
			try {
				$MBInfo = Get-WmiObject Win32_BaseBoard -ComputerName $ComputerName -ea Stop
				[PSCustomObject]@{
					ComputerName = $ComputerName
					Name         = $MBInfo.Product
					Manufacturer = $MBInfo.Manufacturer
					Version      = $MBInfo.Version
					SerialNumber = $MBInfo.SerialNumber
				} | Add-Member -MemberType NoteProperty -Name PSTypeName -Value 'BSonPosh.Computer.MotherBoard' -PassThru
			} catch {
				Write-Host " Host [$ComputerName] Failed with Error: $($_.Exception.Message)" -ForegroundColor Red
			}
		} else {
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener tabla de rutas IPv4
#------------------------------------------------------------------
function Get-RouteTable {
	[Cmdletbinding()]
	param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)
	process {
		if ($ComputerName -match "(.*)(\\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\\$)$",'$1'
		}
		if (Test-Host $ComputerName -TCPPort 135) {
			try {
				Get-WmiObject Win32_IP4RouteTable -ComputerName $ComputerName -Property Name,Mask,NextHop,Metric1,Type | ForEach-Object {
					[PSCustomObject]@{
						ComputerName = $ComputerName
						Name         = $_.Name
						NetworkMask  = $_.Mask
						Gateway      = if ($_.NextHop -eq "0.0.0.0") {"On-Link"} else { $_.NextHop }
						Metric       = $_.Metric1
					} | Add-Member -MemberType NoteProperty -Name PSTypeName -Value 'BSonPosh.RouteTable' -PassThru
				}
			} catch {
				Write-Host " Host [$ComputerName] Failed with Error: $($_.Exception.Message)" -ForegroundColor Red
			}
		} else {
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener tipo de sistema (chasis, modelo, fabricante)
#------------------------------------------------------------------
function Get-SystemType {
	[Cmdletbinding()]
	param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
		[string]$ComputerName = $Env:COMPUTERNAME
	)

	begin {
		function ConvertTo-ChassisType ($Type) {
			switch ($Type) {
				1 {"Other"}; 2 {"Unknown"}; 3 {"Desktop"}; 4 {"Low Profile Desktop"};
				5 {"Pizza Box"}; 6 {"Mini Tower"}; 7 {"Tower"}; 8 {"Portable"};
				9 {"Laptop"}; 10 {"Notebook"}; 11 {"Hand Held"}; 12 {"Docking Station"};
				13 {"All in One"}; 14 {"Sub Notebook"}; 15 {"Space-Saving"};
				16 {"Lunch Box"}; 17 {"Main System Chassis"}; 18 {"Expansion Chassis"};
				19 {"SubChassis"}; 20 {"Bus Expansion Chassis"};
				21 {"Peripheral Chassis"}; 22 {"Storage Chassis"};
				23 {"Rack Mount Chassis"}; 24 {"Sealed-Case PC"}
			}
		}

		function ConvertTo-SecurityStatus ($Status) {
			switch ($Status) {
				1 {"Other"}; 2 {"Unknown"}; 3 {"None"};
				4 {"External Interface Locked Out"}; 5 {"External Interface Enabled"}
			}
		}
	}

	process {
		if ($ComputerName -match "(.*)(\\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\\$)$", '$1'
		}
		if (Test-Host $ComputerName -TCPPort 135) {
			try {
				$SystemInfo = Get-WmiObject Win32_SystemEnclosure -ComputerName $ComputerName
				$CSInfo = Get-WmiObject -Query "Select Model, Domain FROM Win32_ComputerSystem" -ComputerName $ComputerName
				
				# Obtener información de logon server
				$LogonServer = $null
				try {
					if ($ComputerName -eq $Env:COMPUTERNAME -or $ComputerName -eq "localhost" -or $ComputerName -eq ".") {
						# Para equipo local usar variables de entorno
						$LogonServer = $env:LOGONSERVER
					} else {
						# Para equipos remotos, usar "No disponible" o intentar obtener del dominio
						$LogonServer = "No disponible (remoto)"
					}
				} catch {
					$LogonServer = "No disponible"
				}

				[PSCustomObject]@{
					ComputerName    = $ComputerName
					Manufacturer    = $SystemInfo.Manufacturer
					Model           = $CSInfo.Model
					SerialNumber    = $SystemInfo.SerialNumber
					SecurityStatus  = ConvertTo-SecurityStatus $SystemInfo.SecurityStatus
					Type            = ConvertTo-ChassisType $SystemInfo.ChassisTypes
					Domain          = $CSInfo.Domain
					LogonServer     = $LogonServer
				} | Add-Member -MemberType NoteProperty -Name PSTypeName -Value 'BSonPosh.SystemType' -PassThru
			} catch {
				Write-Verbose "[Get-SystemType] :: [$ComputerName] Failed with Error: $($_.Exception.Message)"
			}
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener fecha del último reinicio del sistema
#------------------------------------------------------------------
function Get-RebootTime {
	[cmdletbinding()]
	param(
		[alias('dnsHostName')]
		[Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
		[string]$ComputerName = $Env:COMPUTERNAME,
		[Parameter()]
		[Switch]$Last
	)
	process {
		if ($ComputerName -match "(.*)(\\$)$") {
			$ComputerName = $ComputerName -replace "(.*)(\\$)$", '$1'
		}
		if (Test-Host $ComputerName -TCPPort 135) {
			try {
				if ($Last) {
					$date = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ea Stop | ForEach-Object { $_.LastBootUpTime }
					$RebootTime = [DateTime]::ParseExact($date.Split('.')[0], 'yyyyMMddHHmmss', $null)
					[PSCustomObject]@{
						ComputerName = $ComputerName
						RebootTime   = $RebootTime
					} | Add-Member -MemberType NoteProperty -Name PSTypeName -Value 'BSonPosh.RebootTime' -PassThru
				} else {
					$Query = "Select * FROM Win32_NTLogEvent WHERE SourceName='eventlog' AND EventCode='6009'"
					Get-WmiObject -Query $Query -ComputerName $ComputerName -ea 0 | ForEach-Object {
						$RebootTime = [DateTime]::ParseExact($_.TimeGenerated.Split('.')[0], 'yyyyMMddHHmmss', $null)
						[PSCustomObject]@{
							ComputerName = $ComputerName
							RebootTime   = $RebootTime
						} | Add-Member -MemberType NoteProperty -Name PSTypeName -Value 'BSonPosh.RebootTime' -PassThru
					}
				}
			} catch {
				Write-Host " Host [$ComputerName] Failed with Error: $($_.Exception.Message)" -ForegroundColor Red
			}
		} else {
			Write-Host " Host [$ComputerName] Failed Connectivity Test " -ForegroundColor Red
		}
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Convertir código de estado KMS a descripción
#------------------------------------------------------------------
function ConvertTo-KMSStatus {
	[cmdletbinding()]
	param(
		[Parameter(Mandatory = $true)]
		[int]$StatusCode
	)
	switch ($StatusCode) {
		0 { "Unlicensed" }
		1 { "Licensed" }
		2 { "OOBGrace" }
		3 { "OOTGrace" }
		4 { "NonGenuineGrace" }
		5 { "Notification" }
		6 { "ExtendedGrace" }
		default { "Unknown" }
	}
}
#------------------------------------------------------------------
# SUBBLOQUE: Obtener detalles de activación del servidor KMS
#------------------------------------------------------------------
function Get-KMSActivationDetail {
	[Cmdletbinding()]
	param(
		[Parameter(Mandatory = $true)]
		[string]$KMS,
		[Parameter()]
		[string]$Filter = "*",
		[Parameter()]
		[datetime]$After,
		[Parameter()]
		[switch]$Unique
	)

	$Events = if ($After) {
		Get-Eventlog -LogName "Key Management Service" -ComputerName $KMS -After $After -Message "*$Filter*"
	} else {
		Get-Eventlog -LogName "Key Management Service" -ComputerName $KMS -Message "*$Filter*"
	}

	$MyObjects = foreach ($Event in $Events) {
		$Message = $Event.Message.Split(",")
		[PSCustomObject]@{
			ComputerName = $Message[3]
			Date         = $Event.TimeGenerated
		} | Add-Member -MemberType NoteProperty -Name PSTypeName -Value 'BSonPosh.KMS.ActivationDetail' -PassThru
	}

	if ($Unique) {
		$MyObjects | Group-Object -Property ComputerName | ForEach-Object {
			[PSCustomObject]@{
				ComputerName = $_.Name
				Count        = $_.Count
			} | Add-Member -MemberType NoteProperty -Name PSTypeName -Value 'BSonPosh.KMS.ActivationDetail' -PassThru
		}
	} else {
		$MyObjects
	}
}
#==================================================================
# BLOQUE FUNCIONAL: Gestión de Servidor KMS (Key Management Service)
#==================================================================
#==================================================================
# SUBBLOQUE: Obtener información del servidor KMS
#==================================================================
function Get-KMSServer {
    <# ... documentación interna ... #>
    [Cmdletbinding()]
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [string]$KMS
    )

    if (!$KMS) {
        Write-Verbose " [Get-KMSServer] :: No KMS Server Passed... Using Discovery"
        $KMS = Test-KMSServerDiscovery | Select-Object -ExpandProperty ComputerName
    }

    try {
        Write-Verbose " [Get-KMSServer] :: Querying KMS Service using WMI"
        $KMSService = Get-WmiObject "SoftwareLicensingService" -ComputerName $KMS

        $myobj = @{
            ComputerName            = $KMS
            Version                 = $KMSService.Version
            KMSEnable               = $KMSService.KeyManagementServiceActivationDisabled -eq $false
            CurrentCount            = $KMSService.KeyManagementServiceCurrentCount
            Port                    = $KMSService.KeyManagementServicePort
            DNSPublishing           = $KMSService.KeyManagementServiceDnsPublishing
            TotalRequest            = $KMSService.KeyManagementServiceTotalRequests
            FailedRequest           = $KMSService.KeyManagementServiceFailedRequests
            Unlicensed              = $KMSService.KeyManagementServiceUnlicensedRequests
            Licensed                = $KMSService.KeyManagementServiceLicensedRequests
            InitialGracePeriod      = $KMSService.KeyManagementServiceOOBGraceRequests
            LicenseExpired          = $KMSService.KeyManagementServiceOOTGraceRequests
            NonGenuineGracePeriod   = $KMSService.KeyManagementServiceNonGenuineGraceRequests
            LicenseWithNotification = $KMSService.KeyManagementServiceNotificationRequests
            ActivationInterval      = $KMSService.VLActivationInterval
            RenewalInterval         = $KMSService.VLRenewalInterval
        }

        $obj = New-Object PSObject -Property $myobj
        $obj.PSTypeNames.Clear()
        $obj.PSTypeNames.Add('BSonPosh.KMS.Server')
        $obj
    } catch {
        Write-Verbose " [Get-KMSServer] :: Error: $($lastError[0])"
    }
}
#==================================================================
# SUBBLOQUE: Obtener estado de activación KMS de un equipo
#==================================================================
function Get-KMSStatus {
    [Cmdletbinding()]
    Param(
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [string]$ComputerName = $Env:COMPUTERNAME
    )

    Process {
        Write-Verbose " [Get-KMSStatus] :: Process Start"
        if ($ComputerName -match "(.*)(\$)$") {
            $ComputerName = $ComputerName -replace "(.*)(\$)$",'$1'
        }

        $Query = "Select * FROM SoftwareLicensingProduct WHERE Description LIKE '%VOLUME_KMSCLIENT%'"

        Write-Verbose " [Get-KMSStatus] :: ComputerName = $ComputerName"
        Write-Verbose " [Get-KMSStatus] :: Query = $Query"

        try {
            Write-Verbose " [Get-KMSStatus] :: Calling WMI"
            $WMIResult = Get-WmiObject -ComputerName $ComputerName -Query $Query

            foreach ($result in $WMIResult) {
                $myobj = @{
                    ComputerName  = $ComputerName
                    KMSServer     = $result.KeyManagementServiceMachine
                    KMSPort       = $result.KeyManagementServicePort
                    LicenseFamily = $result.LicenseFamily
                    Status        = ConvertTo-KMSStatus $result.LicenseStatus  # <-- Función no incluida aún
                }

                $obj = New-Object PSObject -Property $myobj
                $obj.PSTypeNames.Clear()
                $obj.PSTypeNames.Add('BSonPosh.KMS.Status')
                $obj
            }
        } catch {
            Write-Verbose " [Get-KMSStatus] :: Error - $($lastError[0])"
        }
    }
}
#==================================================================
# SUBBLOQUE: Testear si una máquina está activada vía KMS
#==================================================================
function Test-KMSIsActivated {
    [Cmdletbinding()]
    Param(
        [Alias('dnsHostName')]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [string]$ComputerName = $Env:COMPUTERNAME
    )

    Process {
        Write-Verbose " [Test-KMSActivation] :: Process start"
        if ($ComputerName -match "(.*)(\$)$") {
            $ComputerName = $ComputerName -replace "(.*)(\$)$",'$1'
        }

        if (Test-Host $ComputerName -TCP 135) {
            $status = Get-KMSStatus -ComputerName $ComputerName
            if ($status.Status -eq "Licensed") {
                $_
            }
        }
    }
}
#==================================================================
# SUBBLOQUE: Descubrimiento automático del servidor KMS por DNS
#==================================================================
function Test-KMSServerDiscovery {
    [Cmdletbinding()]
    Param($DNSSuffix)

    if (!$DNSSuffix) {
        $key = Get-Item -Path HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters
        $DNSSuffix = $key.GetValue("Domain")
    }

    $record = "_vlmcs._tcp.${DNSSuffix}"
    $NameRegEx = "\s+svr hostname   = (?<HostName>.*)$"
    $PortRegEX = "\s+(port)\s+ = (?<Port>\d+)"

    try {
        $results = nslookup -type=srv $record 2>&1 | Select-String "svr hostname" -Context 4,0
        if (!$results) { return }

        $myobj = @{}
        switch -regex ($results -split "\n") {
            $NameRegEx  { $myobj.ComputerName = $Matches.HostName }
            $PortRegEX  { $myobj.Port = $Matches.Port }
        }

        $obj = New-Object PSObject -Property $myobj
        $obj.PSTypeNames.Clear()
        $obj.PSTypeNames.Add('BSonPosh.KMS.DiscoveryResult')
        $obj
    } catch {
        Write-Verbose " [Test-KMSServerDiscovery] :: Error: $($lastError[0])"
    }
}
#==================================================================
# SUBBLOQUE: Verificar si un equipo soporta activación KMS
#==================================================================
function Test-KMSSupport {
    [Cmdletbinding()]
    Param(
        [Alias('dnsHostName')]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true)]
        [string]$ComputerName = $Env:COMPUTERNAME
    )

    Process {
        if ($ComputerName -match "(.*)(\$)$") {
            $ComputerName = $ComputerName -replace "(.*)(\$)$",'$1'
        }

        if (Test-Host -ComputerName $ComputerName -TCPPort 135) {
            $Query = "Select __CLASS FROM SoftwareLicensingProduct"
            try {
                $Result = Get-WmiObject -Query $Query -ComputerName $ComputerName
                if ($Result) { $_ }
            } catch {
                Write-Verbose " [Test-KMSSupport] :: Error: $($lastError[0])"
            }
        }
    }
}
#==================================================================
# BLOQUE FUNCIONAL: Wake-On-LAN (WOL)
#==================================================================
#==================================================================
# SUBBLOQUE: Función principal para encender equipos mediante WOL
#==================================================================
function Invoke-WakeOnLan {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName
    )
    
    $WOLVbsPath = Join-Path $Global:ScriptRoot "app\WOL\WOL.vbs"
    $WOLDirectory = Split-Path $WOLVbsPath
    
    Add-Logs -text "========================================="
    Add-Logs -text "[ENCENDER EQUIPO] $ComputerName"
    Add-Logs -text "========================================="
    
    # Comprobar que existe el archivo VBS
    if (-not (Test-Path $WOLVbsPath)) {
        Add-Logs -text "[ERROR] No se encuentra WOL.vbs"
        [System.Windows.Forms.MessageBox]::Show("No se encuentra el archivo WOL.vbs", "Error WOL", "OK", "Error")
        return
    }
    
    # Ejecutar el VBS en una ventana CMD visible
    Add-Logs -text "Abriendo ventana CMD para ejecutar WOL..."
    Add-Logs -text "Ejecutando: cscript.exe WOL.vbs $ComputerName"
    
    try {
        # Crear comando CMD que ejecute el VBS y mantenga la ventana abierta
        $cmdCommand = "cd /d `"$WOLDirectory`" && cscript.exe //NOLOGO WOL.vbs $ComputerName"
        
        # Abrir CMD con el comando
        Start-Process "cmd.exe" -ArgumentList "/k $cmdCommand" -WorkingDirectory $WOLDirectory -WindowStyle Normal
        
        Add-Logs -text "Ventana CMD abierta con proceso WOL"
        Add-Logs -text "========================================="
        
        [System.Windows.Forms.MessageBox]::Show("Se ha iniciado el proceso WOL para $ComputerName`n`nSigue el progreso en la ventana CMD que se ha abierto`n`nEl ping continuo se abrirá automáticamente al finalizar", "WOL Iniciado", "OK", "Information")
        
    } catch {
        Add-Logs -text "[ERROR] $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error al ejecutar WOL: $($_.Exception.Message)", "Error WOL", "OK", "Error")
    }
}
#==================================================================
# BLOQUE FUNCIONAL: Cálculos y Utilidades IP/Subnet
#==================================================================
#==================================================================
# SUBBLOQUE: Convertir IP decimal a binario en notación punteada
#==================================================================
function ConvertTo-BinaryIP {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Net.IPAddress]$IPAddress
    )
    Process {
        return [String]::Join('.', $(
            $IPAddress.GetAddressBytes() | ForEach-Object {
                [Convert]::ToString($_, 2).PadLeft(8, '0')
            }
        ))
    }
}
#==================================================================
# SUBBLOQUE: Convertir IP decimal a entero UInt32
#==================================================================
function ConvertTo-DecimalIP {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Net.IPAddress]$IPAddress
    )
    Process {
        $i = 3; $DecimalIP = 0
        $IPAddress.GetAddressBytes() | ForEach-Object {
            $DecimalIP += $_ * [Math]::Pow(256, $i)
            $i--
        }
        return [UInt32]$DecimalIP
    }
}
#==================================================================
# SUBBLOQUE: Convertir entero o binario punteado a IP decimal punteada
#==================================================================
function ConvertTo-DottedDecimalIP {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [String]$IPAddress
    )
    Process {
        switch -Regex ($IPAddress) {
            "([01]{8}\.){3}[01]{8}" {
                return [String]::Join('.', $(
                    $IPAddress.Split('.') | ForEach-Object {
                        [Convert]::ToUInt32($_, 2)
                    }
                ))
            }
            "\d" {
                $IPAddress = [UInt32]$IPAddress
                $DottedIP = for ($i = 3; $i -gt -1; $i--) {
                    $Remainder = $IPAddress % [Math]::Pow(256, $i)
                    ($IPAddress - $Remainder) / [Math]::Pow(256, $i)
                    $IPAddress = $Remainder
                }
                return [String]::Join('.', $DottedIP)
            }
            default {
                Write-Error "Cannot convert this format"
            }
        }
    }
}
#==================================================================
# SUBBLOQUE: Convertir máscara de red a longitud (CIDR)
#==================================================================
function ConvertTo-MaskLength {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Alias("Mask")]
        [Net.IPAddress]$SubnetMask
    )
    Process {
        $Bits = "$( $SubnetMask.GetAddressBytes() | ForEach-Object {
            [Convert]::ToString($_, 2)
        } )" -replace '[\s0]'
        return $Bits.Length
    }
}
#==================================================================
# SUBBLOQUE: Convertir longitud de máscara (CIDR) a máscara en formato punteado
#==================================================================
function ConvertTo-Mask {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Alias("Length")]
        [ValidateRange(0, 32)]
        $MaskLength
    )
    Process {
        return ConvertTo-DottedDecimalIP ([Convert]::ToUInt32(("1" * $MaskLength).PadRight(32, "0"), 2))
    }
}
#==================================================================
# SUBBLOQUE: Calcular dirección de red a partir de IP y máscara
#==================================================================
function Get-NetworkAddress {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Net.IPAddress]$IPAddress,
        [Parameter(Mandatory = $True, Position = 1)]
        [Alias("Mask")]
        [Net.IPAddress]$SubnetMask
    )
    Process {
        return ConvertTo-DottedDecimalIP (
            (ConvertTo-DecimalIP $IPAddress) -band (ConvertTo-DecimalIP $SubnetMask)
        )
    }
}
#==================================================================
# SUBBLOQUE: Calcular dirección de broadcast de una red
#==================================================================
function Get-BroadcastAddress {
    [CmdLetBinding()]
    Param(
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $True)]
        [Net.IPAddress]$IPAddress,
        [Parameter(Mandatory = $True, Position = 1)]
        [Alias("Mask")]
        [Net.IPAddress]$SubnetMask
    )
    Process {
        return ConvertTo-DottedDecimalIP (
            (ConvertTo-DecimalIP $IPAddress) -bor
            ((-bnot (ConvertTo-DecimalIP $SubnetMask)) -band [UInt32]::MaxValue)
        )
    }
}
#==================================================================
# SUBBLOQUE: Resumen de red desde IP y máscara (clase, rango, privados...)
#==================================================================
function Get-NetworkSummary ([String]$IP, [String]$Mask) {
    if ($IP.Contains("/")) {
        $Temp = $IP.Split("/")
        $IP = $Temp[0]
        $Mask = $Temp[1]
    }
    if (!$Mask.Contains(".")) {
        $Mask = ConvertTo-Mask $Mask
    }

    $DecimalIP = ConvertTo-DecimalIP $IP
    $DecimalMask = ConvertTo-DecimalIP $Mask
    $Network = $DecimalIP -band $DecimalMask
    $Broadcast = $DecimalIP -bor ((-bnot $DecimalMask) -band [UInt32]::MaxValue)

    $NetInfo = New-Object PSObject
    Add-Member -InputObject $NetInfo -NotePropertyName "Network" -Value (ConvertTo-DottedDecimalIP $Network)
    Add-Member -InputObject $NetInfo -NotePropertyName "Broadcast" -Value (ConvertTo-DottedDecimalIP $Broadcast)
    Add-Member -InputObject $NetInfo -NotePropertyName "Range" -Value (
        "$(ConvertTo-DottedDecimalIP ($Network + 1)) - $(ConvertTo-DottedDecimalIP ($Broadcast - 1))")
    Add-Member -InputObject $NetInfo -NotePropertyName "Mask" -Value $Mask
    Add-Member -InputObject $NetInfo -NotePropertyName "MaskLength" -Value (ConvertTo-MaskLength $Mask)
    Add-Member -InputObject $NetInfo -NotePropertyName "Hosts" -Value ($Broadcast - $Network - 1)

    $BinaryIP = ConvertTo-BinaryIP $IP
    $Private = $false
    switch -regex ($BinaryIP) {
        "^1111" { $Class = "E" }
        "^1110" { $Class = "D" }
        "^110"  { $Class = "C"; if ($BinaryIP -match "^11000000.10101000") { $Private = $true } }
        "^10"   { $Class = "B"; if ($BinaryIP -match "^10101100.0001") { $Private = $true } }
        "^0"    { $Class = "A"; if ($BinaryIP -match "^00001010") { $Private = $true } }
    }
    Add-Member -InputObject $NetInfo -NotePropertyName "Class" -Value $Class
    Add-Member -InputObject $NetInfo -NotePropertyName "IsPrivate" -Value $Private

    return $NetInfo
}
#==================================================================
# SUBBLOQUE: Obtener rango completo de direcciones de host entre IP y broadcast
#==================================================================
function Get-NetworkRange ([String]$IP, [String]$Mask) {
    if ($IP.Contains("/")) {
        $Temp = $IP.Split("/")
        $IP = $Temp[0]
        $Mask = $Temp[1]
    }
    if (!$Mask.Contains(".")) {
        $Mask = ConvertTo-Mask $Mask
    }
    $DecimalIP = ConvertTo-DecimalIP $IP
    $DecimalMask = ConvertTo-DecimalIP $Mask
    $Network = $DecimalIP -band $DecimalMask
    $Broadcast = $DecimalIP -bor ((-bnot $DecimalMask) -band [UInt32]::MaxValue)

    for ($i = $Network + 1; $i -lt $Broadcast; $i++) {
        ConvertTo-DottedDecimalIP $i
    }
}
#==================================================================
# BLOQUE FUNCIONAL: Comprobación de Servidores Remotos (Test-Server)
#==================================================================
#==================================================================
# SUBBLOQUE: Función principal Test-Server
#==================================================================
function Test-Server {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$ComputerName,

        [Parameter(Mandatory = $false)]
        [switch]$CredSSP,

        [Management.Automation.PSCredential]$Credential
    )

    begin {
        $total = Get-Date
        $results = @()
        if ($CredSSP -and -not $Credential) {
            Write-Host "Debe suministrar credenciales con CredSSP"; break
        }
    }

    process {
        foreach ($name in $ComputerName) {
            $dt = $cdt = Get-Date
            Write-Verbose "Testing: $name"
            $failed = 0
            try {
                # SUBBLOQUE: Resolución DNS
                $DNSEntity = [Net.Dns]::GetHostEntry($name)
                $domain = ($DNSEntity.HostName).Replace("$name.", "")
                $ips = $DNSEntity.AddressList | ForEach-Object { $_.IPAddressToString }
            } catch {
                $rst = "" | Select-Object Name, IP, Domain, Ping, WSMAN, CredSSP, RemoteReg, RPC, RDP
                $rst.Name = $name
                $results += $rst
                $failed = 1
            }

            Write-Verbose "DNS:  $((New-TimeSpan $dt ($dt = Get-Date)).TotalSeconds)"
            if ($failed -eq 0) {
                foreach ($ip in $ips) {
                    $rst = "" | Select-Object Name, IP, Domain, Ping, WSMAN, CredSSP, RemoteReg, RPC, RDP
                    $rst.Name = $name
                    $rst.IP = $ip
                    $rst.Domain = $domain

                    # SUBBLOQUE: RDP (puerto 3389)
                    try {
                        $socket = New-Object Net.Sockets.TcpClient($name, 3389)
                        $rst.RDP = $null -ne $socket
                        $socket.Close()
                    } catch { $rst.RDP = $false }
                    Write-Verbose "RDP: $((New-TimeSpan $dt ($dt = Get-Date)).TotalSeconds)"

                    # SUBBLOQUE: Ping
                    if (Test-Connection $ip -Count 1 -Quiet) {
                        $rst.Ping = $true
                        Write-Verbose "PING: $((New-TimeSpan $dt ($dt = Get-Date)).TotalSeconds)"

                        # SUBBLOQUE: WSMan
                        try {
                            Test-WSMan $ip | Out-Null
                            $rst.WSMAN = $true
                        } catch { $rst.WSMAN = $false }
                        Write-Verbose "WSMAN: $((New-TimeSpan $dt ($dt = Get-Date)).TotalSeconds)"

                        # SUBBLOQUE: CredSSP si WSMan activo y credenciales presentes
                        if ($rst.WSMAN -and $CredSSP) {
                            try {
                                Test-WSMan $ip -Authentication Credssp -Credential $Credential
                                $rst.CredSSP = $true
                            } catch { $rst.CredSSP = $false }
                            Write-Verbose "CredSSP: $((New-TimeSpan $dt ($dt = Get-Date)).TotalSeconds)"
                        }

                        # SUBBLOQUE: Registro remoto
                        try {
                            [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
                                [Microsoft.Win32.RegistryHive]::LocalMachine, $ip) | Out-Null
                            $rst.RemoteReg = $true
                        } catch { $rst.RemoteReg = $false }
                        Write-Verbose "remote reg: $((New-TimeSpan $dt ($dt = Get-Date)).TotalSeconds)"

                        # SUBBLOQUE: RPC/WMI básico
                        try {
                            $w = [Wmi]''
                            $w.psbase.options.timeout = 15000000
                            $w.path = "\\$name\root\cimv2:Win32_ComputerSystem.Name='$name'"
                            $w | Select-Object none | Out-Null
                            $rst.RPC = $true
                        } catch { $rst.RPC = $false }
                        Write-Verbose "WMI: $((New-TimeSpan $dt ($dt = Get-Date)).TotalSeconds)"

                    } else {
                        $rst.Ping = $false
                        $rst.WSMAN = $false
                        $rst.CredSSP = $false
                        $rst.RemoteReg = $false
                        $rst.RPC = $false
                    }
                    $results += $rst
                }
            }
            Write-Verbose "Tiempo para $name $((New-TimeSpan $cdt ($dt)).TotalSeconds)"
            Write-Verbose "----------------------------"
        }
    }

    end {
        Write-Verbose "Tiempo total: $((New-TimeSpan $total ($dt)).TotalSeconds)"
        Write-Verbose "----------------------------"
        return $results
    }
}
#==================================================================
# SUBBLOQUE: Obtener configuración IP de los adaptadores de red
#==================================================================
function Get-IPConfig {
    param (
        $Computername = "LocalHost",
        $OnlyConnectedNetworkAdapters = $true
    )
    Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName |
        Where-Object { $_.IPEnabled -eq $OnlyConnectedNetworkAdapters } |
        Format-List `
            @{ Label = "Computer Name"; Expression = { $_.__SERVER } },
            IPEnabled, Description, MACAddress, IPAddress, IPSubnet,
            DefaultIPGateway, DHCPEnabled, DHCPServer,
            @{ Label = "DHCP Lease Expires"; Expression = { [datetime]$_.DHCPLeaseExpires } },
            @{ Label = "DHCP Lease Obtained"; Expression = { [datetime]$_.DHCPLeaseObtained } }
}
#==================================================================
# SUBBLOQUE: Obtener propiedades de sitios y aplicaciones IIS remotas
#==================================================================
function get-iisProperties {
    [CmdletBinding(DefaultParameterSetName = 'ComputerName', ConfirmImpact = 'low')]
    Param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string][ValidatePattern(".{2,}")]$ComputerName
    )

    Begin {
        $lastError.Clear()
        $ComputerName = $ComputerName.ToUpper()
        $array = @()
    }

    Process {
        $objWMI = [WmiSearcher] "Select * From IIsWebServer"
        $objWMI.Scope.Path = "\\$ComputerName\root\microsoftiisv2"
        $objWMI.Scope.Options.Authentication = [System.Management.AuthenticationLevel]::PacketPrivacy

        try {
            $obj = $objWMI.Get()
            $obj | ForEach-Object {
                $Identifier = $_.Name
                $adsiPath = "IIS://$ComputerName/" + $_.Name
                $iis = [adsi]$adsiPath
                $iis.Psbase.Children |
                    Where-Object { $_.SchemaClassName -in "IIsWebVirtualDir", "IIsWebDirectory" } |
                    ForEach-Object {
                        $currentPath = "$adsiPath/$($_.Name)"
                        $_.Psbase.Children |
                            Where-Object { $_.SchemaClassName -eq "IIsWebVirtualDir" } |
                            Select-Object Name, AppPoolId, SchemaClassName, Path |
                            ForEach-Object {
                                $subIIS = [adsi]"$currentPath/$($_.Name)"
                                foreach ($mapping in $subIIS.ScriptMaps) {
                                    if ($mapping.StartsWith(".aspx")) {
                                        $NETversion = $mapping.ToLower().Substring($mapping.ToLower().IndexOf("framework\") + 10, 9)
                                    }
                                }
                                $tmpObj = New-Object PSObject -Property @{
                                    Name            = $_.Name
                                    Identifier      = $Identifier
                                    "ASP.NET"       = $NETversion
                                    AppPoolId       = $_.AppPoolId
                                    SchemaClassName = $_.SchemaClassName
                                    Path            = $_.Path
                                }
                                $array += $tmpObj
                            }
                    }
            }
        } catch {
            Write-Warning "Error: $($_.Exception.Message)"
        }
    }

    End {
        $array | Format-Table -AutoSize
    }
}
#==================================================================
# SUBBLOQUE: Obtener comentario (descripción) del equipo en registro
#==================================================================
function Get-ComputerComment($ComputerName) {
    $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $ComputerName)
    if (!$Registry) { return "Can't connect to the registry" }
    $RegKey = $Registry.OpenSubKey("SYSTEM\CurrentControlSet\Services\lanmanserver\parameters")
    if (!$RegKey) { return "No Computer Description" }
    $Description = $RegKey.GetValue("srvcomment")
    if (-not $Description) { $Description = "No Computer Description" }
    return "Computer Description: $Description"
}
#==================================================================
# SUBBLOQUE: Establecer comentario (descripción) del equipo en registro
#==================================================================
function Set-ComputerComment {
    param(
        [string]$ComputerName,
        [string]$Description
    )

    try {
        $Registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $ComputerName)
        if (!$Registry) { return $false }

        $RegKey = $Registry.OpenSubKey(
            "SYSTEM\CurrentControlSet\Services\lanmanserver\parameters",
            [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree,
            [System.Security.AccessControl.RegistryRights]::SetValue
        )
        if (!$RegKey) { return $false }

        $RegKey.SetValue("srvcomment", $Description)
        $RegKey.Close()
        $Registry.Close()

        return $true
    }
    catch {
        return $false
    }
}
#==================================================================
# SUBBLOQUE: Obtener dominio DNS del equipo
#==================================================================
function Get-DnsDomain {
    if (!$Global:DnsDomain) {
        $WmiInfo = Get-WmiObject "Win32_NTDomain" | Where-Object { $_.DnsForestName }
        $Global:DnsDomain = $WmiInfo.DnsForestName
    }
    return $Global:DnsDomain
}
#==================================================================
# SUBBLOQUE: Obtener ruta LDAP del dominio actual
#==================================================================
function Get-AdDomainPath {
    $DnsDomain = Get-DnsDomain
    $Tokens = $DnsDomain.Split(".")
    return ($Tokens | ForEach-Object { "DC=$_" }) -join ","
}
#==================================================================
# SUBBLOQUE: Obtener descripción del equipo en AD
#==================================================================
function Get-ComputerAdDescription($ComputerName) {
    $Path = Get-AdDomainPath
    $Dom = "LDAP://$Path"
    $Root = New-Object DirectoryServices.DirectoryEntry $Dom
    $Selector = New-Object DirectoryServices.DirectorySearcher
    $Selector.SearchRoot = $Root
    $Selector.Filter = "(objectclass=computer)"
    $AdObjects = $Selector.FindAll() | Where-Object { $_.Properties.cn -match $ComputerName }
    if (!$AdObjects) { return $null }
    return $AdObjects.Properties["description"]
}
#==================================================================
# SUBBLOQUE: Mostrar cuadro de mensaje gráfico (tipo MsgBox)
#==================================================================
function Show-MsgBox {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true)] [string]$Prompt,
        [Parameter(Position = 1)] [string]$Title = "",
        [Parameter(Position = 2)] [ValidateSet("Information", "Question", "Critical", "Exclamation")] [string]$Icon = "Information",
        [Parameter(Position = 3)] [ValidateSet("OKOnly", "OKCancel", "AbortRetryIgnore", "YesNoCancel", "YesNo", "RetryCancel")] [string]$BoxType = "OkOnly",
        [Parameter(Position = 4)] [ValidateSet(1, 2, 3)] [int]$DefaultButton = 1
    )
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null

    $vb_icon = [Microsoft.VisualBasic.MsgBoxStyle]::$Icon
    $vb_box = [Microsoft.VisualBasic.MsgBoxStyle]::$BoxType
    $vb_defaultbutton = [Microsoft.VisualBasic.MsgBoxStyle]::("DefaultButton$DefaultButton")

    $popuptype = $vb_icon -bor $vb_box -bor $vb_defaultbutton
    return [Microsoft.VisualBasic.Interaction]::MsgBox($Prompt, $popuptype, $Title)
}
#==================================================================
# SUBBLOQUE: Ejecutar un comando en una máquina remota usando WMI
#==================================================================
function Invoke-RemoteCMD {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$ComputerName,
        [string]$Command
    )
    begin {
        [string]$cmd = "CMD.EXE /C " + $Command
    }
    process {
        $newproc = Invoke-WmiMethod -Class Win32_Process -Name Create -ArgumentList ($cmd) -ComputerName $ComputerName
        if ($newproc.ReturnValue -eq 0) {
            Write-Output "Command '$Command' invoked successfully on $ComputerName"
        }
    }
    end {
        Write-Output "Script ...END"
    }
}
#==================================================================
# SUBBLOQUE: Comprobar si el equipo tiene PSRemoting habilitado
#==================================================================
function Test-PSRemoting {
    Param(
        [Alias('dnsHostName')]
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ValueFromPipeline = $true)]
        [string]$ComputerName
    )
    Process {
        Write-Verbose " [Test-PSRemoting] :: Start Process"
        if ($ComputerName -match "(.*)(\$)$") {
            $ComputerName = $ComputerName -replace "(.*)(\$)$",'$1'
        }
        try {
            $result = Invoke-Command -ComputerName $ComputerName { 1 } -ErrorAction SilentlyContinue
            return ($result -eq 1)
        } catch {
            return $False
        }
    }
}
#==================================================================
# SUBBLOQUE: Mostrar cuadro de entrada (InputBox)
#==================================================================
function Show-InputBox {
    Param(
        [string]$message = $(Throw "You must enter a prompt message"),
        [string]$title = "Input",
        [string]$default
    )
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null
    return [Microsoft.VisualBasic.Interaction]::InputBox($message, $title, $default)
}
#==================================================================
# BLOQUE FUNCIONAL: Ventana Acerca de (About)
#==================================================================
function Show-AboutPff
{
	#==================================================================
	# SUBBLOQUE: Importación de Ensamblados - Import the Assemblies
	#==================================================================
	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	#==================================================================
	# SUBBLOQUE: Objetos del Formulario Generado - Generated Form Objects
	#==================================================================
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form_Author = New-Object 'System.Windows.Forms.Form'
	$labelLastUpdateApplicatio = New-Object 'System.Windows.Forms.Label'
	$labelAbout = New-Object 'System.Windows.Forms.Label'
	$labelLazyWinAdminIsAPower = New-Object 'System.Windows.Forms.Label'
	$labelAuthorName = New-Object 'System.Windows.Forms.Label'
	$labelEmail = New-Object 'System.Windows.Forms.Label'
	$linklabel_Email = New-Object 'System.Windows.Forms.LinkLabel'
	$label_Author = New-Object 'System.Windows.Forms.Label'
	$button_AuthorOK = New-Object 'System.Windows.Forms.Button'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#==================================================================
	# SUBBLOQUE: Script Generado por el Usuario - User Generated Script
	#==================================================================
	#==================================================================
	# SUBBLOQUE: Función de Carga Inicial - OnApplicationLoad
	#==================================================================
	function OnApplicationLoad {
		# Nota: Esta función se ejecuta antes de que se cree el formulario
		# Nota: Para obtener el directorio del script en el Packager, usar: Split-Path $hostinvocation.MyCommand.path
		# Nota: Para obtener la salida de la consola en el Packager (modo Windows), usar: $ConsoleOutput (Tipo: System.Collections.ArrayList)
		# Importante: No se puede acceder a los controles del formulario en esta función
		# TODO: Agregar snap-ins y código personalizado para validar la carga de la aplicación
		return $true
	}
	#==================================================================
	# SUBBLOQUE: Función de Finalización - OnApplicationExit
	#==================================================================
	function OnApplicationExit {
		# Nota: Esta función se ejecuta después de que se cierra el formulario
		# TODO: Agregar código personalizado para limpiar y descargar snap-ins al salir de la aplicación
		$script:ExitCode = 0 # Establecer el código de salida para el Packager
	}
	#==================================================================
	# SUBBLOQUE: Eventos del Formulario Principal - FormEvent_Load y Links
	#==================================================================
	$FormEvent_Load = {
		$linklabel_Email.Text = $AuthorEmail
	}

	$linklabel_AuthorEmail_LinkClicked = [System.Windows.Forms.LinkLabelLinkClickedEventHandler]{
		[System.Diagnostics.Process]::Start("mailto:$AuthorEmail")
	}	
	#==================================================================
	# SUBBLOQUE: Eventos Generados - Generated Events
	#==================================================================
	$Form_StateCorrection_Load = {
		# Corrige el estado inicial del formulario para evitar el problema de maximizado en .NET
		$form_Author.WindowState = $InitialFormWindowState
	}

	$Form_StoreValues_Closing = {
		# Guardar los valores de los controles al cerrar (pendiente de implementación)
	}

	$Form_Cleanup_FormClosed = {
		try {
			$linklabel_Email.remove_LinkClicked($linklabel_AuthorEmail_LinkClicked)
			$form_Author.remove_Load($FormEvent_Load)
			$form_Author.remove_Load($Form_StateCorrection_Load)
			$form_Author.remove_Closing($Form_StoreValues_Closing)
			$form_Author.remove_FormClosed($Form_Cleanup_FormClosed)
		} catch [Exception] {
			# Manejo silencioso de errores
		}
	}
	#==================================================================
	#region Generated Form Code
	#==================================================================
	$form_Author.Controls.Add($labelLastUpdateApplicatio)
	$form_Author.Controls.Add($labelAbout)
	$form_Author.Controls.Add($labelLazyWinAdminIsAPower)
	$form_Author.Controls.Add($labelAuthorName)
	$form_Author.Controls.Add($labelEmail)
	$form_Author.Controls.Add($linklabel_Email)
	$form_Author.Controls.Add($label_Author)
	$form_Author.Controls.Add($button_AuthorOK)
	$form_Author.AcceptButton = $button_AuthorOK
	$form_Author.ClientSize = '290, 200'
	$form_Author.FormBorderStyle = 'FixedDialog'
	$form_Author.MaximizeBox = $False
	$form_Author.MinimizeBox = $False
	$form_Author.Name = "form_Author"
	$form_Author.Text = "Author"
	$form_Author.add_Load($FormEvent_Load)
	#==================================================================
	# SUBBLOQUE: labelLastUpdateApplicatio
	#==================================================================
	$labelLastUpdateApplicatio.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelLastUpdateApplicatio.Location = '21, 32'
	$labelLastUpdateApplicatio.Name = "labelLastUpdateApplicatio"
	$labelLastUpdateApplicatio.Size = '242, 23'
	$labelLastUpdateApplicatio.TabIndex = 11
	$labelLastUpdateApplicatio.Text = "Last Update: $ApplicationLastUpdate"
	#==================================================================
	# SUBBLOQUE: labelAbout
	#==================================================================
	$labelAbout.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$labelAbout.Location = '21, 9'
	$labelAbout.Name = "labelAbout"
	$labelAbout.Size = '242, 23'
	$labelAbout.TabIndex = 10
	$labelAbout.Text = "$ApplicationName $ApplicationVersion"
	#==================================================================
	# labelLazyWinAdminIsAPower
	#==================================================================
	$labelLazyWinAdminIsAPower.Location = '21, 61'
	$labelLazyWinAdminIsAPower.Name = "labelLazyWinAdminIsAPower"
	$labelLazyWinAdminIsAPower.Size = '244, 67'
	$labelLazyWinAdminIsAPower.TabIndex = 7
	$labelLazyWinAdminIsAPower.Text = "RCNG - Remote Computer Network GUI Tool.

	Basado en LazyWinAdmin (Sapien PowerShell Studio 2012)."
	#==================================================================
	# labelAuthorName
	#==================================================================
	$labelAuthorName.Location = '78, 137'
	$labelAuthorName.Name = "labelAuthorName"
	$labelAuthorName.Size = '198, 23'
	$labelAuthorName.TabIndex = 6
	$labelAuthorName.Text = "$AuthorName"
	#==================================================================
	# labelEmail
	#==================================================================
	$labelEmail.Location = '21, 137'
	$labelEmail.Name = "labelEmail"
	$labelEmail.Size = '51, 23'
	$labelEmail.TabIndex = 5
	$labelEmail.Text = "Email"
	#==================================================================
	# linklabel_Email
	#==================================================================
	$linklabel_Email.Location = '78, 137'
	$linklabel_Email.Name = "linklabel_Email"
	$linklabel_Email.Size = '187, 23'
	$linklabel_Email.TabIndex = 1
	$linklabel_Email.TabStop = $True
	$linklabel_Email.Text = "$AuthorEmail"
	$linklabel_Email.add_LinkClicked($linklabel_AuthorEmail_LinkClicked)
	#==================================================================
	# label_Author
	#==================================================================
	$label_Author.Location = '21, 137'
	$label_Author.Name = "label_Author"
	$label_Author.Size = '43, 23'
	$label_Author.TabIndex = 2
	$label_Author.Text = "Author:"
	#==================================================================
	# button_AuthorOK
	#==================================================================
	$button_AuthorOK.DialogResult = 'OK'
	$button_AuthorOK.Location = '106, 170'
	$button_AuthorOK.Name = "button_AuthorOK"
	$button_AuthorOK.Size = '75, 23'
	$button_AuthorOK.TabIndex = 0
	$button_AuthorOK.Text = "OK"
	$button_AuthorOK.UseVisualStyleBackColor = $True
	#==================================================================
	# SUBBLOQUE: Configuración Final del Formulario
	#==================================================================
	# Se guarda el estado inicial de la ventana, se asignan los eventos
	# y se muestra el formulario modal al usuario.
	# Guardar el estado inicial de la ventana
	$InitialFormWindowState = $form_Author.WindowState

	# Asignar eventos de ciclo de vida del formulario
	$form_Author.add_Load($Form_StateCorrection_Load)
	$form_Author.add_FormClosed($Form_Cleanup_FormClosed)
	$form_Author.add_Closing($Form_StoreValues_Closing)

	# Mostrar el formulario en modo modal
	return $form_Author.ShowDialog()

}

#==================================================================
# BLOQUE FUNCIONAL: Ventana Principal (MainForm)
#==================================================================
function Show-MainForm_pff {
    #==================================================================
    # REGION: Importación de Ensamblados necesarios
    #==================================================================
    [void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    [void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    [void][reflection.assembly]::Load("System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	#==================================================================
	#region Generated Form Objects
	#==================================================================
	[System.Windows.Forms.Application]::EnableVisualStyles()
	$form_MainForm = New-Object 'System.Windows.Forms.Form'
	$iconPath = Join-Path -Path $PSScriptRoot -ChildPath 'NRC.ico'
	if (Test-Path $iconPath) {
		$form_MainForm.Icon = [System.Drawing.Icon]::new($iconPath)
	} else {
		Write-Warning "No se encontró el icono en: $iconPath"
	}
	$richtextbox_output = New-Object 'System.Windows.Forms.RichTextBox'
	$panel_ContentArea = New-Object 'System.Windows.Forms.Panel'
	$panel_Utilities = New-Object 'System.Windows.Forms.Panel'
	$panel_PKHeader = New-Object 'System.Windows.Forms.Panel'
	$label_UtilsTitle = New-Object 'System.Windows.Forms.Label'
	$button_PK_Add = New-Object 'System.Windows.Forms.Button'
	$panel_PKButtons = New-Object 'System.Windows.Forms.Panel'
	$panel_RTBButtons = New-Object 'System.Windows.Forms.Panel'
	$button_formExit = New-Object 'System.Windows.Forms.Button'
	$button_outputClear = New-Object 'System.Windows.Forms.Button'
	$button_ExportRTF = New-Object 'System.Windows.Forms.Button'
	$button_outputCopy = New-Object 'System.Windows.Forms.Button'
	$tabcontrol_computer = New-Object 'System.Windows.Forms.TabControl'
	$tabpage_general = New-Object 'System.Windows.Forms.TabPage'
	$buttonSendCommand = New-Object 'System.Windows.Forms.Button'
	$groupbox_ManagementConsole = New-Object 'System.Windows.Forms.GroupBox'
	$button_mmcCompmgmt = New-Object 'System.Windows.Forms.Button'
	$buttonServices = New-Object 'System.Windows.Forms.Button'

	$RegistroRemoto = New-Object 'System.Windows.Forms.Button'
	$buttonEventVwr = New-Object 'System.Windows.Forms.Button'
	$button_GPupdate = New-Object 'System.Windows.Forms.Button'
	$button_ping = New-Object 'System.Windows.Forms.Button'
	$button_remot = New-Object 'System.Windows.Forms.Button'
	$buttonRemoteAssistance = New-Object 'System.Windows.Forms.Button'
	$button_PsRemoting = New-Object 'System.Windows.Forms.Button'
	$buttonC = New-Object 'System.Windows.Forms.Button'
	$button_networkconfig = New-Object 'System.Windows.Forms.Button'
	$button_Restart = New-Object 'System.Windows.Forms.Button'
	$button_PowerOn = New-Object 'System.Windows.Forms.Button'
	$button_Shutdown = New-Object 'System.Windows.Forms.Button'
	$tabpage_ComputerOSSystem = New-Object 'System.Windows.Forms.TabPage'
	$groupbox_UsersAndGroups = New-Object 'System.Windows.Forms.GroupBox'
	$button_UsersGroupLocalUsers = New-Object 'System.Windows.Forms.Button'
	$button_UsersGroupLocalGroups = New-Object 'System.Windows.Forms.Button'
	$groupbox_software = New-Object 'System.Windows.Forms.GroupBox'
	$groupbox_ComputerDescription = New-Object 'System.Windows.Forms.GroupBox'
	$button_ComputerDescriptionChange = New-Object 'System.Windows.Forms.Button'
	$button_ComputerDescriptionQuery = New-Object 'System.Windows.Forms.Button'
	$groupbox_WindowsUpdate = New-Object 'System.Windows.Forms.GroupBox'
	$button_HotFix = New-Object 'System.Windows.Forms.Button'
	$groupbox_RemoteDesktop = New-Object 'System.Windows.Forms.GroupBox'
	$button_RDPDisable = New-Object 'System.Windows.Forms.Button'
	$button_RDPEnable = New-Object 'System.Windows.Forms.Button'
	$groupbox_WinRM_SO = New-Object 'System.Windows.Forms.GroupBox'
	$button_ActivarWinRM_SO = New-Object 'System.Windows.Forms.Button'
	$button_DeshabilitarWinRM_SO = New-Object 'System.Windows.Forms.Button'
	$button_PSRemotoGeneral = New-Object 'System.Windows.Forms.Button'
	$buttonApplications = New-Object 'System.Windows.Forms.Button'
	$button_PageFile = New-Object 'System.Windows.Forms.Button'
	$button_StartupCommand = New-Object 'System.Windows.Forms.Button'
	$groupbox_Hardware = New-Object 'System.Windows.Forms.GroupBox'
	$button_MotherBoard = New-Object 'System.Windows.Forms.Button'
	$button_Processor = New-Object 'System.Windows.Forms.Button'
	$button_Memory = New-Object 'System.Windows.Forms.Button'
	$button_SystemType = New-Object 'System.Windows.Forms.Button'
	$button_Printers = New-Object 'System.Windows.Forms.Button'
	$button_USBDevices = New-Object 'System.Windows.Forms.Button'
	$tabpage_network = New-Object 'System.Windows.Forms.TabPage'
	$button_ConnectivityTesting = New-Object 'System.Windows.Forms.Button'
	$button_NIC = New-Object 'System.Windows.Forms.Button'
	$button_networkIPConfig = New-Object 'System.Windows.Forms.Button'
	$button_networkTestPort = New-Object 'System.Windows.Forms.Button'
	$button_networkRouteTable = New-Object 'System.Windows.Forms.Button'
	$tabpage_processes = New-Object 'System.Windows.Forms.TabPage'
	$buttonCommandLineGridView = New-Object 'System.Windows.Forms.Button'
	$button_processAll = New-Object 'System.Windows.Forms.Button'
	$buttonCommandLine = New-Object 'System.Windows.Forms.Button'
	$groupbox1 = New-Object 'System.Windows.Forms.GroupBox'
	$textbox_processName = New-Object 'System.Windows.Forms.TextBox'
	$label_processEnterAProcessName = New-Object 'System.Windows.Forms.Label'
	$button_processTerminate = New-Object 'System.Windows.Forms.Button'
	$button_process100MB = New-Object 'System.Windows.Forms.Button'
	$button_ProcessGrid = New-Object 'System.Windows.Forms.Button'
	$button_processOwners = New-Object 'System.Windows.Forms.Button'
	$button_processLastHour = New-Object 'System.Windows.Forms.Button'
	$tabpage_services = New-Object 'System.Windows.Forms.TabPage'
	$button_servicesNonStandardUser = New-Object 'System.Windows.Forms.Button'
	$button_mmcServices = New-Object 'System.Windows.Forms.Button'
	$button_servicesAutoNotStarted = New-Object 'System.Windows.Forms.Button'
	$groupbox_Service_QueryStartStop = New-Object 'System.Windows.Forms.GroupBox'
	$textbox_servicesAction = New-Object 'System.Windows.Forms.TextBox'
	$button_servicesRestart = New-Object 'System.Windows.Forms.Button'
	$label_servicesEnterAServiceName = New-Object 'System.Windows.Forms.Label'
	$button_servicesQuery = New-Object 'System.Windows.Forms.Button'
	$button_servicesStart = New-Object 'System.Windows.Forms.Button'
	$button_servicesStop = New-Object 'System.Windows.Forms.Button'
	$button_servicesRunning = New-Object 'System.Windows.Forms.Button'
	$button_servicesAll = New-Object 'System.Windows.Forms.Button'
	$button_servicesGridView = New-Object 'System.Windows.Forms.Button'
	$button_servicesAutomatic = New-Object 'System.Windows.Forms.Button'
	$tabpage_diskdrives = New-Object 'System.Windows.Forms.TabPage'
	$button_DiskUsage = New-Object 'System.Windows.Forms.Button'
	$button_DiskPartition = New-Object 'System.Windows.Forms.Button'
	$button_DiskLogical = New-Object 'System.Windows.Forms.Button'
	$button_DiskMountPoint = New-Object 'System.Windows.Forms.Button'
	$button_DiskRelationship = New-Object 'System.Windows.Forms.Button'
	$button_DiskMappedDrive = New-Object 'System.Windows.Forms.Button'
	$tabpage_shares = New-Object 'System.Windows.Forms.TabPage'
	$button_mmcShares = New-Object 'System.Windows.Forms.Button'
	$button_SharesGrid = New-Object 'System.Windows.Forms.Button'
	$button_Shares = New-Object 'System.Windows.Forms.Button'
	$tabpage_eventlog = New-Object 'System.Windows.Forms.TabPage'
	$button_RebootHistory = New-Object 'System.Windows.Forms.Button'
	$button_mmcEvents = New-Object 'System.Windows.Forms.Button'
	$button_EventsLogNames = New-Object 'System.Windows.Forms.Button'
	$tabpage_ExternalTools = New-Object 'System.Windows.Forms.TabPage'
	$button_Rwinsta = New-Object 'System.Windows.Forms.Button'
	$button_AD_ShowGroups = New-Object 'System.Windows.Forms.Button'
	$button_AD_AddToGroup = New-Object 'System.Windows.Forms.Button'
	$groupbox_DomainGroups = New-Object 'System.Windows.Forms.GroupBox'  # para los botones de AD (Grupos de dominio)
	$tabpage_PowershellRemoting = New-Object 'System.Windows.Forms.TabPage'
	$button_ActivarWinRM = New-Object 'System.Windows.Forms.Button'
	$button_DeshabilitarWinRM = New-Object 'System.Windows.Forms.Button'
	$button_PowershellRemota = New-Object 'System.Windows.Forms.Button'
	$button_Qwinsta = New-Object 'System.Windows.Forms.Button'
	$button_MsInfo32 = New-Object 'System.Windows.Forms.Button'
	$button_DriverQuery = New-Object 'System.Windows.Forms.Button'
	$button_SystemInfoexe = New-Object 'System.Windows.Forms.Button'
	$button_PAExec = New-Object 'System.Windows.Forms.Button'
	$button_psexec = New-Object 'System.Windows.Forms.Button'
	$textbox_networktracertparam = New-Object 'System.Windows.Forms.TextBox'
	$button_networkTracert = New-Object 'System.Windows.Forms.Button'
	$button_networkNsLookup = New-Object 'System.Windows.Forms.Button'
	$button_networkPing = New-Object 'System.Windows.Forms.Button'
	$textbox_networkpathpingparam = New-Object 'System.Windows.Forms.TextBox'
	$textbox_pingparam = New-Object 'System.Windows.Forms.TextBox'
	$button_networkPathPing = New-Object 'System.Windows.Forms.Button'
	$groupbox_ComputerName = New-Object 'System.Windows.Forms.GroupBox'
	$label_UptimeStatus = New-Object 'System.Windows.Forms.Label'
	$textbox_computername = New-Object 'System.Windows.Forms.TextBox'
	$label_OSStatus = New-Object 'System.Windows.Forms.Label'
	$button_Check = New-Object 'System.Windows.Forms.Button'
	$label_PingStatus = New-Object 'System.Windows.Forms.Label'
	$label_Ping = New-Object 'System.Windows.Forms.Label'
	$label_PSRemotingStatus = New-Object 'System.Windows.Forms.Label'
	$label_Uptime = New-Object 'System.Windows.Forms.Label'
	$label_RDPStatus = New-Object 'System.Windows.Forms.Label'
	$label_OS = New-Object 'System.Windows.Forms.Label'
	$label_PermissionStatus = New-Object 'System.Windows.Forms.Label'
	$label_Permission = New-Object 'System.Windows.Forms.Label'
	$label_PSRemoting = New-Object 'System.Windows.Forms.Label'
	$label_RDP = New-Object 'System.Windows.Forms.Label'
	$label_WinRM = New-Object 'System.Windows.Forms.Label'
	$label_WinRMStatus = New-Object 'System.Windows.Forms.Label'
	$richtextbox_Logs = New-Object 'System.Windows.Forms.RichTextBox'
	$statusbar1 = New-Object 'System.Windows.Forms.StatusBar'
	$menustrip_principal = New-Object 'System.Windows.Forms.MenuStrip'
	$ToolStripMenuItem_AdminArsenal = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_CommandPrompt = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_Powershell = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_localhost = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_compmgmt = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_taskManager = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_services = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_regedit = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_mmc = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_about = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_AboutInfo = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$contextmenustripServer = New-Object 'System.Windows.Forms.ContextMenuStrip'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Ping = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_RDP = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_compmgmt = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_services = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_eventvwr = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_InternetExplorer = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_TerminalAdmin = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_ADSearchDialog = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_ADPrinters = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_DHCP = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_systemInformationMSinfo32exe = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_netstatsListening = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_registeredSnappins = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_otherLocalTools = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_certificateManager = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_devicemanager = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolstripseparator1 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$toolstripseparator3 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$ToolStripMenuItem_systemproperties = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$toolstripseparator4 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$toolstripseparator5 = New-Object 'System.Windows.Forms.ToolStripSeparator'
	$ToolStripMenuItem_sharedFolders = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_performanceMonitor = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_groupPolicyEditor = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_localUsersAndGroups = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_diskManagement = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_localSecuritySettings = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_scheduledTasks = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_PowershellISE = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$errorprovider1 = New-Object 'System.Windows.Forms.ErrorProvider'
	$tooltipinfo = New-Object 'System.Windows.Forms.ToolTip'
	$ToolStripMenuItem_sysInternals = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_adExplorer = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_resetCredenciaisVNC = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Qwinsta = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_rwinsta = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_GeneratePassword = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_scripts = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_apps = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$ToolStripMenuItem_configuracion = New-Object 'System.Windows.Forms.ToolStripMenuItem'
	$imagelistAnimation = New-Object 'System.Windows.Forms.ImageList'
	$timerCheckJob = New-Object 'System.Windows.Forms.Timer'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
	#==================================================================
	#endregion Generated Form Objects
	#==================================================================
    #==================================================================
    # SUBBLOQUE: Configuración Inicial y Variables Globales
    #==================================================================

    $ApplicationName       = "RCNG"
    $ApplicationVersion    = "2.8.0 - Remote Computer Network GUI"
    $ApplicationLastUpdate = "22/03/2026"

    $AuthorName        = if (Get-Command 'Get-AppSettingValue' -ErrorAction SilentlyContinue) {
        Get-AppSettingValue -Key 'SupportDisplayName' -DefaultValue 'NRC_APP Support'
    } else {
        'NRC_APP Support'
    }
    $AuthorEmail       = if (Get-Command 'Get-AppSettingValue' -ErrorAction SilentlyContinue) {
        Get-AppSettingValue -Key 'SupportEmail' -DefaultValue 'support@example.local'
    } else {
        'support@example.local'
    }

    $StatusBarStartUp  = "$AuthorName $AuthorEmail"
    $domain            = $env:userdomain.ToUpper()
    $MainFormTitle     = "$ApplicationName $ApplicationVersion - Última Actualización: $ApplicationLastUpdate - $domain\$env:username"

    $ErrorActionPreference = "SilentlyContinue"
    $ScriptPath   = Split-Path -Path $MyInvocation.MyCommand.Path
    $ToolsFolder  = $ScriptPath + "tools"
    $lastError = ""
    #==================================================================
    # SUBBLOQUE: Carga de Equipos desde CSV ( COMPUTERLIST )
    #==================================================================
	#CONFIGURACIÓN INICIAL
	Set-Location $ScriptPath
	$ComputersList_File = Join-Path -Path $ScriptRoot -ChildPath "csv\equipos_ejemplo.csv"

	# Inicialización: ahora usamos la base de datos `ComputerNames.sqlite` como fuente única
	# Dejar una variable por compatibilidad, pero no precargamos todo en memoria.
	$ComputersHashTable = @{}

	function Import-Computers {
		# Mantener función por compatibilidad; inicializa la conexión a la DB.
		try {
			Initialize-ComputerDB | Out-Null
			Write-Output "Usando base de datos ComputerNames para búsquedas."
		} catch {
			Write-Output "No se pudo inicializar DB de equipos: $($_.Exception.Message)"
		}
	}

	# Inicializar conexión DB al inicio
	Import-Computers

	#FORMULARIO Y CONTROLES
	# ListBox para las sugerencias
	$listbox_suggestions = New-Object Windows.Forms.ListBox
	$listbox_suggestions.Visible = $false
	$listbox_suggestions.Width = 300
	$listbox_suggestions.Height = 150
	$listbox_suggestions.Location = New-Object Drawing.Point(10, 75)
	$listbox_suggestions.BackColor = [System.Drawing.Color]::WhiteSmoke
	$listbox_suggestions.ForeColor = [System.Drawing.Color]::DarkSlateGray
	$listbox_suggestions.BorderStyle = 'FixedSingle'
	$listbox_suggestions.Font = New-Object System.Drawing.Font("Segoe UI", 10)
	$form_MainForm.Controls.Add($listbox_suggestions)

	#FUNCIONES DEL FORMULARIO
	# Mostrar sugerencias en el ListBox
	function ShowListBoxSuggestions {
		param ([array]$sugerencias)

		$listbox_suggestions.Items.Clear()
		foreach ($sugerencia in $sugerencias) {
			$listbox_suggestions.Items.Add($sugerencia)
		}

		if ($listbox_suggestions.Items.Count -gt 0) {
			$listbox_suggestions.Height = [Math]::Min($listbox_suggestions.Items.Count * 20, 200)
			$listbox_suggestions.Visible = $true
			$form_MainForm.Refresh()
		} else {
			$listbox_suggestions.Visible = $false
		}
	}

	# Filtrar equipos por nombre u OU
	function FiltrarEquipos {
		param ([string]$filtro)

		if ([string]::IsNullOrWhiteSpace($filtro)) { return @() }

		# Usar la base de datos para obtener resultados (case-insensitive LIKE)
		try {
			$rows = Get-ComputerByFilterDB -Filter $filtro -Limit 300
			if (-not $rows) { return @() }
			$sugerencias = $rows | ForEach-Object { $_.equipo }
			return $sugerencias | Sort-Object -Unique
		} catch {
			# Fallback a hashtable (en caso de error)
			$filtroUp = $filtro.ToUpper()
			$sugerencias = @()
			foreach ($key in $ComputersHashTable.Keys) {
				$equipo = $ComputersHashTable[$key]
				if (
					$equipo.Equipo.ToUpper() -like "*$filtroUp*" -or
					$equipo.OU.ToUpper() -like "*$filtroUp*"
				) {
					$sugerencias += $equipo.Equipo
				}
				if ($sugerencias.Count -ge 300) { break }
			}
			return $sugerencias | Sort-Object -Unique
		}
	}

	#EVENTOS
	# Manejar texto ingresado por el usuario en el TextBox
	function OnTextBoxComputerNameTextChanged {
		$filtroIngresado = $textbox_computername.Text.Trim()

		if ($filtroIngresado.Length -gt 1) {
			$sugerencias = FiltrarEquipos -filtro $filtroIngresado
			ShowListBoxSuggestions -sugerencias $sugerencias
		} else {
			$listbox_suggestions.Visible = $false
		}
	}

	# Asignar el evento TextChanged al TextBox
	$textbox_computername.add_TextChanged({ OnTextBoxComputerNameTextChanged })

	# Manejar interacción con las teclas en el TextBox
	$textbox_computername.add_KeyDown({
		param($eventSender, $e)
		if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Down -and $listbox_suggestions.Visible -and $listbox_suggestions.Items.Count -gt 0) {
			$listbox_suggestions.Focus()
			$listbox_suggestions.SelectedIndex = 0
		}
	})

	# Manejar interacción con el ListBox
	$listbox_suggestions.add_KeyDown({
		param($eventSender, $e)
		if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
			if ($listbox_suggestions.SelectedItem) {
				$textbox_computername.Text = $listbox_suggestions.SelectedItem
				$listbox_suggestions.Visible = $false
				$textbox_computername.Focus()
			}
		}
	})

	# Manejar doble clic en el ListBox
	$listbox_suggestions.add_DoubleClick({
		if ($listbox_suggestions.SelectedItem) {
			$textbox_computername.Text = $listbox_suggestions.SelectedItem
			$listbox_suggestions.Visible = $false
			$textbox_computername.Focus()
		}
	})

	# Ocultar ListBox al hacer clic fuera de él
	$form_MainForm.add_MouseClick({
		$cursor_position = $form_MainForm.PointToClient([System.Windows.Forms.Cursor]::Position)

		if (-not ($listbox_suggestions.Bounds.Contains($cursor_position) -or $textbox_computername.Bounds.Contains($cursor_position))) {
			$listbox_suggestions.Visible = $false
		}
	})

	#==================================================================
	# SUBBLOQUE: Botón de recarga manual de procedimientos
	#==================================================================

	# Crear un botón estilizado para actualizar los datos manualmente
	$button_reload = New-Object Windows.Forms.Button
	$button_reload.Text = "Actualizar Datos"
	$button_reload.Width = 120
	$button_reload.Height = 30
	$button_reload.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$button_reload.FlatStyle = 'Flat'
	$button_reload.FlatAppearance.BorderSize = 1
	$button_reload.FlatAppearance.BorderColor = [System.Drawing.Color]::SlateGray
	$button_reload.Location = '30, 68'
	$button_reload.Anchor = [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom

	# Asociar evento para recargar datos: sincroniza desde servidor y reconstruye menús
	$button_reload.Add_Click({
		$button_reload.Enabled = $false
		$button_reload.Text    = "Actualizando..."
		try {
			# --- 1) Sincronización desde servidor (copia datos y definiciones compartidas) ---
			$syncResult = $null
			if (Get-Command 'Invoke-FullSyncFromServer' -ErrorAction SilentlyContinue) {
				$syncResult = Invoke-FullSyncFromServer
				Add-Logs -text "Sync: $($syncResult.Message)" -ErrorAction SilentlyContinue
			}

			# --- 2) Reconectar la base local de equipos tras la sincronización ---
			Initialize-ComputerDB | Out-Null

			# --- 3) Limpiar resultados auxiliares de la UI ---
			ClearListBoxSuggestions

			# --- 4) Reconstruir menú Scripts (puede haber nuevos custom tras el sync) ---
			if (Get-Command 'Initialize-ScriptsMenu' -ErrorAction SilentlyContinue) {
				Initialize-ScriptsMenu -Menu $ToolStripMenuItem_scripts
			}

			# --- 5) Reconstruir menú Aplicacións (puede haber nuevas apps o iconos) ---
			if (Get-Command 'Initialize-AplicacionsMenu' -ErrorAction SilentlyContinue) {
				Initialize-AplicacionsMenu -Menu $ToolStripMenuItem_apps
			}

			# --- 6) Gestionar scripts eliminados del servidor ---
			if ($syncResult -and $syncResult.RemovedFromServer -and $syncResult.RemovedFromServer.Count -gt 0) {
				foreach ($removed in $syncResult.RemovedFromServer) {
					$rDel = [System.Windows.Forms.MessageBox]::Show(
						"El script '$($removed.Name)' ($($removed.FileName)) ya no existe en el servidor.`n`n¿Desea eliminarlo también localmente?",
						"Script eliminado del servidor",
						[System.Windows.Forms.MessageBoxButtons]::YesNo,
						[System.Windows.Forms.MessageBoxIcon]::Question
					)
					if ($rDel -eq [System.Windows.Forms.DialogResult]::Yes) {
						# Eliminar de BD local
						$db = @(Get-ScriptsDatabase | Where-Object { $_.FileName -ne $removed.FileName })
						Save-ScriptsDatabase -Database $db
						# Eliminar archivo local si existe
						$localFile = Join-Path $Global:ScriptRoot "scripts\$($removed.FileName)"
						if (Test-Path $localFile) { Remove-Item $localFile -Force -ErrorAction SilentlyContinue }
						Add-Logs -text "Script local eliminado: $($removed.FileName)" -ErrorAction SilentlyContinue
					}
				}
				# Reconstruir menú Scripts tras posibles eliminaciones
				if (Get-Command 'Initialize-ScriptsMenu' -ErrorAction SilentlyContinue) {
					Initialize-ScriptsMenu -Menu $ToolStripMenuItem_scripts
				}
			}

			# --- 7) Gestionar scripts con versión más nueva en el servidor ---
			if ($syncResult -and $syncResult.UpdatedOnServer -and $syncResult.UpdatedOnServer.Count -gt 0) {
				foreach ($upd in $syncResult.UpdatedOnServer) {
					$rUpd = [System.Windows.Forms.MessageBox]::Show(
						"El script '$($upd.Name)' ($($upd.FileName)) tiene una versión más nueva en el servidor.`n`n¿Desea actualizar la copia local?",
						"Versión más nueva en servidor",
						[System.Windows.Forms.MessageBoxButtons]::YesNo,
						[System.Windows.Forms.MessageBoxIcon]::Question
					)
					if ($rUpd -eq [System.Windows.Forms.DialogResult]::Yes) {
						try {
							Copy-Item -Path $upd.ServerPath -Destination $upd.LocalPath -Force -ErrorAction Stop
							Add-Logs -text "Script actualizado desde servidor: $($upd.FileName)" -ErrorAction SilentlyContinue
						} catch {
							Add-Logs -text "Error actualizando $($upd.FileName): $_" -ErrorAction SilentlyContinue
						}
					}
				}
			}

			Add-Logs -text "Actualización completada." -ErrorAction SilentlyContinue
		} catch {
			Add-Logs -text "Error durante la actualización: $_" -ErrorAction SilentlyContinue
		} finally {
			$button_reload.Text    = "Actualizar Datos"
			$button_reload.Enabled = $true
		}
	})

	$panel_RTBButtons.Controls.Add($button_reload)


	#==================================================================
	# SUBBLOQUE: Función para limpiar sugerencias del ListBox
	#==================================================================

	function ClearListBoxSuggestions {
		$listbox_suggestions.Items.Clear()
		$listbox_suggestions.Visible = $false
	}


	#==================================================================
	# SUBBLOQUE: Inicialización del formulario principal
	#==================================================================
	# Mensaje de bienvenida y log inicial
	$RichTexBoxLogsDefaultMessage = "Benvido a $ApplicationName"
	Add-logs -text "Path: $ToolsFolder"

	# Información básica del sistema operativo
	$current_OS = Get-WmiObject Win32_OperatingSystem
	$current_OS_caption = $current_OS.caption

	# Evento de carga del formulario principal
	$OnLoadFormEvent = {
		$statusbar1.Text = $StatusBarStartUp
		$form_MainForm.Text = $MainFormTitle
		$textbox_computername.Text = $env:COMPUTERNAME
		Add-Logs -text $RichTexBoxLogsDefaultMessage

		# Verificar herramientas externas disponibles

		if (Test-Path "$ToolsFolder\psexec.exe") {
			$button_psexec.ForeColor = 'green'
			Add-Logs -text "External Tools check - PsExec.exe found" } 
		else {
			$button_psexec.ForeColor = 'Red'; $button_psexec.enabled = $false
			Add-Logs -text "External Tools check - PsExec.exe not found - Button Disabled" }

		if (Test-Path "$ToolsFolder\paexec.exe") {
			$button_PAExec.ForeColor = 'Green'
			Add-Logs -text "External Tools check - PAExec.exe found" } 
		else {
			$button_PAExec.ForeColor = 'Red'; $button_paexec.enabled = $false
			Add-Logs -text "External Tools check - PAExec.exe not found - Button Disabled" }

		if ( Test-Path ( $adExplorerPath = Join-Path $ScriptRoot 'tools\AdExplorer\adexplorer.exe' ) ) {
			$ToolStripMenuItem_adExplorer.ForeColor = 'Green'
			Add-Logs -text "External Tools check - ADExplorer.exe found"
		}
		else {
			$ToolStripMenuItem_adExplorer.Enabled = $false
			Add-Logs -text "External Tools check - ADExplorer.exe not found - Button Disabled"
		}

		if (Test-Path "$env:systemroot/system32/msra.exe") {
			Add-Logs -text "External Tools check - MSRA.exe found" } 
		else {
			$buttonRemoteAssistance.enabled = $false
			Add-Logs -text "External Tools check - MSRA.exe not found (Remote Assistance) - Button Disabled" }

		if (Test-Path "$env:systemroot/system32/systeminfo.exe") {
			Add-Logs -text "External Tools check - Systeminfo.exe found" } 
		else {
			$button_SystemInfoexe.enabled = $false
			Add-Logs -text "External Tools check - Systeminfo.exe not found - Button Disabled" }

		if (Test-Path "$env:programfiles\Common Files\Microsoft Shared\MSInfo\msinfo32.exe") {
			Add-Logs -text "External Tools check - msinfo32.exe found" } 
		else {
			$button_MsInfo32.enabled = $false
			Add-Logs -text "External Tools check - msinfo32.exe not found - Button Disabled" }

		if (Test-Path "$env:systemroot/system32/driverquery.exe") {
			Add-Logs -text "External Tools check - Driverquery.exe found" } 
		else {
			$button_DriverQuery.enabled = $false
			Add-Logs -text "External Tools check - Driverquery.exe not found - Button Disabled" }
	}
	#==================================================================
	# SUBBLOQUE: Timers y acciones rápidas del menú y botones principales
	#==================================================================
	# Timer adicional para gestión visual del botón de ejecución
	$timerCheckJob_Tick2 = {
		if ($null -ne $timerCheckJob.Tag) {
			if ($timerCheckJob.Tag.State -ne 'Running') {
				$buttonStart.ImageIndex = -1
				$buttonStart.Enabled = $true
				$buttonStart.Visible = $true
				$timerCheckJob.Tag = $null
				$timerCheckJob.Stop()
			} else {
				if ($buttonStart.ImageIndex -lt $buttonStart.ImageList.Images.Count - 1) {
					$buttonStart.ImageIndex += 1
				} else {
					$buttonStart.ImageIndex = 0
				}
			}
		}
	}

	# Exponer controles UI en scope global para los módulos de sections
	$global:richtextbox_output                   = $richtextbox_output
	$global:richtextbox_Logs                     = $richtextbox_Logs
	$global:textbox_computername                 = $textbox_computername
	$global:panel_PKButtons                      = $panel_PKButtons

	# Accesos rápidos desde menú o botones
	# Delegar handlers a sections\FerramentasAdmin.psm1 (migradas)
	$sectionsPath = Join-Path $ScriptRoot 'sections\FerramentasAdmin.psm1'
	if (Test-Path $sectionsPath) { Import-Module $sectionsPath -Force -ErrorAction SilentlyContinue }

	# Delegar handlers del menú LocalHost a sections\LocalHost.psm1 (migradas)
	$sectionsLocalPath = Join-Path $ScriptRoot 'sections\LocalHost.psm1'
	if (Test-Path $sectionsLocalPath) { Import-Module $sectionsLocalPath -Force -ErrorAction SilentlyContinue }

	# ScriptRunner - utilidad genérica de ejecución de scripts (local y remoto)
	$scriptRunnerPath = Join-Path $ScriptRoot 'modules\ScriptRunner.psm1'
	if (Test-Path $scriptRunnerPath) { Import-Module $scriptRunnerPath -Force -ErrorAction SilentlyContinue }

	# Scripts section - menú Scripts ordenado con soporte de scripts en caliente
	$scriptsMenuPath = Join-Path $ScriptRoot 'sections\Scripts.psm1'
	if (Test-Path $scriptsMenuPath) { Import-Module $scriptsMenuPath -Force -ErrorAction SilentlyContinue }

	# Aplicacions section - menú Aplicacións: lanzador de apps y ferramentas internas
	$aplicacionsMenuPath = Join-Path $ScriptRoot 'sections\Aplicacions.psm1'
	if (Test-Path $aplicacionsMenuPath) { Import-Module $aplicacionsMenuPath -Force -ErrorAction SilentlyContinue }

	# Configuracion section - menú Configuración: gestión de bases de datos locales
	$configuracionMenuPath = Join-Path $ScriptRoot 'sections\Configuracion.psm1'
	if (Test-Path $configuracionMenuPath) { Import-Module $configuracionMenuPath -Force -ErrorAction SilentlyContinue }

	# PassKeeper section - panel lateral cifrado con almacenamiento por usuario.
	$passKeeperPath = Join-Path $ScriptRoot 'sections\PassKeeper.psm1'
	if (Test-Path $passKeeperPath) {
		Import-Module $passKeeperPath -Force -ErrorAction SilentlyContinue
		# Si existe un passkeeper previo en database\\, se toma como origen de migración.
		$pkDataFile = Join-Path $ScriptRoot 'database\passkeeper.json'
		Initialize-PassKeeper -ButtonsPanel $global:panel_PKButtons -DataFile $pkDataFile
	}

	$ToolStripMenuItem_CommandPrompt_Click = { Invoke-Ferramentas_CommandPrompt }
	$ToolStripMenuItem_Powershell_Click = { Invoke-Ferramentas_Powershell }
	$ToolStripMenuItem_compmgmt_Click = { Invoke-Ferramentas_compmgmt }
	$ToolStripMenuItem_taskManager_Click = { Invoke-Ferramentas_taskManager }
	$ToolStripMenuItem_services_Click = { Invoke-Ferramentas_services }
	$ToolStripMenuItem_regedit_Click = { Invoke-Ferramentas_regedit }
	$ToolStripMenuItem_mmc_Click = { Invoke-Ferramentas_mmc }

	# DHCP MMC - abrir consola DHCP para el equipo seleccionado (delegado al módulo)
	$ToolStripMenuItem_DHCP_Click = { Invoke-Ferramentas_DHCP }
	# LocalHost menu delegations
	$ToolStripMenuItem_netstatsListening_Click = { Invoke-LocalHost_netstatsListening }
	$ToolStripMenuItem_registeredSnappins_Click = { Invoke-LocalHost_registeredSnappins }
	$ToolStripMenuItem_resetCredenciaisVNC_Click = { Invoke-LocalHost_resetCredenciaisVNC }
	# Exponer también estas entradas como variables globales para que
	# las funciones del módulo puedan habilitarlas/deshabilitarlas.
	$global:ToolStripMenuItem_netstatsListening = $ToolStripMenuItem_netstatsListening
	$global:ToolStripMenuItem_registeredSnappins = $ToolStripMenuItem_registeredSnappins

	$button_ping_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Lanzando ping..."
		$button_ping.Enabled = $false
		Start-Process ping -ArgumentList $($textbox_computername.Text), -t
		$button_ping.Enabled = $true
	}

	$button_remot_Click={
		Get-ComputerTxtBox
		add-logs -text "$ComputerName - Lanzando Escritorio remoto..."
		$port=":3389"
		$command = "mstsc"
		$argument = "/v:$computername$port /admin"
		Start-Process $command $argument
	}

	# Handler moved to sections\LocalHost.psm1: Invoke-LocalHost_registeredSnappins

	$button_outputClear_Click = { Clear-RichTextBox }
	$ToolStripMenuItem_AboutInfo_Click = { Show-AboutPff }

	$button_mmcCompmgmt_Click = {
		Get-ComputerTxtBox
		$button_mmcCompmgmt.Enabled = $false
		if ($ComputerName -match "(?i)^(localhost|\.|127\.0\.0\.1|$env:COMPUTERNAME)$") {
			Add-Logs -text "Localhost - Computer Management MMC (compmgmt.msc)"
			Start-Process compmgmt.msc
		} else {
			Add-Logs -text "$ComputerName - Computer Management MMC (compmgmt.msc /computer:$ComputerName)"
			Start-Process compmgmt.msc "/computer:$ComputerName"
		}
		$button_mmcCompmgmt.Enabled = $true
	}

	$ToolStripMenuItem_InternetExplorer_Click = { Invoke-Ferramentas_InternetExplorer }

	$button_Shares_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Shares list"
		$SharesList = Get-WmiObject win32_share -ComputerName $ComputerName | Sort-Object name | Format-Table -AutoSize | Out-String
		Add-RichTextBox -text $SharesList
	}

	$button_formExit_Click = {
		$ExitConfirmation = Show-MsgBox -Prompt "Do you really want to Exit ?" -Title "$ApplicationName $ApplicationVersion - Exit" -BoxType YesNo
		if ($ExitConfirmation -eq "YES") { $form_MainForm.Close() }
	}
	
	$button_mmcServices_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Services MMC (services.msc /computer:$ComputerName)"
		Start-Process "services.msc" "/computer:$ComputerName"
	}
	
	$button_servicesRunning_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Services - Status: Running"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		$Services_running = Get-Service -ComputerName $ComputerName | Where-Object { $_.Status -eq "Running" } | Format-Table -AutoSize | Out-String
		Add-RichTextBox -text $Services_running
	}
	
	$button_process100MB_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Processes >100MB"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		$owners = @{}
		Get-WmiObject win32_process -ComputerName $ComputerName | ForEach-Object { $owners[$_.handle] = $_.getowner().user }
		$Processes_Over100MB = Get-Process -ComputerName $ComputerName | Where-Object { $_.WorkingSet -gt 100mb } |
			Select-Object Handles, NPM, PM, WS, VM, CPU, ID, ProcessName, @{l="Owner";e={$owners[$_.id.tostring()]}} |
			Sort-Object ws | Format-Table -AutoSize | Out-String
		Add-RichTextBox $Processes_Over100MB
	}
	
	$button_mmcEvents_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Event Viewer MMC (eventvwr $ComputerName)"
		Start-Process "eventvwr" "$ComputerName"
	}
	
	$button_servicesAutomatic_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Services - StartMode:Automatic"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		$Services_StartModeAuto = Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "startmode='auto'" |
			Select-Object DisplayName, Name, ProcessID, StartMode, State |
			Format-Table -AutoSize | Out-String
		Add-RichTextBox $Services_StartModeAuto
	}
	
	$button_servicesQuery_Click = {
		Get-ComputerTxtBox
		Add-Logs "$ComputerName - Query Service"
		$Service_query = $textbox_servicesAction.text
		$a = New-Object -ComObject wscript.shell
		$intAnswer = $a.popup("Do you want to continue ?", 0, "$ComputerName - Query Service: $Service_query", 4)
		if ($intAnswer -eq 6) {
			if ($ComputerName -like "localhost") {
				Add-Logs "$ComputerName - Checking Service $Service_query ..."
				$Service_query_return = Get-WmiObject Win32_Service -Filter "Name='$Service_query'" | Out-String
			} else {
				Add-Logs "$ComputerName - Checking Service $Service_query ..."
				$Service_query_return = Get-WmiObject -ComputerName $ComputerName Win32_Service -Filter "Name='$Service_query'" | Out-String
			}
			Add-RichTextBox $Service_query_return
			Add-Logs -Text "$ComputerName - Query Service $Service_query - Done."
		}
	}
	
	$button_servicesAll_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Services - All Services + Owners"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		$Services_All = Get-WmiObject Win32_Service -ComputerName $ComputerName |
			Select-Object Name, ProcessID, StartMode, State, @{Name="Owner";Expression={$_.StartName}} |
			Format-Table -AutoSize | Out-String
		Add-RichTextBox $Services_All
	}
	
	$button_servicesStop_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Stop Service"
		$Service_query = $textbox_servicesAction.text
		$a = New-Object -ComObject wscript.shell
		$intAnswer = $a.popup("Do you want to continue ?", 0, "$ComputerName - Stop Service: $Service_query", 4)
		if ($intAnswer -eq 6) {
			if ($ComputerName -like "localhost") {
				Add-Logs -text "$ComputerName - Stopping Service: $Service_query ..."
				$Service_query_return = Get-WmiObject Win32_Service -Filter "Name='$Service_query'"
				$Service_query_return.StopService()
				Start-Sleep -Milliseconds 1000
				$Service_query_result = Get-WmiObject Win32_Service -Filter "Name='$Service_query'" | Out-String
			} else {
				Add-Logs -text "$ComputerName - Stopping Service: $Service_query ..."
				$Service_query_return = Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "Name='$Service_query'"
				$Service_query_return.StopService()
				Start-Sleep -Milliseconds 1000
				$Service_query_result = Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "Name='$Service_query'" | Out-String
			}
			Add-Logs -Text "$ComputerName - Command Sent! $Service_query should be stopped"
			Add-RichTextBox $Service_query_return
			Add-RichTextBox $Service_query_result
			Add-Logs -Text "$ComputerName - Stop Service $Service_query - Done."
		}
	}

	$button_DiskLogical_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Hard Drive - Logical Disk"
	
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
	
		$tempFile = [System.IO.Path]::GetTempFileName()
	
		try {
			$Disks = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName -Filter "DriveType != 1"
	
			if (!$Disks) {
				Add-RichTextBox -text "No se encontraron discos lógicos en $ComputerName."
				return
			}
	
			$info = $Disks | Select-Object `
				@{Name="Unidad";Expression={$_.DeviceID}},
				@{Name="Tipo";Expression={
					switch ($_.DriveType) {
						0 {"Desconocido"}
						1 {"Sin raíz"}
						2 {"USB"}
						3 {"Disco Local"}
						4 {"Red"}
						5 {"CD/DVD"}
						6 {"RAM"}
					}
				}},
				@{Name="FS";Expression={$_.FileSystem}},
				@{Name="Libre(GB)";Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
				@{Name="Total(GB)";Expression={[math]::Round($_.Size / 1GB, 2)}},
				@{Name="% Libre";Expression={if ($_.Size) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 2) } else { 0 }}},
				@{Name="% Uso";Expression={if ($_.Size) { [math]::Round((($_.Size - $_.FreeSpace) / $_.Size) * 100, 2) } else { 0 }}},
				@{Name="Volumen";Expression={$_.VolumeName}},
				@{Name="Dirty";Expression={$_.VolumeDirty}},
				@{Name="Serie";Expression={$_.VolumeSerialNumber}}
	
			# Guardar en archivo temporal
			$info | Format-Table -AutoSize | Out-File -FilePath $tempFile -Encoding UTF8
	
			# Leer contenido
			$contenido = Get-Content $tempFile -Raw
	
			# Mostrar
			Add-RichTextBox -text $contenido
	
			# Eliminar el archivo temporal justo después
			Remove-Item $tempFile -Force
		}
		catch {
			Add-RichTextBox -text "Error al obtener discos lógicos: $_"
			if (Test-Path $tempFile) {
				Remove-Item $tempFile -Force
			}
		}
	}	
	
	$button_EventsLogNames_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - EventLog - LogNames list"
		if ($ComputerName -eq "localhost") {
			$EventsLog = Get-EventLog -list | Format-List | Out-String
		} else {
			$EventsLog = Get-EventLog -list -ComputerName $ComputerName | Format-List | Out-String
		}
		Add-RichTextBox $EventsLog
	}
	
	$button_servicesStart_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Start Service"
		$Service_query = $textbox_servicesAction.text
		$a = New-Object -ComObject wscript.shell
		$intAnswer = $a.popup("Do you want to continue ?", 0, "$ComputerName - Start Service: $Service_query", 4)
		if ($intAnswer -eq 6) {
			if ($ComputerName -like "localhost") {
				Add-Logs -text "$ComputerName - Starting Service: $Service_query ..."
				$Service_query_return = Get-WmiObject Win32_Service -Filter "Name='$Service_query'"
				$Service_query_return.StartService()
				Start-Sleep -Milliseconds 1000
				$Service_query_result = Get-WmiObject Win32_Service -Filter "Name='$Service_query'" | Out-String
			} else {
				Add-Logs -text "$ComputerName - Starting Service: $Service_query ..."
				$Service_query_return = Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "Name='$Service_query'"
				$Service_query_return.StartService()
				Start-Sleep -Milliseconds 1000
				$Service_query_result = Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "Name='$Service_query'" | Out-String
			}
			Add-Logs -Text "$ComputerName - Command Sent! $Service_query should be started"
			Add-RichTextBox $Service_query_return
			Add-RichTextBox $Service_query_result
			Add-Logs -Text "$ComputerName - Start Service $Service_query - Done."
		}
	}
	
	$button_processOwners_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Processes with owners"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		$owners = @{}
		Get-WmiObject win32_process -ComputerName $ComputerName | ForEach-Object { $owners[$_.handle] = $_.getowner().user }
		$ProcessALL = Get-Process -ComputerName $ComputerName |
			Select-Object ProcessName, @{l="Owner";e={$owners[$_.id.tostring()]}}, CPU, WorkingSet, Handles, Id |
			Format-Table -AutoSize | Out-String
		Add-RichTextBox $ProcessALL
	}
	
	$button_processAll_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - All Processes"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		$ProcessALL = Get-Process -ComputerName $ComputerName | Out-String
		Add-RichTextBox $ProcessALL
	}
	
	$button_ProcessGrid_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - All Processes - GridView"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		$owners = @{}
		Get-WmiObject win32_process -ComputerName $ComputerName | ForEach-Object { $owners[$_.handle] = $_.getowner().user }
		Get-Process -ComputerName $ComputerName | Select-Object @{l="Owner";e={$owners[$_.id.tostring()]}}, * | Out-GridView
	}
	
	$button_servicesGridView_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - All Services - GridView"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		Get-WmiObject Win32_Service -ComputerName $ComputerName | Select-Object *, @{Name="Owner";Expression={$_.StartName}} | Out-GridView
	}
	
	$button_SharesGrid_Click = {
		Get-ComputerTxtBox
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		Get-WmiObject win32_share -ComputerName $ComputerName |
			Select-Object -Property __SERVER, Name, Path, Status, Description, * |
			Sort-Object name | Out-GridView
	}
	
	$ToolStripMenuItem_TerminalAdmin_Click = { Invoke-Ferramentas_TerminalAdmin }
	
	$ToolStripMenuItem_ADSearchDialog_Click = { Invoke-Ferramentas_ADSearchDialog }
	
	$ToolStripMenuItem_ADPrinters_Click = { Invoke-Ferramentas_ADPrinters }
	
	$button_outputCopy_Click = {
		Add-Logs -text "Copiando contido dos Logs ao Portapapeles..."
		$texte = $richtextbox_output.Text
		Add-ClipBoard -text $texte
	}
	
	$button_ExportRTF_Click = {
		$filename = [System.IO.Path]::GetTempFileName()
		$richtextbox_output.SaveFile($filename)
		Add-Logs -text "Enviando arquivo a notepad...."
		Start-Process notepad $filename
		Start-Sleep -Seconds 5
	}
	
	$button_networkPing_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Network - Ping"
		$cmd = "cmd"
		$param_user = $textbox_pingparam.text
		$param = "/k ping $param_user $ComputerName"
		Start-Process $cmd $param
	}
	
	$button_networkPathPing_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Network - PathPing"
		$cmd = "cmd"
		$param_user = $textbox_networkpathpingparam.Text
		$param = "/k pathping $param_user $ComputerName"
		Start-Process $cmd $param
	}
	
	#==================================================================
	# SUBBLOQUE: Acciones adicionales y administrativas
	#==================================================================

	$button_servicesNonStandardUser_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Services - Non-Standard Windows Service Accounts"
		$wql = 'Select Name, DisplayName, StartName, __Server From Win32_Service WHERE ((StartName != "LocalSystem") and (StartName != "NT Authority\\LocalService") and (StartName != "NT Authority\\NetworkService"))'
		$query = Get-WmiObject -Query $wql -ComputerName $ComputerName -ErrorAction Stop |
			Select-Object __SERVER, StartName, Name, DisplayName |
			Format-Table -AutoSize | Out-String
		if ($null -eq $query) {
			Add-RichTextBox "$ComputerName - All the services use Standard Windows Service Accounts"
		} else {
			Add-RichTextBox $query
		}
	}

	$button_networkTestPort_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Network - Test-Port"
		$port = Show-Inputbox -message "Enter a port to test" -title "$ComputerName - Test-Port" -default "80"
		if ($port -ne "") {
			$result = Test-TcpPort $ComputerName $port
			Add-RichTextBox $result
		}
	}

	$button_networkNsLookup_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Network - Nslookup"
		$cmd = "cmd"
		$param = "/k nslookup $ComputerName"
		Start-Process $cmd $param -WorkingDirectory c:\
	}

	$button_networkTracert_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Network - Trace Route (Tracert)"
		$cmd = "cmd"
		$param = "/k tracert $($textbox_networktracertparam.text) $ComputerName"
		Start-Process $cmd $param -WorkingDirectory c:\
	}

	$button_processLastHour_Click = {
		Get-ComputerTxtBox
		Add-Logs "$ComputerName - Processes - Processes started in last hour"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		$owners = @{}
		Get-WmiObject win32_process -ComputerName $ComputerName | ForEach-Object { $owners[$_.handle] = $_.getowner().user }
		$ProcessALL = Get-Process -ComputerName $ComputerName |
			Where-Object { trap { continue } (New-Timespan $_.StartTime).TotalMinutes -le 60 } |
			Select-Object ProcessName,
						@{l="StartTime";e={$_.StartTime}},
						@{l="Owner";e={$owners[$_.id.tostring()]}},
						CPU, WorkingSet, Handles, Id |
			Format-List | Out-String
		Add-RichTextBox $ProcessALL
	}

	$button_PasswordGen_Click = { Invoke-Ferramentas_GeneratePassword }

	$ToolStripMenuItem_systemInformationMSinfo32exe_Click = { Invoke-LocalHost_systemInformationMSinfo32exe }
	$ToolStripMenuItem_certificateManager_Click = { Invoke-LocalHost_certificateManager }
	$button_mmcShares_Click = {
		$ComputerName = $textbox_computername.Text
		Add-logs -text "$ComputerName - Shared Folders MMC (fsmgmt.msc /computer:$ComputerName)"
		$cmd = "fsmgmt.msc"
		$param = "/computer:$ComputerName"
		Start-Process $cmd $param
	}
	$ToolStripMenuItem_systemproperties_Click = { Invoke-LocalHost_systemproperties }
	$ToolStripMenuItem_sharedFolders_Click = { Invoke-LocalHost_sharedFolders }
	$ToolStripMenuItem_performanceMonitor_Click = { Invoke-LocalHost_performanceMonitor }
	$ToolStripMenuItem_devicemanager_Click = { Invoke-LocalHost_devicemanager }
	$ToolStripMenuItem_groupPolicyEditor_Click = { Invoke-LocalHost_groupPolicyEditor }
	$ToolStripMenuItem_localUsersAndGroups_Click = { Invoke-LocalHost_localUsersAndGroups }
	$ToolStripMenuItem_diskManagement_Click = { Invoke-LocalHost_diskManagement }
	$ToolStripMenuItem_localSecuritySettings_Click = { Invoke-LocalHost_localSecuritySettings }
	$ToolStripMenuItem_scheduledTasks_Click = { Invoke-LocalHost_scheduledTasks }
	$ToolStripMenuItem_PowershellISE_Click = { Invoke-LocalHost_PowershellISE }

	$button_servicesAutoNotStarted_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Services - Services with StartMode: Automatic and Status: NOT Running"
		if ($ComputerName -eq "localhost") { $ComputerName = "." }
		$Services_StartModeAuto = Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "startmode='auto' AND state!='running'" |
			Select-Object DisplayName, Name, StartMode, State |
			Format-Table -AutoSize | Out-String
		Add-RichTextBox $Services_StartModeAuto
	}

	$textbox_computername_TextChanged = {
		$label_OSStatus.Text = ""
		$label_PermissionStatus.Text = ""
		$label_PingStatus.Text = ""
		$label_RDPStatus.Text = ""
		$label_PSRemotingStatus.Text = ""
		$label_UptimeStatus.Text = ""
		$label_WinRMStatus.Text = ""
		if ($textbox_computername.Text -eq "") {
			$textbox_computername.BackColor = [System.Drawing.Color]::FromArgb(255, 128, 128)
			Add-logs -text "Please Enter a ComputerName"
			$errorprovider1.SetError($textbox_computername, "Please enter a ComputerName.")
		}
		if ($textbox_computername.Text -ne "") {
			$textbox_computername.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 192)
			$errorprovider1.SetError($textbox_computername, "")
		}
		$tabcontrol_computer.Enabled = $textbox_computername.Text -ne ""
		# No re-habilitar mientras haya una recogida de datos en curso
		if (-not $global:StreamRunning) {
			$button_Check.Enabled = $textbox_computername.Text -ne ""
		}
	}

	#==================================================================
	# SUBBLOQUE: Diagnósticos hardware y acciones generales
	#==================================================================

	$button_Processor_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Processor Information"
		$result = Get-Processor -ComputerName $ComputerName | Out-String
		Add-RichTextBox $result
	}

	$button_UsersGroupLocalUsers_Click = {
		Get-ComputerTxtBox
		$result = Get-WmiObject -class "Win32_UserAccount" -namespace "root\CIMV2" -filter "LocalAccount = True" -computername $ComputerName |
			Select-Object AccountType, Caption, Description, Disabled, Domain, FullName, InstallDate, LocalAccount, Lockout, Name, PasswordChangeable, PasswordExpires, PasswordRequired, SID, SIDType, Status |
			Format-List | Out-String
		Add-RichTextBox $result
	}

	$button_UsersGroupLocalGroups_Click = {
		$button_UsersGroupLocalGroups.Enabled = $false
		Get-ComputerTxtBox
		$result = Get-WmiObject -Class Win32_Group -ComputerName $ComputerName | Where-Object { $_.LocalAccount } | Format-Table -auto | Out-String
		Add-RichTextBox $result
		$button_UsersGroupLocalGroups.Enabled = $true
	}	

	$button_PageFile_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Page File Information and Settings"
	
		try {
			$ResultPageFile = Get-PageFile -ComputerName $ComputerName | Out-String
			$PageFileSetting = Get-PageFileSetting -ComputerName $ComputerName
	
			if ($PageFileSetting) {
				$ResultPageFileSettings = $PageFileSetting | Out-String
			} else {
				$ResultPageFileSettings = "Configuración de PageFile no encontrada. Probablemente esté gestionado automáticamente por el sistema."
			}
	
			Add-RichTextBox -text "$ResultPageFile `r`n$ResultPageFileSettings"
		}
		catch {
			Add-RichTextBox -text "Error al obtener información del archivo de paginación: $_"
		}
	}
	
	
	$button_DiskPartition_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Hard Drive - Partition"
		$result = Get-DiskPartition -ComputerName $ComputerName | Out-String
		Add-RichTextBox $result
	}
	
	$button_DiskUsage_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Hard Drive - DiskSpace"
	
		$result = Get-DiskSpace -ComputerName $ComputerName
	
		if (![string]::IsNullOrWhiteSpace($result)) {
			Add-RichTextBox -text $result
		} else {
			Add-RichTextBox -text "No se obtuvo información de disco para $ComputerName"
		}
	}
	
	
	$button_networkIPConfig_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Network - Configuration"
		$result = Get-IP -ComputerName $ComputerName |
			Format-Table Name, IP4, IP4Subnet, DefaultGWY, MacAddress, DNSServer, WinsPrimary, WinsSecondary -AutoSize |
			Out-String -Width $richtextbox_output.Width
		Add-RichTextBox "$result`n"
	}
	
	$button_DiskRelationship_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Hard Disk - Disks Relationship"
		$result = Get-DiskRelationship -ComputerName $ComputerName | Out-String
		Add-RichTextBox $result
	}
	
	$button_DiskMountPoint_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Hard Disk - MountPoint"
		$result = Get-MountPoint -ComputerName $ComputerName | Out-String
		if ($null -ne $result) {
			Add-RichTextBox $result
		} else {
			Show-MsgBox -BoxType "OKOnly" -Title "$ComputerName - Hard Disk - MountPoint" -Prompt "$ComputerName - No MountPoint detected" -Icon "Information"
		}
	}
	
	$button_DiskMappedDrive_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Hard Disk - Mapped Drive"
		$result = Get-MappedDrive -ComputerName $ComputerName | Out-String
		if ($null -ne $result) {
			Add-RichTextBox $result
		} else {
			Show-MsgBox -BoxType "OKOnly" -Title "$ComputerName - Mapped Drive" -Prompt "$ComputerName - No Mapped Drive detected" -Icon "Information"
		}
	}
	
	$button_Memory_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Memory Configuration"
		$result = Get-MemoryConfiguration -ComputerName $ComputerName | Out-String
		Add-RichTextBox $result
	}
	
	$button_NIC_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Network Interface Card Configuration (slow)"
		$result = Get-NICInfo -ComputerName $ComputerName | Out-String
		Add-RichTextBox $result
	}
	
	$button_MotherBoard_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - MotherBoard"
		$result = Get-MotherBoard -ComputerName $ComputerName | Out-String
		Add-RichTextBox $result
	}
	
	$button_networkRouteTable_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Route table"
		$result = Get-Routetable -ComputerName $ComputerName | Format-Table -auto | Out-String
		Add-RichTextBox $result
	}
	
	$button_SystemType_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - System Type"
		$result = get-systemtype -ComputerName $ComputerName | Out-String
		Add-RichTextBox $result
	}
	
	#==================================================================
	# SUBBLOQUE: Eventos de cambios en RichTextBox
	#==================================================================
	
	$richtextbox_output_TextChanged = {
		# Evitar interferencia con actualizaciones in-situ de DataCollection
		if (-not $global:StreamUpdating) {
			$richtextbox_output.SelectionStart = $richtextbox_output.Text.Length
			$richtextbox_output.ScrollToCaret()
		}
	}
	
	$richtextbox_Logs_TextChanged = {
		$richtextbox_Logs.SelectionStart = $richtextbox_Logs.Text.Length
		$richtextbox_Logs.ScrollToCaret()
		if ($lastError[0]) {
			Add-logs -text $($lastError[0].Exception.Message)
		}
	}
	
	$ToolStripMenuItem_adExplorer_Click = { Invoke-Ferramentas_adExplorer }
	
	#==================================================================
	# BLOQUE: Acciones Varias (Servicios, Hosts, Procesos, etc.)
	#==================================================================
	# SUBBLOQUE: Click en TextBox de servicios (limpia texto anterior)
	$textbox_servicesAction_Click = {
		$textbox_servicesAction.text = ""
	}

	# SUBBLOQUE: Botón Reiniciar Servicio
	$button_servicesRestart_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Restart Service"
		$Service_query = $textbox_servicesAction.text
		Add-logs -text "$ComputerName - Service to Restart: $Service_query"

		$a = New-Object -ComObject wscript.shell
		$intAnswer = $a.popup("Do you want to continue ?", 0, "$ComputerName - Start Service: $Service_query", 4)

		if (($ComputerName -like "localhost") -and ($intAnswer -eq 6)) {
			# Localhost: detener y reiniciar
			Add-logs -text "$ComputerName - Stopping Service: $Service_query ..."
			$Service_query_return = Get-WmiObject Win32_Service -Filter "Name='$Service_query'"
			$Service_query_return.stopservice()
			Add-Logs -Text "$ComputerName - Command Sent! $Service_query should be stopped"
			Add-RichTextBox $Service_query_return

			Start-Sleep -Milliseconds 1000
			$Service_query_result = Get-WmiObject Win32_Service -Filter "Name='$Service_query'" | Out-String
			Add-RichTextBox $Service_query_result

			Add-Logs -Text "$ComputerName - Restarting the Service $Service_query ..."
			$Service_query_return = Get-WmiObject Win32_Service -Filter "Name='$Service_query'"
			$Service_query_return.startservice()
			Add-Logs -Text "$ComputerName - Command Sent! $Service_query should be started"
			Add-RichTextBox $Service_query_return

			Start-Sleep -Milliseconds 1000
			$Service_query_result = Get-WmiObject Win32_Service -Filter "Name='$Service_query'" | Out-String
			Add-RichTextBox $Service_query_result
		}
		elseif ($intAnswer -eq 6) {
			# Remoto
			Add-logs -text "$ComputerName - Stopping Service: $Service_query ..."
			$Service_query_return = Get-WmiObject Win32_Service -Filter "Name='$Service_query'"
			$Service_query_return.stopservice()
			Add-Logs -Text "$ComputerName - Command Sent! $Service_query should be stopped"
			Add-RichTextBox $Service_query_return

			Start-Sleep -Milliseconds 1000
			$Service_query_result = Get-WmiObject Win32_Service -Filter "Name='$Service_query'" | Out-String
			Add-RichTextBox $Service_query_result

			Add-Logs -Text "$ComputerName - Restarting the Service $Service_query ..."
			$Service_query_return = Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "Name='$Service_query'"
			$Service_query_return.startservice()
			Add-Logs -Text "$ComputerName - Command Sent! $Service_query should be started"
			Add-RichTextBox $Service_query_return

			Start-Sleep -Milliseconds 1000
			$Service_query_result = Get-WmiObject Win32_Service -ComputerName $ComputerName -Filter "Name='$Service_query'" | Out-String
			Add-RichTextBox $Service_query_result
		}
	}

	# SUBBLOQUE: Menú - Resetear credenciales VNC
	# Handler moved to sections\LocalHost.psm1: Invoke-LocalHost_resetCredenciaisVNC

	# SUBBLOQUE: Botón - Terminar proceso
	$button_processTerminate_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Terminate Process"
		$Process_query = $textbox_processName.text
		Add-logs -text "$ComputerName - Process to Terminate: $Process_query"
		$a = New-Object -ComObject wscript.shell
		$intAnswer = $a.popup("Do you want to continue ?", 0, "$ComputerName - Terminate Process: $Process_query", 4)

		if (($ComputerName -like "localhost") -and ($intAnswer -eq 6)) {
			Add-logs -text "$ComputerName - Terminate Process: $Process_query - Terminating..."
			$Process_query_return = (Get-WmiObject Win32_Process -Filter "Name='$Process_query'").Terminate() | Out-String
			Add-RichTextBox $Process_query_return
			Start-Sleep -Milliseconds 1000
			$Process_query_return = Get-WmiObject Win32_Process -Filter "Name='$Process_query'" | Out-String
			if (!($Process_query_return)) { Add-Logs -Text "$ComputerName - $Process_query has been terminated" }
			Add-logs -text "$ComputerName - Terminate Process: $Process_query - Terminated "
		}
		elseif ($intAnswer -eq 6) {
			Add-logs -text "$ComputerName - Terminate Process: $Process_query - Terminating..."
			$Process_query_return = (Get-WmiObject Win32_Process -Filter "Name='$Process_query'").Terminate() | Out-String
			Add-RichTextBox $Process_query_return
			Start-Sleep -Milliseconds 1000
			$Process_query_return = Get-WmiObject Win32_Process -ComputerName $ComputerName -Filter "Name='$Process_query'" | Out-String
			if (!($Process_query_return)) { Add-Logs -Text "$ComputerName - Terminate Process: $Process_query - Terminated " }
		}
	}

	# SUBBLOQUE: Botón - Obtener comandos de inicio
	$button_StartupCommand_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Startup Commands"
		$result = Get-WmiObject Win32_StartupCommand -ComputerName $ComputerName | Sort-Object Caption | Format-Table __Server,Caption,Command,User -AutoSize | Out-String -Width $richtextbox_output.Width
		Add-RichTextBox $result
		Add-Logs -text "$ComputerName - Startup Commands - Done."
	}

	# SUBBLOQUE: Botón - Prueba de conectividad
	$button_ConnectivityTesting_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Connectivity Testing..."
		$result = Test-Server -ComputerName $ComputerName | Out-String
		Add-RichTextBox "$result`n"
	}


	# SUBBLOQUE: Botón "Check" para iniciar el análisis de equipo remoto/local. Ejecución Asíncrona
	$button_Check_Click = {
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Comprobando Conectividade e Propiedades Básicas"

		$button_Check.Enabled = $false
		$label_OSStatus.Text     = "Cargando..."
		$label_OSStatus.ForeColor = "blue"
		[System.Windows.Forms.Application]::DoEvents()

		# Limpiar suggestions
		ClearListBoxSuggestions

		# ── PING RÁPIDO EN HILO PRINCIPAL ─────────────────────────────
		$pingOK = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue
		if (-not $pingOK) {
			$label_PingStatus.Text      = "FAIL"
			$label_PingStatus.ForeColor = "red"
			$label_OSStatus.Text        = ""
			$label_OSStatus.ForeColor   = "black"
			Add-logs -text "❌ $ComputerName non responde ao ping."
			Add-RichTextBoxCheck -text "❌ Equipo '$ComputerName' non responde ao ping."
			$button_Check.Enabled = $true
			return
		}

		# Ping OK: lanzar recogida asíncrona
		$label_PingStatus.Text      = "OK"
		$label_PingStatus.ForeColor = "green"
		[System.Windows.Forms.Application]::DoEvents()

		# Verificar que la función de streaming está disponible
		if (-not (Get-Command 'Start-StreamingDataCollection' -ErrorAction SilentlyContinue)) {
			# Reintentar carga del módulo vía Import-Module
			$_dcPath = Join-Path $Global:ScriptRoot 'modules\DataCollection.psm1'
			if (Test-Path $_dcPath) {
				try {
					Import-Module $_dcPath -Force -DisableNameChecking -ErrorAction Stop
					Add-logs -text "⚠️ DataCollection.psm1 recargado (no estaba en memoria)"
				} catch {
					$_loadErr = $_.Exception.Message
					Add-logs -text "❌ ERROR cargando DataCollection.psm1: $_loadErr"
					[System.Windows.Forms.MessageBox]::Show(
						"Error cargando DataCollection.psm1:`n$_loadErr",
						"Error de módulo", "OK", "Error")
					$button_Check.Enabled = $true
					return
				}
			} else {
				Add-logs -text "❌ ERROR: modules\DataCollection.psm1 no encontrado. ScriptRoot='$Global:ScriptRoot'"
				[System.Windows.Forms.MessageBox]::Show(
					"No se encontró el módulo DataCollection.psm1.`nRuta buscada: $_dcPath`n`nScriptRoot: $Global:ScriptRoot",
					"Error de módulo", "OK", "Error")
				$button_Check.Enabled = $true
				return
			}
		}

		try {
			$launched = Start-StreamingDataCollection -ComputerNameParam $ComputerName -UI @{
				ButtonCheck   = $button_Check
				RTBOutput     = $richtextbox_output
				RTBLogs       = $richtextbox_Logs
				LblOS         = $label_OSStatus
				LblPing       = $label_PingStatus
				LblRDP        = $label_RDPStatus
				LblVNC        = $label_PSRemotingStatus
				LblWinRM      = $label_WinRMStatus
				LblPermission = $label_PermissionStatus
				LblUptime     = $label_UptimeStatus
			}
			if ($launched -eq $false) { $button_Check.Enabled = $true }
		} catch {
			Add-logs -text "❌ ERROR en Start-StreamingDataCollection: $($_.Exception.Message)"
			$label_OSStatus.Text      = "ERROR"
			$label_OSStatus.ForeColor = "red"
			$button_Check.Enabled     = $true
		}
	}

	#==================================================================
	# BLOQUE: Acciones Remotas Avanzadas y Utilidades
	#==================================================================

	# SUBBLOQUE: PsExec Terminal (CMD remoto)
	$button_psexec_Click={
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - PSEXEC (Terminal)"
		if(Test-Path "$ToolsFolder\psexec.exe"){
			$argument = "/k $ToolsFolder\psexec.exe \\$ComputerName cmd.exe"
			Start-Process cmd.exe $argument
		}
		else {$button_psexec.ForeColor = 'Red'}
	}

	# SUBBLOQUE: PAExec Terminal (CMD remoto con sesión del sistema)
	$button_PAExec_Click={
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - PAEXEC (Terminal)"
		$argument = "/k $ToolsFolder\paexec.exe \\$ComputerName -s cmd.exe"
		Start-Process cmd.exe $argument
	}

	# Handler moved to sections\LocalHost.psm1: Invoke-LocalHost_netstatsListening

	# Handler moved to sections\LocalHost.psm1: Invoke-LocalHost_limpaTemps

	# SUBBLOQUE: Consultar drivers instalados remotamente
	$button_DriverQuery_Click={
		$button_DriverQuery.Enabled = $False
		Get-ComputerTxtBox
		$DriverQuery_command="cmd.exe"
		$DriverQuery_arguments = "/k driverquery /s $ComputerName"
		Start-Process $DriverQuery_command $DriverQuery_arguments
		$button_DriverQuery.Enabled = $true
	}

	# SUBBLOQUE: Sesión remota con PsExec
	$button_PsRemoting_Click={
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Abre una sesión de cmd remota"
		$argument = "/k $ToolsFolder\psexec.exe \\$ComputerName cmd.exe"
		Start-Process cmd.exe $argument
	}

	# SUBBLOQUE: Información del sistema con msinfo32
	$button_MsInfo32_Click={
		Get-ComputerTXTBOX
		Add-Logs "$ComputerName - System Information (MSinfo32.exe)"
		$cmd = "$env:programfiles\Common Files\Microsoft Shared\MSInfo\msinfo32.exe"
		$param = "/computer $ComputerName"
		Start-Process $cmd $param
	}

	# SUBBLOQUE: Sesiones de usuario activas con QWINSTA
	$button_Qwinsta_Click={
		$button_Qwinsta.Enabled = $false
		Get-ComputerTXTBOX
		if ($current_OS_caption -notlike "*64*"){
			Add-Logs -text "$ComputerName - QWINSTA (Consultar sesión de terminal) - 32 bits"
			$Qwinsta_cmd = "cmd"
			$Qwinsta_argument = "/k qwinsta /server:$computername"
		} else {
			Add-Logs -text "$ComputerName - QWINSTA (Consultar sesión de terminal) - 64 bits"
			$Qwinsta_cmd = "cmd"
			$Qwinsta_argument = "/k $env:SystemRoot\Sysnative\qwinsta /server:$computername"
		}
		Start-Process $Qwinsta_cmd $Qwinsta_argument
		$button_Qwinsta.Enabled = $true
	}

	# SUBBLOQUE: Finalizar sesión de terminal con RWINSTA
	$button_Rwinsta_Click={
		$button_Rwinsta.Enabled = $false
		Get-ComputerTXTBOX
		Add-Logs -text "$ComputerName - RWINSTA (Reset Terminal Sessions)"
		$Rwinsta_ID = Show-Inputbox -message "Introduce ID sesión activa" -title "$ComputerName - Rwinsta (Restablecer sesión de terminal)"
		if ($Rwinsta_ID -ne ""){
			if ($current_OS_caption -notlike "*64*"){
				$Rwinsta_argument = "/k $env:SystemRoot\System32\rwinsta $Rwinsta_ID /server:$computername"
			} else {
				$Rwinsta_argument = "/k $env:SystemRoot\Sysnative\rwinsta $Rwinsta_ID /server:$computername"
			}
			Start-Process cmd $Rwinsta_argument
		}
		$button_Rwinsta.Enabled = $true
	}

	# SUBBLOQUE: Historial de reinicios (event ID 6009)
	$button_RebootHistory_Click={
		$button_RebootHistory.Enabled = $false
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Historial de reinicios"
		start-sleep -s 1
		$job = get-reboottime -ComputerName $ComputerName |Out-string
		Add-Richtextbox $job
		$button_RebootHistory.Enabled = $true
	}

	#==================================================================
	# SUBBLOQUE: Active Directory - Mostrar Grupos del Equipo
	#==================================================================
	$button_AD_ShowGroups_Click={
		$button_AD_ShowGroups.Enabled = $false
		Get-ComputerTxtBox
		
		if ([string]::IsNullOrWhiteSpace($ComputerName)) {
			Add-logs -text "❌ No hay equipo cargado. Por favor, ingresa un nombre de equipo primero."
			$button_AD_ShowGroups.Enabled = $true
			return
		}
		
		Add-logs -text "$ComputerName - Consultando grupos de Active Directory..."
		
		try {
			# Verificar si el módulo ActiveDirectory está disponible
			if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
				Add-RichTextBox "❌ ERROR: El módulo Active Directory no está instalado."
				Add-logs -text "Error: Módulo ActiveDirectory no disponible"
				$button_AD_ShowGroups.Enabled = $true
				return
			}
			
			Import-Module ActiveDirectory -ErrorAction Stop
			
			# Buscar el equipo en AD
			$computer = Get-ADComputer -Identity $ComputerName -Properties MemberOf -ErrorAction Stop
			
			if ($computer.MemberOf.Count -eq 0) {
				Add-RichTextBox "ℹ️ El equipo '$ComputerName' no pertenece a ningún grupo de AD (excepto Domain Computers)."
				Add-logs -text "$ComputerName - No pertenece a grupos adicionales"
			} else {
				$output = "GRUPOS DE ACTIVE DIRECTORY - $ComputerName`n`n"
				$output += "Total de grupos: $($computer.MemberOf.Count)`n`n"
				
				foreach ($groupDN in $computer.MemberOf) {
					# Extraer el nombre del grupo del Distinguished Name
					$groupName = ($groupDN -split ',')[0] -replace 'CN=', ''
					$output += "  • $groupName`n"
				}
				
				$output += "`nFinalizado: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')`n"
				
				Add-RichTextBox $output
				Add-logs -text "$ComputerName - Grupos mostrados correctamente ($($computer.MemberOf.Count) grupos)"
			}
		}
		catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
			Add-RichTextBox "❌ ERROR: El equipo '$ComputerName' no se encontró en Active Directory."
			Add-logs -text "$ComputerName - Equipo no encontrado en AD"
		}
		catch {
			Add-RichTextBox "❌ ERROR al consultar Active Directory: $($_.Exception.Message)"
			Add-logs -text "$ComputerName - Error consultando AD: $($_.Exception.Message)"
		}
		
		$button_AD_ShowGroups.Enabled = $true
	}

	#==================================================================
	# SUBBLOQUE: Powershell Remoting - Activar WinRM
	#==================================================================
	$button_ActivarWinRM_Click={
		$button_ActivarWinRM.Enabled = $false
		Get-ComputerTxtBox
		
		if ([string]::IsNullOrWhiteSpace($ComputerName)) {
			Add-logs -text "❌ No hay equipo cargado. Por favor, ingresa un nombre de equipo primero."
			$button_ActivarWinRM.Enabled = $true
			return
		}
		
		Add-logs -text "$ComputerName - Activando Powershell Remoting..."
		
		try {
			# Ruta al psexec
			$psexecPath = Join-Path $PSScriptRoot "tools\psexec.exe"
			
			if (-not (Test-Path $psexecPath)) {
				Add-logs -text "$ComputerName - Error: psexec.exe no encontrado"
				$button_ActivarWinRM.Enabled = $true
				return
			}
			
			$result1 = & $psexecPath -accepteula \\$ComputerName -s cmd /c "sc config winrm start=demand" 2>&1
			Start-Sleep -Milliseconds 500
			
			$result2 = & $psexecPath -accepteula \\$ComputerName -s cmd /c "net start winrm" 2>&1
			Start-Sleep -Milliseconds 500
			
			$result3 = & $psexecPath -accepteula \\$ComputerName -s powershell.exe -Command "Enable-PSRemoting -Force" 2>&1
			Start-Sleep -Milliseconds 1000
			
			# IMPORTANTE: Enable-PSRemoting configura el servicio como automático, volver a manual
			$result4 = & $psexecPath -accepteula \\$ComputerName -s cmd /c "sc config winrm start=demand" 2>&1
			Start-Sleep -Milliseconds 1000
			
			# Verificar el tipo de inicio real usando sc qc (query config)
			$result5 = & $psexecPath -accepteula \\$ComputerName -s cmd /c "sc qc winrm" 2>&1
			
			# Determinar el StartType real
			$startTypeActual = "Desconocido"
			if ($result5 -match "AUTO_START") {
				$startTypeActual = "Automático"
			} elseif ($result5 -match "DEMAND_START") {
				$startTypeActual = "Manual"
			} elseif ($result5 -match "DISABLED") {
				$startTypeActual = "Deshabilitado"
			}
			
			# Verificar estado del servicio
			$result6 = & $psexecPath -accepteula \\$ComputerName -s cmd /c "sc query winrm" 2>&1
			
			if ($result6 -match "RUNNING") {
				$label_WinRMStatus.Text = "ACTIVO ($startTypeActual)"
				if ($startTypeActual -eq "Manual") {
					$label_WinRMStatus.ForeColor = "green"
					Add-logs -text "$ComputerName - Activado Powershell Remoting (Inicio: Manual)"
				} else {
					$label_WinRMStatus.ForeColor = "orange"
					Add-logs -text "$ComputerName - ⚠️ Activado pero quedó como $startTypeActual - Se requiere ajuste manual"
				}
			} else {
				Add-logs -text "$ComputerName - Error activando Powershell Remoting"
			}
		}
		catch {
			Add-logs -text "$ComputerName - Error activando Powershell Remoting: $($_.Exception.Message)"
		}
		
		$button_ActivarWinRM.Enabled = $true
	}

	#==================================================================
	# SUBBLOQUE: Powershell Remoting - Deshabilitar WinRM
	#==================================================================
	$button_DeshabilitarWinRM_Click={
		$button_DeshabilitarWinRM.Enabled = $false
		Get-ComputerTxtBox
		
		if ([string]::IsNullOrWhiteSpace($ComputerName)) {
			Add-logs -text "❌ No hay equipo cargado. Por favor, ingresa un nombre de equipo primero."
			$button_DeshabilitarWinRM.Enabled = $true
			return
		}
		
		Add-logs -text "$ComputerName - Deshabilitando Powershell Remoting..."
		
		try {
			$psexecPath = Join-Path $PSScriptRoot "tools\psexec.exe"
			
			if (-not (Test-Path $psexecPath)) {
				Add-logs -text "$ComputerName - Error: psexec.exe no encontrado"
				$button_DeshabilitarWinRM.Enabled = $true
				return
			}
			
			# Detener el servicio
			$result1 = & $psexecPath -accepteula \\$ComputerName -s cmd /c "net stop winrm" 2>&1
			Start-Sleep -Milliseconds 300
			
			# Configurar como deshabilitado
			$result2 = & $psexecPath -accepteula \\$ComputerName -s cmd /c "sc config winrm start=disabled" 2>&1
			Start-Sleep -Milliseconds 300
			
			# Verificar
			$result3 = & $psexecPath -accepteula \\$ComputerName -s cmd /c "sc query winrm" 2>&1
			
			if ($result3 -match "STOPPED") {
				$label_WinRMStatus.Text = "PARADO (Deshabilitado)"
				$label_WinRMStatus.ForeColor = "red"
				Add-logs -text "$ComputerName - Deshabilitado Powershell Remoting"
			} else {
				Add-logs -text "$ComputerName - Error deshabilitando Powershell Remoting"
			}
		}
		catch {
			Add-logs -text "$ComputerName - Error deshabilitando Powershell Remoting: $($_.Exception.Message)"
		}
		
		$button_DeshabilitarWinRM.Enabled = $true
	}

	#==================================================================
	# SUBBLOQUE: Powershell Remoting - Abrir Powershell Remota
	#==================================================================
	$button_PowershellRemota_Click={
		$button_PowershellRemota.Enabled = $false
		Get-ComputerTxtBox
		
		if ([string]::IsNullOrWhiteSpace($ComputerName)) {
			Add-logs -text "❌ No hay equipo cargado. Por favor, ingresa un nombre de equipo primero."
			$button_PowershellRemota.Enabled = $true
			return
		}
		
		Add-logs -text "$ComputerName - Verificando WinRM antes de abrir sesión remota..."
		
		try {
			# Verificar si WinRM está activo
			$winrmTest = Test-WSMan -ComputerName $ComputerName -ErrorAction Stop
			
			if ($winrmTest) {
				Add-logs -text "$ComputerName - Abriendo sesión remota de PowerShell..."
				
				# Abrir PowerShell en nueva ventana con Enter-PSSession
				$psCommand = "Enter-PSSession -ComputerName $ComputerName"
				Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", $psCommand
				
				Add-logs -text "$ComputerName - Sesión remota de PowerShell iniciada"
			}
		}
		catch {
			[System.Windows.Forms.MessageBox]::Show(
				"WinRM no está habilitado o no es accesible en el equipo '$ComputerName'.`n`n" +
				"Para habilitarlo, ve a la pestaña 'Equipo y Sistema Operativo' y pulsa el botón 'Activar WinRM'.",
				"WinRM no disponible",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Warning
			)
			Add-logs -text "$ComputerName - WinRM no disponible para sesión remota"
		}
		
		$button_PowershellRemota.Enabled = $true
	}

	#==================================================================
	# SUBBLOQUE: Active Directory - Agregar Equipo a Grupo
	#==================================================================
	$button_AD_AddToGroup_Click={
		$button_AD_AddToGroup.Enabled = $false
		Get-ComputerTxtBox
		
		if ([string]::IsNullOrWhiteSpace($ComputerName)) {
			Add-logs -text "❌ No hay equipo cargado. Por favor, ingresa un nombre de equipo primero."
			$button_AD_AddToGroup.Enabled = $true
			return
		}
		
		# Crear ventana personalizada para seleccionar o escribir grupo
		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing
		
		$formGrupo = New-Object System.Windows.Forms.Form
		$formGrupo.Text = "Agregar '$ComputerName' a Grupo AD"
		$formGrupo.Size = New-Object System.Drawing.Size(550,300)
		$formGrupo.StartPosition = "CenterParent"
		$formGrupo.FormBorderStyle = 'FixedDialog'
		$formGrupo.MaximizeBox = $false
		
		# Etiqueta instrucciones
		$lblInstrucciones = New-Object System.Windows.Forms.Label
		$lblInstrucciones.Text = "Selecciona un grupo de las listas o escribe el nombre manualmente:"
		$lblInstrucciones.Location = New-Object System.Drawing.Point(20,15)
		$lblInstrucciones.Size = New-Object System.Drawing.Size(500,20)
		$lblInstrucciones.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
		$formGrupo.Controls.Add($lblInstrucciones)
		
		# Label para Primaria
		$lblPrimaria = New-Object System.Windows.Forms.Label
		$lblPrimaria.Text = "Grupos sugeridos (OU principal):"
		$lblPrimaria.Location = New-Object System.Drawing.Point(20,45)
		$lblPrimaria.Size = New-Object System.Drawing.Size(500,20)
		$lblPrimaria.Font = New-Object System.Drawing.Font("Segoe UI",9)
		$formGrupo.Controls.Add($lblPrimaria)
		
		# ComboBox Primaria
		$comboPrimaria = New-Object System.Windows.Forms.ComboBox
		$comboPrimaria.Location = New-Object System.Drawing.Point(20,68)
		$comboPrimaria.Size = New-Object System.Drawing.Size(500,25)
		$comboPrimaria.DropDownStyle = 'DropDownList'
		$comboPrimaria.Font = New-Object System.Drawing.Font("Segoe UI",9)
		$formGrupo.Controls.Add($comboPrimaria)
		
		# Label para Salud Pública
		$lblSaludPublica = New-Object System.Windows.Forms.Label
		$lblSaludPublica.Text = "Grupos sugeridos (OU secundaria):"
		$lblSaludPublica.Location = New-Object System.Drawing.Point(20,105)
		$lblSaludPublica.Size = New-Object System.Drawing.Size(500,20)
		$lblSaludPublica.Font = New-Object System.Drawing.Font("Segoe UI",9)
		$formGrupo.Controls.Add($lblSaludPublica)
		
		# ComboBox Salud Pública
		$comboSaludPublica = New-Object System.Windows.Forms.ComboBox
		$comboSaludPublica.Location = New-Object System.Drawing.Point(20,128)
		$comboSaludPublica.Size = New-Object System.Drawing.Size(500,25)
		$comboSaludPublica.DropDownStyle = 'DropDownList'
		$comboSaludPublica.Font = New-Object System.Drawing.Font("Segoe UI",9)
		$formGrupo.Controls.Add($comboSaludPublica)
		
		# Label para campo manual
		$lblManual = New-Object System.Windows.Forms.Label
		$lblManual.Text = "O escribe el nombre del grupo manualmente:"
		$lblManual.Location = New-Object System.Drawing.Point(20,165)
		$lblManual.Size = New-Object System.Drawing.Size(500,20)
		$lblManual.Font = New-Object System.Drawing.Font("Segoe UI",9)
		$formGrupo.Controls.Add($lblManual)
		
		# TextBox para nombre manual
		$txtGrupoManual = New-Object System.Windows.Forms.TextBox
		$txtGrupoManual.Location = New-Object System.Drawing.Point(20,188)
		$txtGrupoManual.Size = New-Object System.Drawing.Size(500,25)
		$txtGrupoManual.Font = New-Object System.Drawing.Font("Segoe UI",9)
		$formGrupo.Controls.Add($txtGrupoManual)
		
		# Botón Aceptar
		$btnAceptar = New-Object System.Windows.Forms.Button
		$btnAceptar.Text = "Aceptar"
		$btnAceptar.Location = New-Object System.Drawing.Point(280,225)
		$btnAceptar.Size = New-Object System.Drawing.Size(100,30)
		$btnAceptar.BackColor = [System.Drawing.Color]::LightGreen
		$btnAceptar.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
		$btnAceptar.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$formGrupo.Controls.Add($btnAceptar)
		
		# Botón Cancelar
		$btnCancelar = New-Object System.Windows.Forms.Button
		$btnCancelar.Text = "Cancelar"
		$btnCancelar.Location = New-Object System.Drawing.Point(390,225)
		$btnCancelar.Size = New-Object System.Drawing.Size(100,30)
		$btnCancelar.BackColor = [System.Drawing.Color]::LightCoral
		$btnCancelar.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
		$btnCancelar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$formGrupo.Controls.Add($btnCancelar)
		
		$formGrupo.AcceptButton = $btnAceptar
		$formGrupo.CancelButton = $btnCancelar
		
		# Cargar grupos desde AD
		try {
			Import-Module ActiveDirectory -ErrorAction Stop
			
			# Cargar grupos desde la base LDAP principal definida en configuracion
			$ouPrimaria = if (Get-Command 'Get-AppSettingValue' -ErrorAction SilentlyContinue) {
				Get-AppSettingValue -Key 'PrimaryGroupSearchBase' -DefaultValue 'OU=GrupoPrincipal,DC=example,DC=local'
			} else {
				'OU=GrupoPrincipal,DC=example,DC=local'
			}
			$gruposPrimaria = Get-ADGroup -Filter * -SearchBase $ouPrimaria -ErrorAction SilentlyContinue | 
				Select-Object -ExpandProperty Name | Sort-Object
			
			if ($gruposPrimaria) {
				$comboPrimaria.Items.Add("-- Selecciona un grupo --")
				foreach ($grupo in $gruposPrimaria) {
					$comboPrimaria.Items.Add($grupo)
				}
				$comboPrimaria.SelectedIndex = 0
			} else {
				$comboPrimaria.Items.Add("-- No se encontraron grupos --")
				$comboPrimaria.SelectedIndex = 0
				$comboPrimaria.Enabled = $false
			}
			
			# Cargar grupos desde la base LDAP secundaria definida en configuracion
			$ouSaludPublica = if (Get-Command 'Get-AppSettingValue' -ErrorAction SilentlyContinue) {
				Get-AppSettingValue -Key 'SecondaryGroupSearchBase' -DefaultValue 'OU=GrupoSecundario,DC=example,DC=local'
			} else {
				'OU=GrupoSecundario,DC=example,DC=local'
			}
			$gruposSaludPublica = Get-ADGroup -Filter * -SearchBase $ouSaludPublica -ErrorAction SilentlyContinue | 
				Select-Object -ExpandProperty Name | Sort-Object
			
			if ($gruposSaludPublica) {
				$comboSaludPublica.Items.Add("-- Selecciona un grupo --")
				foreach ($grupo in $gruposSaludPublica) {
					$comboSaludPublica.Items.Add($grupo)
				}
				$comboSaludPublica.SelectedIndex = 0
			} else {
				$comboSaludPublica.Items.Add("-- No se encontraron grupos --")
				$comboSaludPublica.SelectedIndex = 0
				$comboSaludPublica.Enabled = $false
			}
		}
		catch {
			$comboPrimaria.Items.Add("-- Error cargando grupos --")
			$comboPrimaria.SelectedIndex = 0
			$comboPrimaria.Enabled = $false
			$comboSaludPublica.Items.Add("-- Error cargando grupos --")
			$comboSaludPublica.SelectedIndex = 0
			$comboSaludPublica.Enabled = $false
		}
		
		# Evento para limpiar otros selectores al seleccionar Primaria
		$comboPrimaria.Add_SelectedIndexChanged({
			if ($comboPrimaria.SelectedIndex -gt 0 -and $comboPrimaria.SelectedItem -ne "-- Selecciona un grupo --") {
				$comboSaludPublica.SelectedIndex = 0
				$txtGrupoManual.Text = ""
			}
		})
		
		# Evento para limpiar otros selectores al seleccionar Salud Pública
		$comboSaludPublica.Add_SelectedIndexChanged({
			if ($comboSaludPublica.SelectedIndex -gt 0 -and $comboSaludPublica.SelectedItem -ne "-- Selecciona un grupo --") {
				$comboPrimaria.SelectedIndex = 0
				$txtGrupoManual.Text = ""
			}
		})
		
		# Evento para limpiar selectores al escribir manualmente
		$txtGrupoManual.Add_TextChanged({
			if ($txtGrupoManual.Text.Trim() -ne "") {
				$comboPrimaria.SelectedIndex = 0
				$comboSaludPublica.SelectedIndex = 0
			}
		})
		
		# Mostrar el formulario
		$resultado = $formGrupo.ShowDialog()
		
		# Procesar resultado
		if ($resultado -ne [System.Windows.Forms.DialogResult]::OK) {
			Add-logs -text "Operación cancelada por el usuario"
			$button_AD_AddToGroup.Enabled = $true
			return
		}
		
		# Determinar qué grupo se seleccionó
		$groupName = ""
		if ($txtGrupoManual.Text.Trim() -ne "") {
			$groupName = $txtGrupoManual.Text.Trim()
		}
		elseif ($comboPrimaria.SelectedIndex -gt 0 -and $comboPrimaria.SelectedItem -ne "-- Selecciona un grupo --") {
			$groupName = $comboPrimaria.SelectedItem
		}
		elseif ($comboSaludPublica.SelectedIndex -gt 0 -and $comboSaludPublica.SelectedItem -ne "-- Selecciona un grupo --") {
			$groupName = $comboSaludPublica.SelectedItem
		}
		
		if ([string]::IsNullOrWhiteSpace($groupName)) {
			Add-logs -text "Operación cancelada - No se seleccionó o ingresó ningún grupo"
			$button_AD_AddToGroup.Enabled = $true
			return
		}
		
		Add-logs -text "$ComputerName - Intentando agregar al grupo: $groupName"
		
		try {
			# Verificar si el módulo ActiveDirectory está disponible
			if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
				Add-RichTextBox "❌ ERROR: El módulo Active Directory no está instalado."
				Add-logs -text "Error: Módulo ActiveDirectory no disponible"
				$button_AD_AddToGroup.Enabled = $true
				return
			}
			
			Import-Module ActiveDirectory -ErrorAction Stop
			
			# Verificar que el equipo existe en AD
			$computer = Get-ADComputer -Identity $ComputerName -ErrorAction Stop
			
			# Verificar que el grupo existe
			$group = Get-ADGroup -Identity $groupName -ErrorAction Stop
			
			# Agregar el equipo al grupo
			Add-ADGroupMember -Identity $groupName -Members $computer -ErrorAction Stop
			
			$output = "✅ ÉXITO: El equipo '$ComputerName' ha sido agregado al grupo '$groupName'`n`n"
			$output += "Detalles:`n"
			$output += "  • Equipo: $($computer.Name)`n"
			$output += "  • DN: $($computer.DistinguishedName)`n"
			$output += "  • Grupo: $($group.Name)`n"
			$output += "  • Grupo DN: $($group.DistinguishedName)`n"
			$output += "  • Fecha: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')`n"
			
			Add-RichTextBox $output
			Add-logs -text "$ComputerName - Agregado exitosamente al grupo: $groupName"
			
			[System.Windows.Forms.MessageBox]::Show("Equipo agregado exitosamente al grupo '$groupName'", "Operación Exitosa", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
		}
		catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
			$errorMsg = if ($_.Exception.Message -match "computer") {
				"❌ ERROR: El equipo '$ComputerName' no se encontró en Active Directory."
			} else {
				"❌ ERROR: El grupo '$groupName' no se encontró en Active Directory."
			}
			Add-RichTextBox $errorMsg
			Add-logs -text "$ComputerName - Error: Equipo o grupo no encontrado"
			[System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
		}
		catch {
			$errorDetail = $_.Exception.Message
			Add-RichTextBox "❌ ERROR al agregar equipo al grupo: $errorDetail"
			Add-logs -text "$ComputerName - Error agregando a grupo: $errorDetail"
			[System.Windows.Forms.MessageBox]::Show("Error al agregar equipo al grupo:`n`n$errorDetail", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
		}
		
		$button_AD_AddToGroup.Enabled = $true
	}

	# SUBBLOQUE: Dispositivos USB conectados
	function Get-PhysicalUSBAndHIDDevices {
		param($ComputerName = "localhost")
		Get-WmiObject Win32_PnPEntity -ComputerName $ComputerName |
			Where-Object {
				(
					# Físicos USB EXCLUYENDO los lógicos y USB Input Device
					($_.DeviceID -like "USB*") -and
					($_.Name -notmatch "Hub|Composite|Bridge|Generic|Root Hub|Controller") -and
					($_.Description -notmatch "Hub|Composite|Bridge|Generic|Root Hub|Controller") -and
					($_.Service -notmatch "usbccgp|USBHUB3|USBHUB") -and
					($_.Name -ne "USB Input Device") -and
					($_.Description -ne "USB Input Device")
				) -or
				(
					# Solo HID que tengan 'Keyboard' en Name o Description
					($_.DeviceID -like "HID*") -and
					(
						($_.Name -match "Keyboard") -or
						($_.Description -match "Keyboard")
					)
				)
			} |
			Select-Object Manufacturer, Name, DeviceID, PNPDeviceID, Service, DriverVersion, Status, Description, @{Name="ClaseWMI";Expression={ $_.__Class }}
	}

	$button_USBDevices_Click = {
		$button_USBDevices.Enabled = $false
		try {
			Get-ComputerTxtBox
			Add-Logs "$ComputerName - USB/HID Physical Devices"
			$result = Get-PhysicalUSBAndHIDDevices -ComputerName $ComputerName |
				ForEach-Object {
					"Fabricante: $($_.Manufacturer)"
					"Nombre: $($_.Name)"
					"DeviceID: $($_.DeviceID)"
					"PNPDeviceID: $($_.PNPDeviceID)"
					"Servicio: $($_.Service)"
					"Versión de driver: $($_.DriverVersion)"
					"Estado: $($_.Status)"
					"Descripción: $($_.Description)"
					"ClaseWMI: $($_.ClaseWMI)"
					""
				} | Out-String
			Add-RichTextBox $result
		} catch {
			Add-RichTextBox "Error: $_"
			Add-Logs "Error querying USB/HID physical devices: $_"
		} finally {
			$button_USBDevices.Enabled = $true
		}
	}

	# SUBBLOQUE: Habilitar escritorio remoto
	$button_RDPEnable_Click={
		$button_RDPEnable.Enabled = $false
		Get-ComputerTxtBox
		Add-Logs "$ComputerName - Enable RDP"
		$RDPresult = Set-RDPEnable -ComputerName $ComputerName
		Add-logs -text "$ComputerName - Resultado de Set-RDPEnable: $RDPresult"
		Add-RichTextBox "Habilitación RDP: $RDPresult"
		$button_RDPEnable.Enabled = $true
	}

	# SUBBLOQUE: Deshabilitar escritorio remoto
	$button_RDPDisable_Click={
		$button_RDPDisable.Enabled = $false
		Get-ComputerTxtBox
		Add-Logs "$ComputerName - Disable RDP"
		Set-RDPDisable -ComputerName $ComputerName
		$button_RDPDisable.Enabled = $true
	}

	# SUBBLOQUE: Consultar actualizaciones instaladas
	$button_HotFix_Click={
		$button_HotFix.Enabled = $false
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Get the Windows Updates Installed"
		$result = Get-HotFix -ComputerName $ComputerName | Sort-Object InstalledOn | Format-Table __SERVER, Description, HotFixID, InstalledBy, InstalledOn,Caption -AutoSize | Out-String -Width $richtextbox_output.Width
		Add-RichTextBox $result
		$button_HotFix.Enabled = $true
	}

	# SUBBLOQUE: Listar aplicaciones instaladas en el equipo remoto (HTML report)
	$buttonApplications_Click = {
		$buttonApplications.Enabled = $false
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Aplicaciones Instaladas"
		try {
			$apps = Get-InstalledSoftware -ComputerName $ComputerName | Sort-Object Name
			if ($apps) {
				# Función de escape HTML sin dependencia de System.Web
				function HtmlEnc($s) {
					if ($null -eq $s) { return "" }
					[string]$s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
				}
				# Construir tabla HTML con filtro y estilos
				$rows = ""
				foreach ($a in $apps) {
					$name    = HtmlEnc $a.Name
					$version = HtmlEnc $a.Version
					$vendor  = HtmlEnc $a.Vendor
					$rows += "<tr><td>$name</td><td>$version</td><td>$vendor</td></tr>`n"
				}
				$htmlApps = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Aplicaciones - $ComputerName</title>
  <style>
    body { font-family:'Segoe UI',Arial,sans-serif; margin:16px; background:#f5f5f5; }
    h1   { font-size:16px; background:#e0e0e0; padding:8px; border-radius:4px; }
    input[type=text] { width:340px; padding:5px 8px; margin-bottom:12px; border:1px solid #ccc; border-radius:3px; }
    table { width:100%; border-collapse:collapse; background:#fff; box-shadow:0 2px 4px rgba(0,0,0,.1); }
    th { background:#b8c3d5; color:#00008b; font-size:12px; padding:6px 8px; text-align:left; position:sticky; top:0; }
    td { padding:5px 8px; font-size:12px; border-bottom:1px solid #eee; }
    tr:nth-child(even) td { background:#f9f9f9; }
    tr:hover td { background:#e8f0fe; }
    .cnt { color:#555; font-size:11px; margin-bottom:6px; }
  </style>
</head>
<body>
  <h1>&#x1F4E6; Aplicaciones instaladas &mdash; $ComputerName ($($apps.Count) encontradas)</h1>
  <input type="text" id="filtro" onkeyup="filtrar()" placeholder="Filtrar por nombre, versión o fabricante..." autofocus>
  <table id="tabla">
    <thead><tr><th>Nombre</th><th>Versión</th><th>Fabricante</th></tr></thead>
    <tbody>
$rows
    </tbody>
  </table>
  <script>
    function filtrar(){
      var f=document.getElementById('filtro').value.toLowerCase();
      var rows=document.querySelectorAll('#tabla tbody tr');
      rows.forEach(function(r){r.style.display=r.innerText.toLowerCase().includes(f)?'':'none';});
    }
  </script>
</body>
</html>
"@
				$tmpFile = "C:\Temp\Apps_$($ComputerName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
				if (-not (Test-Path 'C:\Temp')) { New-Item -ItemType Directory -Path 'C:\Temp' -Force | Out-Null }
				$htmlApps | Out-File -FilePath $tmpFile -Encoding UTF8
				Start-Process "explorer.exe" -ArgumentList $tmpFile
				Add-Logs -text "✅ Informe de aplicaciones abierto: $tmpFile"
			} else {
				Add-RichTextBox "$ComputerName - No se encontraron aplicaciones instaladas"
			}
		} catch {
			Add-Logs -text "[ERROR] Aplicaciones: $($_.Exception.Message)"
			Add-RichTextBox "Error al obtener aplicaciones: $($_.Exception.Message)"
		}
		$buttonApplications.Enabled = $true
	}

	# SUBBLOQUE: Consultar impresoras instaladas (lanza ListPrinters.ps1)
	$button_Printers_Click = {
		$button_Printers.Enabled = $false
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Impresoras (ListPrinters)"
		$scriptPS1 = Join-Path -Path $Global:ScriptRoot -ChildPath "scripts\ListPrinters.ps1"
		$remoteComputerName = if ([string]::IsNullOrWhiteSpace($textbox_computername.Text.Trim())) { $env:COMPUTERNAME } else { $textbox_computername.Text.Trim() }
		if (-not (Test-Path -Path $scriptPS1)) {
			Add-Logs -text "ERROR: No se encontró: $scriptPS1"
			[System.Windows.Forms.MessageBox]::Show("No se encontró el script:`n$scriptPS1", "Error", "OK", "Error")
			$button_Printers.Enabled = $true
		} else {
			try {
				$arguments = "-NoExit -NoProfile -ExecutionPolicy Bypass -File `"$scriptPS1`" -ComputerName `"$remoteComputerName`""
				Start-Process -FilePath "powershell.exe" -ArgumentList $arguments
				Add-Logs -text "✅ ListPrinters lanzado para $remoteComputerName"
				$button_Printers.Enabled = $true
			} catch {
				Add-Logs -text "ERROR al lanzar ListPrinters: $($_.Exception.Message)"
				$button_Printers.Enabled = $true
			}
		}
	}

	# SUBBLOQUE: Reiniciar equipo remotamente
	$button_Restart_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Reiniciar Equipo"
		$PsExecPath = "$ToolsFolder\psexec.exe"
		$command = "shutdown /r /t 0 /f"
		& $PsExecPath "\\$ComputerName" -s cmd /c $command
		Show-MsgBox -Prompt "$ComputerName - Reiniciando..." -Title "$ComputerName - Reinicio Equipo" -Icon Information -BoxType OKOnly
	}

	# SUBBLOQUE: Encender equipo mediante Wake-On-LAN
	$button_PowerOn_Click={
		Get-ComputerTxtBox
		if ([string]::IsNullOrEmpty($ComputerName)) {
			[System.Windows.Forms.MessageBox]::Show("Por favor, introduce un nombre de equipo", "WOL", "OK", "Warning")
			return
		}
		
		Add-Logs -text "$ComputerName - Encender Equipo (WOL)"
		
		try {
			Invoke-WakeOnLan -ComputerName $ComputerName
		} catch {
			Add-Logs -text "[ERROR] $($_.Exception.Message)"
			[System.Windows.Forms.MessageBox]::Show("Error al ejecutar WOL: $($_.Exception.Message)", "Error WOL", "OK", "Error")
		}
	}

	# SUBBLOQUE: Apagar equipo remotamente
	$button_Shutdown_Click={
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Apagado Equipo"
		$PsExecPath = "$ToolsFolder\psexec.exe"
		$command = "shutdown /s /t 0 /f"
		& $PsExecPath "\\$ComputerName" -s cmd /c $command
		Show-MsgBox -Prompt "$ComputerName - Apagando..." -Title "$ComputerName - Apagado Equipo" -Icon Information -BoxType OKOnly
	}

	# SUBBLOQUE: Obtener línea de comandos de procesos
	$buttonCommandLine_Click={
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Get the CommandLine Argument for each process"
		$result = Get-WmiObject Win32_Process -ComputerName $ComputerName | select-Object Name,ProcessID,CommandLine | Format-Table -AutoSize | Out-String -Width $richtextbox_output.Width
		Add-RichTextBox $result
	}

	# SUBBLOQUE: Obtener descripción del equipo
	$button_ComputerDescriptionQuery_Click={
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Get the Computer Description"
		$result = Get-ComputerComment -ComputerName $ComputerName
		Add-RichTextBox $result
	}

	# SUBBLOQUE: Cambiar descripción del equipo
	$button_ComputerDescriptionChange_Click = {
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Set the Computer Description"
		$Description = Show-Inputbox -message "Introduce una descripción para el equipo" -title "$ComputerName" -default "<Rol> - <Propietario> - <Ticket#>"
		if ($Description -ne "") {
			$result = Set-ComputerComment -ComputerName $ComputerName -Description $Description
			if ($result -eq $true) {
				Add-Logs -text "$ComputerName - Descripción actualizada con éxito: $Description"
				Show-Messagebox -message "Descripción actualizada correctamente." -title "Éxito"
			} else {
				Add-Logs -text "$ComputerName - Error al actualizar la descripción."
				Show-Messagebox -message "No se pudo actualizar la descripción." -title "Error"
			}
		} else {
			Add-Logs -text "$ComputerName - Operación cancelada por el usuario."
		}
	}

	# SUBBLOQUE: Abrir unidad C$ en Explorer++
	$buttonC_Click = {
		$ComputerName = $textbox_computername.Text.Trim()
	
		if ([string]::IsNullOrWhiteSpace($ComputerName)) {
			Add-Logs -text "Error: El nombre del equipo no es válido o está vacío."
			[System.Windows.Forms.MessageBox]::Show("Por favor, ingrese un nombre de equipo válido.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return
		}
	
		$ExplorerPath = Join-Path -Path $Global:ScriptRoot -ChildPath 'explorer++.exe'
		$TargetPath = "\\$ComputerName\c$"
	
		if (-not (Test-Path $ExplorerPath)) {
			Add-Logs -text "Error: Explorer++ no encontrado en la ruta $ExplorerPath."
			[System.Windows.Forms.MessageBox]::Show("El archivo explorer++.exe no se encuentra en la ruta especificada.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
			return
		}
	
		if (-not (Test-Path $TargetPath)) {
			Add-Logs -text "$ComputerName - No hay acceso a $TargetPath. Se requiere autenticación previa."
			[System.Windows.Forms.MessageBox]::Show("No hay acceso a $TargetPath. Introduce credenciales manualmente (Ej: Win + R > \\$ComputerName\c$)", "Acceso denegado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
			return
		}
	
		try {
			Start-Process -FilePath $ExplorerPath -ArgumentList "`"$TargetPath`""
			Add-Logs -text "$ComputerName - Explorer++ lanzado con ruta $TargetPath"
		} catch {
			Add-Logs -text "Error al lanzar Explorer++: $_"
			[System.Windows.Forms.MessageBox]::Show("Error al intentar ejecutar Explorer++ con la ruta especificada.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
		}
	}
		
	# SUBBLOQUE: Conexión VNC con credencial cifrada y validación previa
	$buttonRemoteAssistance_Click = {
		if ([string]::IsNullOrEmpty($env:ContrasenaGuardada)) {
			$rutaTempPass = Join-Path $ScriptRoot 'temp.pass'
			# Derivar clave identica a Launcher_RNCG.ps1 (Machine GUID + PBKDF2)
			$_pkGuid   = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Cryptography' -Name MachineGuid).MachineGuid
			$_pkSalt   = [System.Text.Encoding]::UTF8.GetBytes('NRC_Pass_v6_Salt_2026')
			$_pkDerive = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($_pkGuid, $_pkSalt, 10000)
			[byte[]] $Key = $_pkDerive.GetBytes(32)
			$_pkDerive.Dispose()
			if (Test-Path $rutaTempPass) {
				try {
					$securePass = Get-Content $rutaTempPass | ConvertTo-SecureString -Key $Key
					$env:ContrasenaGuardada = ConvertFrom-SecureString $securePass
				} catch {
					[System.Windows.Forms.MessageBox]::Show("Error al leer o descifrar la contraseña. Se solicitará manualmente.`n$($_.Exception.Message)")
				} finally {
					try { Remove-Item $rutaTempPass -Force } catch {}
				}
			}
		}
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Conexión VNC"
		if ([string]::IsNullOrEmpty($env:ContrasenaGuardada)) {
			$ButtonGuardar = [Windows.Forms.Button]@{
				Text = 'Guardar'; Location = [Drawing.Point]::new(130, 120); Width = 100; Height = 50
			}
			$TextBoxContrasena = [Windows.Forms.TextBox]@{
				Location = [Drawing.Point]::new(30, 80); Width = 320; PasswordChar = "*"
			}
			$LblintroduceCont = [Windows.Forms.Label]@{
				Text = 'Introduce una contraseña:'; Location = [Drawing.Point]::new(30, 30); AutoSize = $false; Width = 300; Height = 20
			}
			$formularioContrasena = New-Object Windows.Forms.Form
			$formularioContrasena.Text = "Conexión"
			$formularioContrasena.Width = 400
			$formularioContrasena.Height = 250
			$formularioContrasena.Controls.AddRange(@($LblintroduceCont, $TextBoxContrasena, $ButtonGuardar))
			$ButtonGuardar.Add_Click({
				if ([string]::IsNullOrWhiteSpace($TextBoxContrasena.text)) {
					[System.Windows.Forms.MessageBox]::Show("La contraseña no puede estar vacía.")
				} else {
					$contrasenaSegura = ConvertTo-SecureString $TextBoxContrasena.text -AsPlainText -Force
					$env:ContrasenaGuardada = ConvertFrom-SecureString $contrasenaSegura
					[System.Windows.Forms.MessageBox]::Show("Contraseña guardada para la sesión actual.")
					$formularioContrasena.Close()
				}
			})
			$formularioContrasena.ShowDialog()
		}
		$puertoVNC = 5700
		$contrasenaDescencriptada = ConvertTo-SecureString $env:ContrasenaGuardada
		$Marshal = [System.Runtime.InteropServices.Marshal]
		$Bstr = $Marshal::SecureStringToBSTR($contrasenaDescencriptada)
		$contrasena = $Marshal::PtrToStringAuto($Bstr)
		$configFileVNC = Join-Path $ScriptRoot 'vnc\options.vnc'
		$pathvncViwer = Join-Path $ScriptRoot 'vnc\vncviewer.exe'
		&$pathvncViwer "$($ComputerName):$puertoVNC" -Config $configFileVNC -User $env:USERNAME -Password $contrasena
		
		# Lanzar recogida de datos solo si el boton no esta ocupado
		if ($button_Check.Enabled) {
			$button_Check.PerformClick()
		} else {
			Add-logs -text "⚠️ Hay un proceso en curso. Conexión VNC iniciada sin recogida de datos."
		}
	}

	# SUBBLOQUE: Mostrar línea de comandos en GridView
	$buttonCommandLineGridView_Click={
		Get-ComputerTxtBox
		Add-Logs -text "$ComputerName - Get the CommandLine Argument for each process - Grid View"
		Get-WmiObject Win32_Process -ComputerName $ComputerName | select-Object Name,ProcessID,CommandLine| Out-GridView
	}

	# SUBBLOQUE: Abrir consola de servicios (MMC)
	$buttonServices_Click={
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Services MMC (services.msc /computer:$ComputerName)"
		$command = "services.msc"
		$arguments = "/computer:$computername"
		Start-Process $command $arguments 
	}

	# SUBBLOQUE: Abrir visor de eventos (MMC)
	$buttonEventVwr_Click={
		Get-ComputerTxtBox
		Add-logs -text "$ComputerName - Event Viewer MMC (eventvwr $Computername)"
		$command="eventvwr"
		$arguments = "$ComputerName"
		Start-Process $command $arguments
	}

	#==================================================================
	# BLOQUE: Herramientas Especiales y Scripts Externos
	#==================================================================
	#==================================================================
	# SUBBLOQUE: Cerrar Sesión Remota - Detectar usuario activo y cerrar sesión
	#==================================================================
	$buttonSendCommand_Click = {
		$ComputerName = $textbox_computername.Text.Trim()
		if ([string]::IsNullOrWhiteSpace($ComputerName)) {
			Add-Logs -text "❌ No hay equipo cargado. Por favor, ingresa un nombre de equipo primero."
			return
		}
		Add-Logs -text "$ComputerName - Comprobando el usuario activo en el equipo remoto"
		try {
			$sessionInfo = (Get-WmiObject -Query "SELECT * FROM Win32_ComputerSystem" -ComputerName $ComputerName | Select-Object -ExpandProperty UserName)
			if ([string]::IsNullOrWhiteSpace($sessionInfo)) {
				Add-Logs -text "$ComputerName - No se detectó ningún usuario logueado."
				[System.Windows.Forms.MessageBox]::Show("No hay un usuario activo en $ComputerName.", "Usuario Activo", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
				return
			}
			$loggedInUser = $sessionInfo
			Add-Logs -text "$ComputerName - Usuario activo detectado: $loggedInUser"
			$confirmClose = [System.Windows.Forms.MessageBox]::Show(
				"El usuario activo en $ComputerName es: $loggedInUser.`n¿Desea proceder con el cierre de sesión?",
				"Confirmar cierre de sesión",
				[System.Windows.Forms.MessageBoxButtons]::YesNo,
				[System.Windows.Forms.MessageBoxIcon]::Question)
			if ($confirmClose -eq [System.Windows.Forms.DialogResult]::Yes) {
				try {
					$os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
					$result = $os.Win32Shutdown(0)
					if ($result.ReturnValue -eq 0) {
						Add-Logs -text "$ComputerName - Cierre de sesión enviado correctamente al usuario $loggedInUser"
						[System.Windows.Forms.MessageBox]::Show("Cierre de sesión iniciado para el usuario $loggedInUser en $ComputerName.", "Cierre de Sesión", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
					} else {
						Add-Logs -text "$ComputerName - Error al cerrar sesión. Código: $($result.ReturnValue)"
						[System.Windows.Forms.MessageBox]::Show("Error al cerrar sesión. Código de error: $($result.ReturnValue)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
					}
				} catch {
					$errorMsg = $_.Exception.Message
					Add-Logs -text "$ComputerName - Error al cerrar sesión: $errorMsg"
					[System.Windows.Forms.MessageBox]::Show("Error al cerrar sesión: $errorMsg", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
				}
			} else {
				Add-Logs -text "$ComputerName - Cierre de sesión cancelado para el usuario $loggedInUser"
			}
		} catch {
			$errorMensaje = $_.Exception.Message
			Add-Logs -text "$ComputerName - Error al obtener la sesión del usuario activo: $errorMensaje"
			[System.Windows.Forms.MessageBox]::Show("Error al obtener la sesión del usuario activo en $ComputerName.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
		}
	}
	# BLOQUE: Scripts externos  gestionado por sections\Scripts.psm1 (Initialize-ScriptsMenu)

	#==================================================================
	# SUBBLOQUE: Ejecutar SystemInfo - Ventana CMD con información del sistema remoto
	#==================================================================
	$button_SystemInfoexe_Click = {
		# Deshabilita el botón mientras se ejecuta el comando
		$button_SystemInfoexe.Enabled = $false

		# Obtiene el nombre del equipo desde el TextBox
		Get-ComputerTxtBox

		# Comando principal y argumentos para abrir systeminfo
		$SystemInfo_cmd_command = "cmd"
		$SystemInfo_cmd_Args = "/k systeminfo /s $Computername"

		# Ejecuta el proceso en una nueva ventana CMD
		Start-Process $SystemInfo_cmd_command $SystemInfo_cmd_Args -WorkingDirectory "c:\"

		# Reactiva el botón
		$button_SystemInfoexe.Enabled = $true
	}

	#==================================================================
	# SUBBLOQUE: Evento Enter en TextBox - Ejecuta Ping al presionar Enter
	#==================================================================
	$textbox_computername_KeyPress = [System.Windows.Forms.KeyPressEventHandler]{
		# Si se presiona Enter (código 13), lanza el ping
		if ($_.KeyChar -eq 13){
			$button_ping.PerformClick()
			$richtextbox_output.Focus()
		}
	}
	#==================================================================
	# SUBBLOQUE: Eventos Generados - Inicialización, almacenamiento y limpieza de eventos
	#==================================================================

	# Evento LOAD del formulario principal: corrige el estado inicial de la ventana
	$Form_StateCorrection_Load = {
		$form_MainForm.WindowState = $InitialFormWindowState
	}

	# Evento CLOSING del formulario: guarda los valores actuales de los controles para persistencia
	$Form_StoreValues_Closing = {
		$script:MainForm_richtextbox_output = $richtextbox_output.Text
		$script:MainForm_textbox_processName = $textbox_processName.Text
		$script:MainForm_textbox_servicesAction = $textbox_servicesAction.Text
		$script:MainForm_textbox_networktracertparam = $textbox_networktracertparam.Text
		$script:MainForm_textbox_networkpathpingparam = $textbox_networkpathpingparam.Text
		$script:MainForm_textbox_pingparam = $textbox_pingparam.Text
		$script:MainForm_textbox_computername = $textbox_computername.Text
		$script:MainForm_richtextbox_Logs = $richtextbox_Logs.Text
	}
	# Evento FORMCLOSED: elimina todos los manejadores de eventos para liberar memoria y evitar errores
	$Form_Cleanup_FormClosed = {
		# Limpiar recolección de datos en streaming si está en curso
		try { if (Get-Command 'Cleanup-StreamingCollection' -ErrorAction SilentlyContinue) { Cleanup-StreamingCollection } } catch {}
		try {
			# (Se omiten comentarios individuales por cantidad masiva de controles, pero todos siguen el patrón):
			# control.remove_eventhandler(event_handler)
			$richtextbox_output.remove_TextChanged($richtextbox_output_TextChanged)
			$button_formExit.remove_Click($button_formExit_Click)
			$button_outputClear.remove_Click($button_outputClear_Click)
			$button_ExportRTF.remove_Click($button_ExportRTF_Click)
			$button_outputCopy.remove_Click($button_outputCopy_Click)
			$buttonSendCommand.remove_Click($buttonSendCommand_Click)
			$button_mmcCompmgmt.remove_Click($button_mmcCompmgmt_Click)
			$buttonServices.remove_Click($buttonServices_Click)
			$buttonEventVwr.remove_Click($buttonEventVwr_Click)
			$button_GPupdate.remove_Click($button_GPupdate_Click)
			$button_ping.remove_Click($button_ping_Click)
			$button_remot.remove_Click($button_remot_Click)
			$buttonRemoteAssistance.remove_Click($buttonRemoteAssistance_Click)
			$button_PsRemoting.remove_Click($button_PsRemoting_Click)
			$buttonC.remove_Click($buttonC_Click)
			$button_networkconfig.remove_Click($button_networkIPConfig_Click)
			$button_Restart.remove_Click($button_Restart_Click)
			$button_PowerOn.remove_Click($button_PowerOn_Click)
			$button_Shutdown.remove_Click($button_Shutdown_Click)
			$button_UsersGroupLocalUsers.remove_Click($button_UsersGroupLocalUsers_Click)
			$button_UsersGroupLocalGroups.remove_Click($button_UsersGroupLocalGroups_Click)
			$button_ComputerDescriptionChange.remove_Click($button_ComputerDescriptionChange_Click)
			$button_ComputerDescriptionQuery.remove_Click($button_ComputerDescriptionQuery_Click)
			$button_HotFix.remove_Click($button_HotFix_Click)
			$button_RDPDisable.remove_Click($button_RDPDisable_Click)
			$button_RDPEnable.remove_Click($button_RDPEnable_Click)
			$button_ActivarWinRM_SO.remove_Click($button_ActivarWinRM_Click)
			$button_DeshabilitarWinRM_SO.remove_Click($button_DeshabilitarWinRM_Click)
			$button_PSRemotoGeneral.remove_Click($button_PowershellRemota_Click)
			$button_PageFile.remove_Click($button_PageFile_Click)
			$button_StartupCommand.remove_Click($button_StartupCommand_Click)
			$buttonApplications.remove_Click($buttonApplications_Click)
			$button_MotherBoard.remove_Click($button_MotherBoard_Click)
			$button_Processor.remove_Click($button_Processor_Click)
			$button_Memory.remove_Click($button_Memory_Click)
			$button_SystemType.remove_Click($button_SystemType_Click)
			$button_Printers.remove_Click($button_Printers_Click)
			$button_USBDevices.remove_Click($button_USBDevices_Click)
			$button_ConnectivityTesting.remove_Click($button_ConnectivityTesting_Click)
			$button_NIC.remove_Click($button_NIC_Click)
			$button_networkIPConfig.remove_Click($button_networkIPConfig_Click)
			$button_networkTestPort.remove_Click($button_networkTestPort_Click)
			$button_networkRouteTable.remove_Click($button_networkRouteTable_Click)
			$buttonCommandLineGridView.remove_Click($buttonCommandLineGridView_Click)
			$button_processAll.remove_Click($button_processAll_Click)
			$buttonCommandLine.remove_Click($buttonCommandLine_Click)
			$button_processTerminate.remove_Click($button_processTerminate_Click)
			$button_process100MB.remove_Click($button_process100MB_Click)
			$button_ProcessGrid.remove_Click($button_ProcessGrid_Click)
			$button_processOwners.remove_Click($button_processOwners_Click)
			$button_processLastHour.remove_Click($button_processLastHour_Click)
			$button_servicesNonStandardUser.remove_Click($button_servicesNonStandardUser_Click)
			$button_mmcServices.remove_Click($button_mmcServices_Click)
			$button_servicesAutoNotStarted.remove_Click($button_servicesAutoNotStarted_Click)
			$textbox_servicesAction.remove_Click($textbox_servicesAction_Click)
			$button_servicesRestart.remove_Click($button_servicesRestart_Click)
			$button_servicesQuery.remove_Click($button_servicesQuery_Click)
			$button_servicesStart.remove_Click($button_servicesStart_Click)
			$button_servicesStop.remove_Click($button_servicesStop_Click)
			$button_servicesRunning.remove_Click($button_servicesRunning_Click)
			$button_servicesAll.remove_Click($button_servicesAll_Click)
			$button_servicesGridView.remove_Click($button_servicesGridView_Click)
			$button_servicesAutomatic.remove_Click($button_servicesAutomatic_Click)
			$button_DiskUsage.remove_Click($button_DiskUsage_Click)
			$button_DiskPartition.remove_Click($button_DiskPartition_Click)
			$button_DiskLogical.remove_Click($button_DiskLogical_Click)
			$button_DiskMountPoint.remove_Click($button_DiskMountPoint_Click)
			$button_DiskRelationship.remove_Click($button_DiskRelationship_Click)
			$button_DiskMappedDrive.remove_Click($button_DiskMappedDrive_Click)
			$button_mmcShares.remove_Click($button_mmcShares_Click)
			$button_SharesGrid.remove_Click($button_SharesGrid_Click)
			$button_Shares.remove_Click($button_Shares_Click)
			$button_RebootHistory.remove_Click($button_RebootHistory_Click)
			$button_mmcEvents.remove_Click($button_mmcEvents_Click)
			$button_EventsLogNames.remove_Click($button_EventsLogNames_Click)
			$button_Rwinsta.remove_Click($button_Rwinsta_Click)
			$button_Qwinsta.remove_Click($button_Qwinsta_Click)
			$button_MsInfo32.remove_Click($button_MsInfo32_Click)
			$button_DriverQuery.remove_Click($button_DriverQuery_Click)
			$button_SystemInfoexe.remove_Click($button_SystemInfoexe_Click)
			$button_PAExec.remove_Click($button_PAExec_Click)
			$button_psexec.remove_Click($button_psexec_Click)
			$button_networkTracert.remove_Click($button_networkTracert_Click)
			$button_networkNsLookup.remove_Click($button_networkNsLookup_Click)
			$button_networkPing.remove_Click($button_networkPing_Click)
			$button_networkPathPing.remove_Click($button_networkPathPing_Click)
			$textbox_computername.remove_TextChanged($textbox_computername_TextChanged)
			$textbox_computername.remove_KeyPress($textbox_computername_KeyPress)
			$button_Check.remove_Click($button_Check_Click)
			$richtextbox_Logs.remove_TextChanged($richtextbox_Logs_TextChanged)
			$form_MainForm.remove_Load($OnLoadFormEvent)
			$ToolStripMenuItem_CommandPrompt.remove_Click($ToolStripMenuItem_CommandPrompt_Click)
			$ToolStripMenuItem_Powershell.remove_Click($ToolStripMenuItem_Powershell_Click)
				$ToolStripMenuItem_compmgmt.remove_Click($ToolStripMenuItem_compmgmt_Click)
			$ToolStripMenuItem_taskManager.remove_Click($ToolStripMenuItem_taskManager_Click)
			$ToolStripMenuItem_services.remove_Click($ToolStripMenuItem_services_Click)
			$ToolStripMenuItem_regedit.remove_Click($ToolStripMenuItem_regedit_Click)
			$ToolStripMenuItem_mmc.remove_Click($ToolStripMenuItem_mmc_Click)
				$ToolStripMenuItem_AboutInfo.remove_Click($ToolStripMenuItem_AboutInfo_Click)
			$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Ping.remove_Click($button_ping_Click)
			$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_RDP.remove_Click($button_remot_Click)
			$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_compmgmt.remove_Click($button_mmcCompmgmt_Click)
			$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_services.remove_Click($button_mmcServices_Click)
			$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_eventvwr.remove_Click($button_mmcEvents_Click)
			$ToolStripMenuItem_InternetExplorer.remove_Click($ToolStripMenuItem_InternetExplorer_Click)
			$ToolStripMenuItem_TerminalAdmin.remove_Click($ToolStripMenuItem_TerminalAdmin_Click)
			$ToolStripMenuItem_ADSearchDialog.remove_Click($ToolStripMenuItem_ADSearchDialog_Click)
			$ToolStripMenuItem_ADPrinters.remove_Click($ToolStripMenuItem_ADPrinters_Click)
			$ToolStripMenuItem_DHCP.remove_Click($ToolStripMenuItem_DHCP_Click)
			$ToolStripMenuItem_systemInformationMSinfo32exe.remove_Click($ToolStripMenuItem_systemInformationMSinfo32exe_Click)
			$ToolStripMenuItem_netstatsListening.remove_Click($ToolStripMenuItem_netstatsListening_Click)
			$ToolStripMenuItem_registeredSnappins.remove_Click($ToolStripMenuItem_registeredSnappins_Click)
			$ToolStripMenuItem_certificateManager.remove_Click($ToolStripMenuItem_certificateManager_Click)
			$ToolStripMenuItem_devicemanager.remove_Click($ToolStripMenuItem_devicemanager_Click)
			$ToolStripMenuItem_systemproperties.remove_Click($ToolStripMenuItem_systemproperties_Click)
			$ToolStripMenuItem_sharedFolders.remove_Click($ToolStripMenuItem_sharedFolders_Click)
			$ToolStripMenuItem_performanceMonitor.remove_Click($ToolStripMenuItem_performanceMonitor_Click)
			$ToolStripMenuItem_groupPolicyEditor.remove_Click($ToolStripMenuItem_groupPolicyEditor_Click)
			$ToolStripMenuItem_localUsersAndGroups.remove_Click($ToolStripMenuItem_localUsersAndGroups_Click)
			$ToolStripMenuItem_diskManagement.remove_Click($ToolStripMenuItem_diskManagement_Click)
			$ToolStripMenuItem_localSecuritySettings.remove_Click($ToolStripMenuItem_localSecuritySettings_Click)
			$ToolStripMenuItem_scheduledTasks.remove_Click($ToolStripMenuItem_scheduledTasks_Click)
			$ToolStripMenuItem_PowershellISE.remove_Click($ToolStripMenuItem_PowershellISE_Click)
			$ToolStripMenuItem_adExplorer.remove_Click($ToolStripMenuItem_adExplorer_Click)
			$ToolStripMenuItem_resetCredenciaisVNC.remove_Click($ToolStripMenuItem_resetCredenciaisVNC_Click)
			$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Qwinsta.remove_Click($button_Qwinsta_Click)
			$ToolStripMenuItem_rwinsta.remove_Click($button_Rwinsta_Click)
			$ToolStripMenuItem_GeneratePassword.remove_Click($button_PasswordGen_Click)
			$timerCheckJob.remove_Tick($timerCheckJob_Tick2)
			$form_MainForm.remove_Load($Form_StateCorrection_Load)
			$form_MainForm.remove_Closing($Form_StoreValues_Closing)
			$form_MainForm.remove_FormClosed($Form_Cleanup_FormClosed)
		} catch [Exception] {}
	}
	#==================================================================
	# BBLOQUE: BINARY DATA & BOTONES & ETIQUETAS (FRONTEND INTERFAZ)
	#==================================================================
	#==================================================================
	# SUBBLOQUE: Código Generado del Formulario - Construcción visual y propiedades
	#==================================================================
	# Configuración de controles del formulario principal (form_MainForm)
	$form_MainForm.Controls.Add($panel_ContentArea)
	$form_MainForm.Controls.Add($panel_RTBButtons)
	$form_MainForm.Controls.Add($tabcontrol_computer)
	$form_MainForm.Controls.Add($groupbox_ComputerName)
	$form_MainForm.Controls.Add($richtextbox_Logs)
	$form_MainForm.Controls.Add($statusbar1)
	$form_MainForm.Controls.Add($menustrip_principal)
	$form_MainForm.AutoScaleMode = 'Inherit'
	$form_MainForm.AutoSize = $False
	$form_MainForm.BackColor = 'Control'
	$form_MainForm.ClientSize = '1324, 719'
	$form_MainForm.Font = "Calibri, 8.25pt"
	$form_MainForm.MainMenuStrip = $menustrip_principal
	$form_MainForm.MinimumSize = '1332, 746'
	$form_MainForm.Name = "form_MainForm"
	$form_MainForm.Text = "LazyWinAdmin"
	$form_MainForm.add_Load($OnLoadFormEvent)
	#==================================================================
	# SUBBLOQUE: RichTextBox Principal de Salida - richtextbox_output
	#==================================================================
	$richtextbox_output.Dock = 'Fill'
	$richtextbox_output.Font = "Calibri, 10pt"
	$richtextbox_output.Location = '0, 224'
	$richtextbox_output.Name = "richtextbox_output"
	$richtextbox_output.Size = '1170, 365'
	$richtextbox_output.TabIndex = 3
	$richtextbox_output.Text = ""
	$tooltipinfo.SetToolTip($richtextbox_output, "Output")
	$richtextbox_output.add_TextChanged($richtextbox_output_TextChanged)
	#==================================================================
	# SUBBLOQUE: Panel Contenedor Principal - panel_ContentArea
	# Contiene el RichTextBox (Fill) y el Panel Utilidades (Right)
	#==================================================================
	# Docking order: index 0 = Fill (laid out last), index 1+ = edge (laid out first)
	$panel_ContentArea.Controls.Add($richtextbox_output)
	$panel_ContentArea.Controls.Add($panel_Utilities)
	$panel_ContentArea.Dock = 'Fill'
	$panel_ContentArea.Name = "panel_ContentArea"
	#==================================================================
	# SUBBLOQUE: Panel Utilidades - panel_Utilities (lateral derecho permanente)
	#==================================================================
	# Estructura: [panel_PKHeader (Top)] + [panel_PKButtons (Fill)]
	$panel_Utilities.Controls.Add($panel_PKButtons)
	$panel_Utilities.Controls.Add($panel_PKHeader)
	$panel_Utilities.Dock = 'Right'
	$panel_Utilities.Width = 360
	$panel_Utilities.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 240)
	$panel_Utilities.BorderStyle = 'FixedSingle'
	$panel_Utilities.Name = "panel_Utilities"
	#==================================================================
	# SUBBLOQUE: Cabecera del panel Pass Keeper
	# Contiene: label_UtilsTitle (Fill) + button_PK_Add (Right)
	#==================================================================
	$panel_PKHeader.Controls.Add($label_UtilsTitle)
	$panel_PKHeader.Controls.Add($button_PK_Add)
	$panel_PKHeader.Dock = 'Top'
	$panel_PKHeader.Height = 32
	$panel_PKHeader.BackColor = [System.Drawing.Color]::FromArgb(210, 215, 230)
	$panel_PKHeader.Name = "panel_PKHeader"
	#==================================================================
	# SUBBLOQUE: Título del Panel Pass Keeper
	#==================================================================
	$label_UtilsTitle.Dock = 'Fill'
	$label_UtilsTitle.Text = "Pass Keeper"
	$label_UtilsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
	$label_UtilsTitle.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 80)
	$label_UtilsTitle.TextAlign = 'MiddleCenter'
	$label_UtilsTitle.Name = "label_UtilsTitle"
	#==================================================================
	# SUBBLOQUE: Botón Añadir entrada Pass Keeper
	#==================================================================
	$button_PK_Add.Text = [char]0x2699  # ⚙
	$button_PK_Add.Dock = 'Right'
	$button_PK_Add.Width = 30
	$button_PK_Add.FlatStyle = 'Flat'
	$button_PK_Add.FlatAppearance.BorderSize = 0
	$button_PK_Add.BackColor = [System.Drawing.Color]::Transparent
	$button_PK_Add.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 80)
	$button_PK_Add.Font = New-Object System.Drawing.Font("Segoe UI", 11)
	$button_PK_Add.Cursor = [System.Windows.Forms.Cursors]::Hand
	$button_PK_Add.Name = "button_PK_Add"
	$tooltipinfo.SetToolTip($button_PK_Add, "Añadir nueva entrada al Pass Keeper")
	$button_PK_Add.Add_Click({ if (Get-Command 'Show-AddPKDialog' -ErrorAction SilentlyContinue) { Show-AddPKDialog } })
	#==================================================================
	# SUBBLOQUE: Panel scrollable de botones Pass Keeper
	#==================================================================
	$panel_PKButtons.Dock = 'Fill'
	$panel_PKButtons.AutoScroll = $true
	$panel_PKButtons.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 240)
	$panel_PKButtons.Name = "panel_PKButtons"
	#==================================================================
	# SUBBLOQUE: Panel Inferior con Botones - panel_RTBButtons
	#==================================================================
	$panel_RTBButtons.Controls.Add($button_formExit)
	$panel_RTBButtons.Controls.Add($button_outputClear)
	$panel_RTBButtons.Controls.Add($button_ExportRTF)
	$panel_RTBButtons.Controls.Add($button_outputCopy)
	$panel_RTBButtons.Dock = 'Bottom'
	$panel_RTBButtons.Location = '0, 589'
	$panel_RTBButtons.Name = "panel_RTBButtons"
	$panel_RTBButtons.Size = '1244, 34'
	$panel_RTBButtons.TabIndex = 63
	#==================================================================
	# SUBBLOQUE: Botón Salir (Exit) - button_formExit
	#==================================================================
	$button_formExit.Dock = 'Right'
	$button_formExit.Font = "Calibri, 9.75pt, style=Bold"
	$button_formExit.ForeColor = 'Red'
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_formExit.ps1')
	$button_formExit.Location = '1129, 0'
	$button_formExit.Name = "button_formExit"
	$button_formExit.Size = '41, 34'
	$button_formExit.TabIndex = 15
	$tooltipinfo.SetToolTip($button_formExit, "Exit")
	$button_formExit.UseVisualStyleBackColor = $True
	$button_formExit.add_Click($button_formExit_Click)
	#==================================================================
	# SUBBLOQUE: Botón Limpiar Logs - button_outputClear
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_outputClear.ps1')
	$button_outputClear.Location = '1, 1'
	$button_outputClear.Name = "button_outputClear"
	$button_outputClear.Size = '38, 31'
	$button_outputClear.TabIndex = 5
	$tooltipinfo.SetToolTip($button_outputClear, "Limpa Logs")
	$button_outputClear.UseVisualStyleBackColor = $True
	$button_outputClear.add_Click($button_outputClear_Click)
	#==================================================================
	# SUBBLOQUE: Botón Exportar a Notepad - button_ExportRTF
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_ExportRTF.ps1')
	$button_ExportRTF.Location = '95, 1'
	$button_ExportRTF.Name = "button_ExportRTF"
	$button_ExportRTF.Size = '41, 31'
	$button_ExportRTF.TabIndex = 23
	$tooltipinfo.SetToolTip($button_ExportRTF, "Exporta a Notepad")
	$button_ExportRTF.UseVisualStyleBackColor = $True
	$button_ExportRTF.add_Click($button_ExportRTF_Click)
	#==================================================================
	# SUBBLOQUE: Botón Copiar al Portapapeles - button_outputCopy
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_outputCopy.ps1')
	$button_outputCopy.Location = '45, 1'
	$button_outputCopy.Name = "button_outputCopy"
	$button_outputCopy.Size = '44, 31'
	$button_outputCopy.TabIndex = 20
	$tooltipinfo.SetToolTip($button_outputCopy, "Copia Portapapeles")
	$button_outputCopy.UseVisualStyleBackColor = $True
	$button_outputCopy.add_Click($button_outputCopy_Click)
	#==================================================================
	# SUBBLOQUE: Pestañas principales del formulario - tabcontrol_computer y tabpage_general
	#==================================================================
	$tabcontrol_computer.Controls.Add($tabpage_general)
	$tabcontrol_computer.Controls.Add($tabpage_ComputerOSSystem)
	$tabcontrol_computer.Controls.Add($tabpage_network)
	$tabcontrol_computer.Controls.Add($tabpage_processes)
	$tabcontrol_computer.Controls.Add($tabpage_services)
	$tabcontrol_computer.Controls.Add($tabpage_diskdrives)
	$tabcontrol_computer.Controls.Add($tabpage_shares)
	$tabcontrol_computer.Controls.Add($tabpage_eventlog)
	$tabcontrol_computer.Controls.Add($tabpage_ExternalTools)
	$tabcontrol_computer.Dock = 'Top'
	$tabcontrol_computer.Location = '0, 87'
	$tabcontrol_computer.Multiline = $True
	$tabcontrol_computer.Name = "tabcontrol_computer"
	$tabcontrol_computer.SelectedIndex = 0
	$tabcontrol_computer.Size = '1244, 137'
	$tabcontrol_computer.TabIndex = 11
	#==================================================================
	# SUBBLOQUE: Tab General - Acciones y accesos directos generales
	#==================================================================
	$tabpage_general.Controls.Add($buttonSendCommand)
	$tabpage_general.Controls.Add($groupbox_ManagementConsole)
	$tabpage_general.Controls.Add($button_GPupdate)
	$tabpage_general.Controls.Add($button_PSRemotoGeneral)
	$tabpage_general.Controls.Add($button_ping)
	$tabpage_general.Controls.Add($button_remot)
	$tabpage_general.Controls.Add($buttonRemoteAssistance)
	$tabpage_general.Controls.Add($button_PsRemoting)
	$tabpage_general.Controls.Add($buttonC)
	$tabpage_general.Controls.Add($button_networkconfig)
	$tabpage_general.Controls.Add($button_Restart)
	$tabpage_general.Controls.Add($button_PowerOn)
	$tabpage_general.Controls.Add($button_Shutdown)
	$tabpage_general.BackColor = 'Control'
	$tabpage_general.Location = '4, 22'
	$tabpage_general.Name = "tabpage_general"
	$tabpage_general.Size = '1236, 111'
	$tabpage_general.TabIndex = 12
	$tabpage_general.Text = "General"
	# SUBBLOQUE: Botón Ejecutar Comando Remoto - buttonSendCommand
	#==================================================================
	$CerrarSesionIconPath = Join-Path $Global:ScriptRoot "icos\CerrarSesion.ico"
	if (Test-Path $CerrarSesionIconPath) {
		$CerrarSesionImg = [System.Drawing.Image]::FromFile($CerrarSesionIconPath)
		$CerrarSesionSize = New-Object System.Drawing.Size(44, 44)
		$buttonSendCommand.Image = New-Object System.Drawing.Bitmap($CerrarSesionImg, $CerrarSesionSize)
	}
	$buttonSendCommand.ImageAlign = 'TopCenter'
	$buttonSendCommand.Location = '599, 4'
	$buttonSendCommand.Name = "buttonSendCommand"
	$buttonSendCommand.Size = '66, 77'
	$buttonSendCommand.TabIndex = 50
	$buttonSendCommand.Text = "Cerrar Sesión"
	$buttonSendCommand.TextAlign = 'BottomCenter'
	$buttonSendCommand.UseVisualStyleBackColor = $True
	$buttonSendCommand.add_Click($buttonSendCommand_Click)
	#==================================================================
	# SUBBLOQUE: Grupo Consola de Administración - groupbox_ManagementConsole
	#==================================================================
	$groupbox_ManagementConsole.Controls.Add($button_mmcCompmgmt)
	$groupbox_ManagementConsole.Controls.Add($buttonServices)
	$groupbox_ManagementConsole.Controls.Add($buttonEventVwr)
	$groupbox_ManagementConsole.Controls.Add($RegistroRemoto)
	$groupbox_ManagementConsole.Location = '903, 4'
	$groupbox_ManagementConsole.Name = "groupbox_ManagementConsole"
	$groupbox_ManagementConsole.Size = '171, 78'
	$groupbox_ManagementConsole.TabIndex = 49
	$groupbox_ManagementConsole.TabStop = $False
	$groupbox_ManagementConsole.Text = "Consola Administración"
	#==================================================================
	# SUBBLOQUE: Botón Administración de Equipos - button_mmcCompmgmt
	#==================================================================
	$button_mmcCompmgmt.Font = "Trebuchet MS, 8.25pt"
	$button_mmcCompmgmt.ForeColor = 'ForestGreen'
	$button_mmcCompmgmt.Location = '2, 19'
	$button_mmcCompmgmt.Name = "button_mmcCompmgmt"
	$button_mmcCompmgmt.Size = '82, 23'
	$button_mmcCompmgmt.TabIndex = 7
	$button_mmcCompmgmt.Text = "AdminEquipos"
	$tooltipinfo.SetToolTip($button_mmcCompmgmt, "Lanza consola de Administración de Equipos")
	$button_mmcCompmgmt.UseVisualStyleBackColor = $True
	$button_mmcCompmgmt.add_Click($button_mmcCompmgmt_Click)
	#==================================================================
	# SUBBLOQUE: Botón Servicios - buttonServices
	#==================================================================
	$buttonServices.Font = "Trebuchet MS, 8.25pt"
	$buttonServices.ForeColor = 'ForestGreen'
	$buttonServices.Location = '84, 19'
	$buttonServices.Name = "buttonServices"
	$buttonServices.Size = '71, 23'
	$buttonServices.TabIndex = 45
	$buttonServices.Text = "Servicios"
	$tooltipinfo.SetToolTip($buttonServices, "Lanza Consola Servicios")
	$buttonServices.UseVisualStyleBackColor = $True
	$buttonServices.add_Click($buttonServices_Click)
	#==================================================================
	# SUBBLOQUE: Botón Visor de Eventos - buttonEventVwr
	#==================================================================
	$buttonEventVwr.Font = "Trebuchet MS, 8.25pt"
	$buttonEventVwr.ForeColor = 'ForestGreen'
	$buttonEventVwr.Location = '2, 48'
	$buttonEventVwr.Name = "buttonEventVwr"
	$buttonEventVwr.Size = '82, 23'
	$buttonEventVwr.TabIndex = 47
	$buttonEventVwr.Text = "Visor Eventos"
	$tooltipinfo.SetToolTip($buttonEventVwr, "Lanza Visor de Eventos")
	$buttonEventVwr.UseVisualStyleBackColor = $True
	$buttonEventVwr.add_Click($buttonEventVwr_Click)
	#==================================================================
	# SUBBLOQUE: Botón Registro Remoto - RegistroRemoto
	#==================================================================
	$RegistroRemoto.Font = "Trebuchet MS, 8.25pt"
	$RegistroRemoto.ForeColor = 'ForestGreen'
	$RegistroRemoto.Location = '84, 48'
	$RegistroRemoto.Name = "RegistroRemoto"
	$RegistroRemoto.Size = '71, 23'
	$RegistroRemoto.TabIndex = 47
	$RegistroRemoto.Text = "Registro"
	$tooltipinfo.SetToolTip($RegistroRemoto, "Abrir Registro Remoto")
	$RegistroRemoto.UseVisualStyleBackColor = $True
	$RegistroRemoto_Click = {
		try {
			Get-ComputerTxtBox
			if (($ComputerName -like "localhost") -or ($ComputerName -like ".") -or ($ComputerName -like "127.0.0.1") -or ($ComputerName -like "$env:computername")) {
				Add-Logs -text "Localhost - Editor de Registro local (regedit.exe)"
				Start-Process -FilePath "regedit.exe"
			} else {
				Add-Logs -text "$ComputerName - Conectando al Editor de Registro remoto"
				Start-Process -FilePath "regedit.exe"
				Start-Sleep -Seconds 1
				$shell = New-Object -ComObject "WScript.Shell"
				$shell.AppActivate("Editor del Registro")
				Start-Sleep -Milliseconds 500
				[System.Windows.Forms.SendKeys]::SendWait("%A")
				Start-Sleep -Milliseconds 500
				[System.Windows.Forms.SendKeys]::SendWait("C")
				Start-Sleep -Milliseconds 500
				[System.Windows.Forms.SendKeys]::SendWait("$ComputerName")
				Start-Sleep -Milliseconds 500
				[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
				Add-Logs -text "Conexión establecida con $ComputerName desde el Editor de Registro"
			}
		} catch {
			$errorMsg = $_.Exception.Message
			Add-Logs -text "Error al abrir el Editor de Registro: $errorMsg"
			[System.Windows.Forms.MessageBox]::Show("No se pudo abrir el Editor de Registro. Error: $errorMsg", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
		}
	}
	$RegistroRemoto.add_Click($RegistroRemoto_Click)
	#==================================================================
	# SUBBLOQUE: Botón Actualizar Directivas - button_GPupdate
	#==================================================================
	$button_GPupdate.Font = "Microsoft Sans Serif, 8.25pt"
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_GPupdate.ps1')
	$button_GPupdate.ImageAlign = 'TopCenter'
	$button_GPupdate.Location = '521, 4'
	$button_GPupdate.Name = "button_GPupdate"
	$button_GPupdate.Size = '74, 77'
	$button_GPupdate.TabIndex = 0
	$button_GPupdate.Text = "GPupdate"
	$button_GPupdate.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($button_GPupdate, "Actualiza Directiva de Grupo")
	$button_GPupdate.UseVisualStyleBackColor = $True
	$button_GPupdate.add_Click($button_GPupdate_Click)
	# SUBBLOQUE: Botón Powershell Remota (pestaña General) - button_PSRemotoGeneral
	#==================================================================
	$button_PSRemotoGeneral.Font = "Microsoft Sans Serif, 8.25pt"
	$PSRemotoIconPath = Join-Path $Global:ScriptRoot "icos\Powershell.ico"
	if (Test-Path $PSRemotoIconPath) {
		$PSRemotoImg = [System.Drawing.Image]::FromFile($PSRemotoIconPath)
		$PSRemotoSize = New-Object System.Drawing.Size(52, 52)
		$button_PSRemotoGeneral.Image = New-Object System.Drawing.Bitmap($PSRemotoImg, $PSRemotoSize)
	} else {
		$button_PSRemotoGeneral.Image = $button_PsRemoting.Image
	}
	$button_PSRemotoGeneral.ImageAlign = 'TopCenter'
	$button_PSRemotoGeneral.Location = '225, 4'
	$button_PSRemotoGeneral.Name = "button_PSRemotoGeneral"
	$button_PSRemotoGeneral.Size = '74, 77'
	$button_PSRemotoGeneral.TabIndex = 65
	$button_PSRemotoGeneral.Text = "PS Remota"
	$button_PSRemotoGeneral.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($button_PSRemotoGeneral, "Abre sesión PowerShell Remota")
	$button_PSRemotoGeneral.UseVisualStyleBackColor = $True
	$button_PSRemotoGeneral.add_Click($button_PowershellRemota_Click)
	#==================================================================
	# SUBBLOQUE: Botón Ping - button_ping
	#==================================================================
	$button_ping.Font = "Microsoft Sans Serif, 8.25pt"
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_ping.ps1')
	$button_ping.ImageAlign = 'TopCenter'
	$button_ping.Location = '3, 4'
	$button_ping.Name = "button_ping"
	$button_ping.Size = '74, 77'
	$button_ping.TabIndex = 0
	$button_ping.Text = "Ping"
	$button_ping.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($button_ping, "Lanza Ping")
	$button_ping.UseVisualStyleBackColor = $True
	$button_ping.add_Click($button_ping_Click)
	#==================================================================
	# SUBBLOQUE: Botón Escritorio Remoto - button_remot
	#==================================================================
	$button_remot.Font = "Microsoft Sans Serif, 8.25pt"
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_remot.ps1')
	$button_remot.ImageAlign = 'TopCenter'
	$button_remot.Location = '77, 4'
	$button_remot.Name = "button_remot"
	$button_remot.Size = '74, 77'
	$button_remot.TabIndex = 4
	$button_remot.Text = "RDP"
	$button_remot.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($button_remot, "Abre conexión a Escritorio Remoto")
	$button_remot.UseVisualStyleBackColor = $True
	$button_remot.add_Click($button_remot_Click)
	#==================================================================
	# SUBBLOQUE: Botón Asistencia Remota - buttonRemoteAssistance
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\buttonRemoteAssistance.ps1')
	$buttonRemoteAssistance.ImageAlign = 'TopCenter'
	$buttonRemoteAssistance.Location = '299, 4'
	$buttonRemoteAssistance.Name = "buttonRemoteAssistance"
	$buttonRemoteAssistance.Size = '74, 77'
	$buttonRemoteAssistance.TabIndex = 44
	$buttonRemoteAssistance.Text = "VNC"
	$buttonRemoteAssistance.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($buttonRemoteAssistance, "Conectarse por VNC al equipo")
	$buttonRemoteAssistance.UseVisualStyleBackColor = $True
	$buttonRemoteAssistance.add_Click($buttonRemoteAssistance_Click)
	#==================================================================
	# SUBBLOQUE: Botón PowerShell Remoting - button_PsRemoting
	#==================================================================
	$button_PsRemoting.Font = "Microsoft Sans Serif, 8.25pt"
	$CMDIconPath = Join-Path $Global:ScriptRoot "icos\CMD.ico"
	if (Test-Path $CMDIconPath) {
		$CMDImg = [System.Drawing.Image]::FromFile($CMDIconPath)
		$CMDSize = New-Object System.Drawing.Size(44, 44)
		$button_PsRemoting.Image = New-Object System.Drawing.Bitmap($CMDImg, $CMDSize)
	}
	$button_PsRemoting.ImageAlign = 'TopCenter'
	$button_PsRemoting.Location = '151, 4'
	$button_PsRemoting.Name = "button_PsRemoting"
	$button_PsRemoting.Size = '74, 77'
	$button_PsRemoting.TabIndex = 27
	$button_PsRemoting.Text = "CMD Remota"
	$button_PsRemoting.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($button_PsRemoting, "Abre una sesión de CMD Remota")
	$button_PsRemoting.UseVisualStyleBackColor = $True
	$button_PsRemoting.add_Click($button_PsRemoting_Click)
	#==================================================================
	# SUBBLOQUE: Botón Unidad C$ Remota - buttonC
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\buttonC.ps1')
	$buttonC.ImageAlign = 'TopCenter'
	$buttonC.Location = '447, 4'
	$buttonC.Name = "buttonC"
	$buttonC.Size = '74, 77'
	$buttonC.TabIndex = 43
	$buttonC.Text = "Explorer++"
	$buttonC.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($buttonC, "Abre Explorer++")
	$buttonC.UseVisualStyleBackColor = $True
	$buttonC.add_Click($buttonC_Click)
	#==================================================================
	# SUBBLOQUE: Botón Configuración de Red - button_networkconfig
	#==================================================================
	$button_networkconfig.Font = "Microsoft Sans Serif, 8.25pt"
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_networkconfig.ps1')
	$button_networkconfig.ImageAlign = 'TopCenter'
	$button_networkconfig.Location = '373, 4'
	$button_networkconfig.Name = "button_networkconfig"
	$button_networkconfig.Size = '74, 77'
	$button_networkconfig.TabIndex = 42
	$button_networkconfig.Text = "Ip Config" 
	$button_networkconfig.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($button_networkconfig, "Información de Tarxeta de Rede")
	$button_networkconfig.UseVisualStyleBackColor = $True
	$button_networkconfig.add_Click($button_networkIPConfig_Click)
	#==================================================================
	# SUBBLOQUE: Botón Reiniciar Equipo - button_Restart
	#==================================================================
	$button_Restart.Font = "Microsoft Sans Serif, 8.25pt"
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_Restart.ps1')
	$button_Restart.ImageAlign = 'TopCenter'
	$button_Restart.Location = '669, 4'
	$button_Restart.Name = "button_Restart"
	$button_Restart.Size = '74, 77'
	$button_Restart.TabIndex = 0
	$button_Restart.Text = "Reinicio"
	$button_Restart.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($button_Restart, "Reinicio Equipo")
	$button_Restart.UseVisualStyleBackColor = $True
	$button_Restart.add_Click($button_Restart_Click)
	#==================================================================
	# SUBBLOQUE: Botón Encender Equipo - button_PowerOn
	#==================================================================
	$button_PowerOn.Font = "Microsoft Sans Serif, 8.25pt"
	# Cargar icono desde archivo
	$PowerOnIconPath = Join-Path $Global:ScriptRoot "icos\Encendido.ico"
	if (Test-Path $PowerOnIconPath) {
		$PowerOnImg = [System.Drawing.Image]::FromFile($PowerOnIconPath)
		$PowerOnSize = New-Object System.Drawing.Size(52, 52)
		$button_PowerOn.Image = New-Object System.Drawing.Bitmap($PowerOnImg, $PowerOnSize)
	}
	$button_PowerOn.ImageAlign = 'TopCenter'
	$button_PowerOn.Location = '747, 4'
	$button_PowerOn.Name = "button_PowerOn"
	$button_PowerOn.Size = '74, 77'
	$button_PowerOn.TabIndex = 0
	$button_PowerOn.Text = "Encender"
	$button_PowerOn.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($button_PowerOn, "Encender Equipo mediante Wake-On-LAN")
	$button_PowerOn.UseVisualStyleBackColor = $True
	$button_PowerOn.add_Click($button_PowerOn_Click)
	#==================================================================
	# SUBBLOQUE: Botón Apagar Equipo - button_Shutdown
	#==================================================================
	$button_Shutdown.Font = "Microsoft Sans Serif, 8.25pt"
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_Shutdown.ps1')
	$button_Shutdown.ImageAlign = 'TopCenter'
	$button_Shutdown.Location = '825, 4'
	$button_Shutdown.Name = "button_Shutdown"
	$button_Shutdown.Size = '74, 77'
	$button_Shutdown.TabIndex = 1
	$button_Shutdown.Text = "Apagado"
	$button_Shutdown.TextAlign = 'BottomCenter'
	$tooltipinfo.SetToolTip($button_Shutdown, "Apagado de Equipo")
	$button_Shutdown.UseVisualStyleBackColor = $True
	$button_Shutdown.add_Click($button_Shutdown_Click)
	#==================================================================
	# SUBBLOQUE: Pestaña Sistema Operativo y Equipo - tabpage_ComputerOSSystem
	#==================================================================
	$tabpage_ComputerOSSystem.Controls.Add($groupbox_software)
	$tabpage_ComputerOSSystem.Controls.Add($groupbox_WindowsUpdate)
	$tabpage_ComputerOSSystem.Controls.Add($groupbox_UsersAndGroups)
	$tabpage_ComputerOSSystem.Controls.Add($groupbox_DomainGroups)
	$tabpage_ComputerOSSystem.Controls.Add($groupbox_Hardware)
	$tabpage_ComputerOSSystem.Location = '4, 22'
	$tabpage_ComputerOSSystem.Name = "tabpage_ComputerOSSystem"
	$tabpage_ComputerOSSystem.Size = '1162, 111'
	$tabpage_ComputerOSSystem.TabIndex = 13
	$tabpage_ComputerOSSystem.Text = "Equipo y Sistema Operativo"
	$tabpage_ComputerOSSystem.UseVisualStyleBackColor = $True
	#==================================================================
	# SUBBLOQUE: Grupo Usuarios y Grupos Locales - groupbox_UsersAndGroups
	#==================================================================
	$groupbox_UsersAndGroups.Controls.Add($button_UsersGroupLocalUsers)
	$groupbox_UsersAndGroups.Controls.Add($button_UsersGroupLocalGroups)
	$groupbox_UsersAndGroups.Location = '800, 1'
	$groupbox_UsersAndGroups.Name = "groupbox_UsersAndGroups"
	$groupbox_UsersAndGroups.Size = '123, 81'
	$groupbox_UsersAndGroups.TabIndex = 61
	$groupbox_UsersAndGroups.TabStop = $False
	$groupbox_UsersAndGroups.Text = "Usuarios y Grupos"

	#==================================================================
	# SUBBLOQUE: Grupo Grupos de Dominio - groupbox_DomainGroups
	#==================================================================
	$groupbox_DomainGroups.Controls.Add($button_AD_ShowGroups)
	$groupbox_DomainGroups.Controls.Add($button_AD_AddToGroup)
	$groupbox_DomainGroups.Location = '935, 1'  # justo a la derecha de Usuarios y Grupos
	$groupbox_DomainGroups.Name = "groupbox_DomainGroups"
	$groupbox_DomainGroups.Size = '123, 81'   # ancho igual a Usuarios y Grupos
	$groupbox_DomainGroups.TabIndex = 62
	$groupbox_DomainGroups.TabStop = $False
	$groupbox_DomainGroups.Text = "Grupos De Dominio"
	#==================================================================
	# SUBBLOQUE: Botón Usuarios Locales - button_UsersGroupLocalUsers
	#==================================================================
	$button_UsersGroupLocalUsers.Location = '14, 21'
	$button_UsersGroupLocalUsers.Name = "button_UsersGroupLocalUsers"
	$button_UsersGroupLocalUsers.Size = '94, 23'
	$button_UsersGroupLocalUsers.TabIndex = 17
	$button_UsersGroupLocalUsers.Text = "Usuarios Locales"
	$button_UsersGroupLocalUsers.UseVisualStyleBackColor = $True
	$button_UsersGroupLocalUsers.add_Click($button_UsersGroupLocalUsers_Click)
	#==================================================================
	# SUBBLOQUE: Botón Grupos Locales - button_UsersGroupLocalGroups
	#==================================================================
	$button_UsersGroupLocalGroups.Location = '14, 50'
	$button_UsersGroupLocalGroups.Name = "button_UsersGroupLocalGroups"
	$button_UsersGroupLocalGroups.Size = '94, 23'
	$button_UsersGroupLocalGroups.TabIndex = 18
	$button_UsersGroupLocalGroups.Text = "Grupos Locales"
	$button_UsersGroupLocalGroups.UseVisualStyleBackColor = $True
	$button_UsersGroupLocalGroups.add_Click($button_UsersGroupLocalGroups_Click)
	#==================================================================
	# SUBBLOQUE: Grupo Software y Sistema Operativo - groupbox_software
	#==================================================================
	$groupbox_software.Controls.Add($groupbox_ComputerDescription)
	$groupbox_software.Controls.Add($groupbox_RemoteDesktop)
	$groupbox_software.Controls.Add($groupbox_WinRM_SO)
	$groupbox_software.Controls.Add($buttonApplications)
	$groupbox_software.Controls.Add($button_PageFile)
	$groupbox_software.Controls.Add($button_StartupCommand)
	$groupbox_software.Location = '212, 1'
	$groupbox_software.Name = "groupbox_software"
	$groupbox_software.Size = '450, 102'   # ampliado para incluir grupo WinRM
	$groupbox_software.TabIndex = 60
	$groupbox_software.TabStop = $False
	$groupbox_software.Text = "S.O / Software"
	#==================================================================
	# SUBBLOQUE: Grupo Computer Description
	#==================================================================
	$groupbox_ComputerDescription.Controls.Add($button_ComputerDescriptionChange)
	$groupbox_ComputerDescription.Controls.Add($button_ComputerDescriptionQuery)
	$groupbox_ComputerDescription.Location = '126, 52'
	$groupbox_ComputerDescription.Name = "groupbox_ComputerDescription"
	$groupbox_ComputerDescription.Size = '138, 48'
	$groupbox_ComputerDescription.TabIndex = 57
	$groupbox_ComputerDescription.TabStop = $False
	$groupbox_ComputerDescription.Text = "Descripción de Equipo"
	#==================================================================
	# SUBBLOQUE: Botón Cambiar Descripción de Equipo - button_ComputerDescriptionChange
	#==================================================================
	$button_ComputerDescriptionChange.Location = '71, 19'
	$button_ComputerDescriptionChange.Name = "button_ComputerDescriptionChange"
	$button_ComputerDescriptionChange.Size = '59, 23'
	$button_ComputerDescriptionChange.TabIndex = 57
	$button_ComputerDescriptionChange.Text = "Modificar"
	$button_ComputerDescriptionChange.UseVisualStyleBackColor = $True
	$button_ComputerDescriptionChange.add_Click($button_ComputerDescriptionChange_Click)
	#==================================================================
	# SUBBLOQUE: Botón Consultar Descripción de Equipo - button_ComputerDescriptionQuery
	#==================================================================
	$button_ComputerDescriptionQuery.Location = '6, 19'
	$button_ComputerDescriptionQuery.Name = "button_ComputerDescriptionQuery"
	$button_ComputerDescriptionQuery.Size = '59, 23'
	$button_ComputerDescriptionQuery.TabIndex = 56
	$button_ComputerDescriptionQuery.Text = "Consultar"
	$button_ComputerDescriptionQuery.UseVisualStyleBackColor = $True
	$button_ComputerDescriptionQuery.add_Click($button_ComputerDescriptionQuery_Click)
	#==================================================================
	# SUBBLOQUE: Grupo Windows Update - groupbox_WindowsUpdate
	#==================================================================
	$groupbox_WindowsUpdate.Controls.Add($button_HotFix)
	$groupbox_WindowsUpdate.Location = '666, 1'   # desplazado para dejar espacio al grupo WinRM en groupbox_software
	$groupbox_WindowsUpdate.Name = "groupbox_WindowsUpdate"
	$groupbox_WindowsUpdate.Size = '123, 52'
	$groupbox_WindowsUpdate.TabIndex = 63
	$groupbox_WindowsUpdate.TabStop = $False
	$groupbox_WindowsUpdate.Text = "Windows Update"
	#==================================================================
	# SUBBLOQUE: Botón "Últimas Actualizaciones"
	#=================================================================
	$button_HotFix.Location = '14, 21'
	$button_HotFix.Name = "button_HotFix"
	$button_HotFix.Size = '94, 23'
	$button_HotFix.TabIndex = 50
	$button_HotFix.Text = "Últimos Updates"
	$button_HotFix.UseVisualStyleBackColor = $True
	$button_HotFix.add_Click($button_HotFix_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Activa Escritorio Remoto"
	#=================================================================
	$groupbox_RemoteDesktop.Controls.Add($button_RDPDisable)
	$groupbox_RemoteDesktop.Controls.Add($button_RDPEnable)
	$groupbox_RemoteDesktop.Location = '126, 7'
	$groupbox_RemoteDesktop.Name = "groupbox_RemoteDesktop"
	$groupbox_RemoteDesktop.Size = '138, 42'
	$groupbox_RemoteDesktop.TabIndex = 58
	$groupbox_RemoteDesktop.TabStop = $False
	$groupbox_RemoteDesktop.Text = "Escritorio Remoto" ##PSREMOTING
	#==================================================================
	# SUBBLOQUE: Botón "Desactiva"
	#=================================================================
	$button_RDPDisable.Location = '71, 16'
	$button_RDPDisable.Name = "button_RDPDisable"
	$button_RDPDisable.Size = '59, 23'
	$button_RDPDisable.TabIndex = 49
	$button_RDPDisable.Text = "Desactiva"
	$button_RDPDisable.UseVisualStyleBackColor = $True
	$button_RDPDisable.add_Click($button_RDPDisable_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Activa"
	#=================================================================
	$button_RDPEnable.Location = '6, 16'
	$button_RDPEnable.Name = "button_RDPEnable"
	$button_RDPEnable.Size = '59, 23'
	$button_RDPEnable.TabIndex = 48
	$button_RDPEnable.Text = "Activa"
	$button_RDPEnable.UseVisualStyleBackColor = $True
	$button_RDPEnable.add_Click($button_RDPEnable_Click)
	#==================================================================
	# SUBBLOQUE: Grupo WinRM (Sistema Operativo) - groupbox_WinRM_SO
	#==================================================================
	$groupbox_WinRM_SO.Controls.Add($button_ActivarWinRM_SO)
	$groupbox_WinRM_SO.Controls.Add($button_DeshabilitarWinRM_SO)
	$groupbox_WinRM_SO.Location = '270, 7'
	$groupbox_WinRM_SO.Name = "groupbox_WinRM_SO"
	$groupbox_WinRM_SO.Size = '174, 80'
	$groupbox_WinRM_SO.TabStop = $False
	$groupbox_WinRM_SO.Text = "WinRM"
	#==================================================================
	# SUBBLOQUE: Botón Activar WinRM (SO)
	#==================================================================
	$button_ActivarWinRM_SO.Location = '6, 17'
	$button_ActivarWinRM_SO.Name = "button_ActivarWinRM_SO"
	$button_ActivarWinRM_SO.Size = '160, 23'
	$button_ActivarWinRM_SO.Text = "Activar WinRM"
	$button_ActivarWinRM_SO.UseVisualStyleBackColor = $True
	$button_ActivarWinRM_SO.add_Click($button_ActivarWinRM_Click)
	#==================================================================
	# SUBBLOQUE: Botón Deshabilitar WinRM (SO)
	#==================================================================
	$button_DeshabilitarWinRM_SO.Location = '6, 47'
	$button_DeshabilitarWinRM_SO.Name = "button_DeshabilitarWinRM_SO"
	$button_DeshabilitarWinRM_SO.Size = '160, 23'
	$button_DeshabilitarWinRM_SO.Text = "Deshabilitar WinRM"
	$button_DeshabilitarWinRM_SO.UseVisualStyleBackColor = $True
	$button_DeshabilitarWinRM_SO.add_Click($button_DeshabilitarWinRM_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Aplicaciones"
	#=================================================================
	$buttonApplications.Location = '6, 19'
	$buttonApplications.Name = "buttonApplications"
	$buttonApplications.Size = '114, 23'
	$buttonApplications.TabIndex = 56
	$buttonApplications.Text = "Aplicaciones"
	$buttonApplications.UseVisualStyleBackColor = $True
	$buttonApplications.add_Click($buttonApplications_Click)
	#==================================================================
	# SUBBLOQUE: Botón "PageFile"
	#=================================================================
	$button_PageFile.Location = '6, 48'
	$button_PageFile.Name = "button_PageFile"
	$button_PageFile.Size = '114, 23'
	$button_PageFile.TabIndex = 52
	$button_PageFile.Text = "PageFile"
	$button_PageFile.UseVisualStyleBackColor = $True
	$button_PageFile.add_Click($button_PageFile_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Comandos de Inicio"
	#=================================================================
	$button_StartupCommand.Location = '6, 76'
	$button_StartupCommand.Name = "button_StartupCommand"
	$button_StartupCommand.Size = '114, 23'
	$button_StartupCommand.TabIndex = 27
	$button_StartupCommand.Text = "StartUp"
	$button_StartupCommand.UseVisualStyleBackColor = $True
	$button_StartupCommand.add_Click($button_StartupCommand_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Hardware"
	#=================================================================
	$groupbox_Hardware.Controls.Add($button_MotherBoard)
	$groupbox_Hardware.Controls.Add($button_Processor)
	$groupbox_Hardware.Controls.Add($button_Memory)
	$groupbox_Hardware.Controls.Add($button_SystemType)
	$groupbox_Hardware.Controls.Add($button_Printers)
	$groupbox_Hardware.Controls.Add($button_USBDevices)
	$groupbox_Hardware.Location = '2, 1'
	$groupbox_Hardware.Name = "groupbox_Hardware"
	$groupbox_Hardware.Size = '204, 102'
	$groupbox_Hardware.TabIndex = 59
	$groupbox_Hardware.TabStop = $False
	$groupbox_Hardware.Text = "Hardware"
	#==================================================================
	# SUBBLOQUE: Botón "Placa Base"
	#=================================================================
	$button_MotherBoard.Location = '6, 19'
	$button_MotherBoard.Name = "button_MotherBoard"
	$button_MotherBoard.Size = '93, 23'
	$button_MotherBoard.TabIndex = 54
	$button_MotherBoard.Text = "Placa Base"
	$button_MotherBoard.UseVisualStyleBackColor = $True
	$button_MotherBoard.add_Click($button_MotherBoard_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Procesador"
	#=================================================================
	$button_Processor.Location = '6, 48'
	$button_Processor.Name = "button_Processor"
	$button_Processor.Size = '93, 23'
	$button_Processor.TabIndex = 53
	$button_Processor.Text = "Procesador"
	$button_Processor.UseVisualStyleBackColor = $True
	$button_Processor.add_Click($button_Processor_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Memoria"
	#=================================================================
	$button_Memory.Font = "Trebuchet MS, 8.25pt"
	$button_Memory.Location = '6, 77'
	$button_Memory.Name = "button_Memory"
	$button_Memory.Size = '93, 23'
	$button_Memory.TabIndex = 22
	$button_Memory.Text = "Memoria"
	$button_Memory.UseVisualStyleBackColor = $True
	$button_Memory.add_Click($button_Memory_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Sistema"
	#=================================================================
	$button_SystemType.Location = '105, 19'
	$button_SystemType.Name = "button_SystemType"
	$button_SystemType.Size = '93, 23'
	$button_SystemType.TabIndex = 55
	$button_SystemType.Text = "Sistema"
	$button_SystemType.UseVisualStyleBackColor = $True
	$button_SystemType.add_Click($button_SystemType_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Impresoras"
	#=================================================================
	$button_Printers.Location = '105, 77'
	$button_Printers.Name = "button_Printers"
	$button_Printers.Size = '93, 23'
	$button_Printers.TabIndex = 51
	$button_Printers.Text = "Impresoras"
	$button_Printers.UseVisualStyleBackColor = $True
	$button_Printers.add_Click($button_Printers_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Dispositivos USB"
	#=================================================================
	$button_USBDevices.Location = '105, 48'
	$button_USBDevices.Name = "button_USBDevices"
	$button_USBDevices.Size = '93, 23'
	$button_USBDevices.TabIndex = 47
	$button_USBDevices.Text = "Dispositivos USB"
	$button_USBDevices.UseVisualStyleBackColor = $True
	$button_USBDevices.add_Click($button_USBDevices_Click)
	##==================================================================
	# SUBBLOQUE: Botón "Red"
	#=================================================================
	$tabpage_network.Controls.Add($button_ConnectivityTesting)
	$tabpage_network.Controls.Add($button_NIC)
	$tabpage_network.Controls.Add($button_networkIPConfig)
	$tabpage_network.Controls.Add($button_networkTestPort)
	$tabpage_network.Controls.Add($button_networkRouteTable)
	$tabpage_network.Location = '4, 22'
	$tabpage_network.Name = "tabpage_network"
	$tabpage_network.Size = '1162, 111'
	$tabpage_network.TabIndex = 6
	$tabpage_network.Text = "Red"
	$tabpage_network.UseVisualStyleBackColor = $True
	#==================================================================
	# SUBBLOQUE: Botón "Test de Conectividad (slow)"
	#=================================================================
	$button_ConnectivityTesting.FlatStyle = 'System'
	$button_ConnectivityTesting.Location = '137, 3'
	$button_ConnectivityTesting.Name = "button_ConnectivityTesting"
	$button_ConnectivityTesting.Size = '145, 23'
	$button_ConnectivityTesting.TabIndex = 46
	$button_ConnectivityTesting.Text = "Test de Conectividad (slow)"
	$tooltipinfo.SetToolTip($button_ConnectivityTesting, "Comprobando Conectividad")
	$button_ConnectivityTesting.UseVisualStyleBackColor = $False
	$button_ConnectivityTesting.add_Click($button_ConnectivityTesting_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Interfaz de Red (modo slow)"
	#=================================================================
	$button_NIC.Location = '137, 32'
	$button_NIC.Name = "button_NIC"
	$button_NIC.Size = '145, 23'
	$button_NIC.TabIndex = 11
	$button_NIC.Text = "Interfaz de Red (modo slow)"
	$tooltipinfo.SetToolTip($button_NIC, "Información de tarjeta(s) de red")
	$button_NIC.UseVisualStyleBackColor = $True
	$button_NIC.add_Click($button_NIC_Click)
	#==================================================================
	# SUBBLOQUE: Botón "IPConfig"
	#=================================================================
	$button_networkIPConfig.Location = '8, 3'
	$button_networkIPConfig.Name = "button_networkIPConfig"
	$button_networkIPConfig.Size = '123, 23'
	$button_networkIPConfig.TabIndex = 9
	$button_networkIPConfig.Text = "IPConfig"
	$tooltipinfo.SetToolTip($button_networkIPConfig, "Configuración IP") ##ACTIVA DHCP
	$button_networkIPConfig.UseVisualStyleBackColor = $True
	$button_networkIPConfig.add_Click($button_networkIPConfig_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Comprueba un puerto"
	#=================================================================
	$button_networkTestPort.Location = '8, 61'
	$button_networkTestPort.Name = "button_networkTestPort"
	$button_networkTestPort.Size = '123, 23'
	$button_networkTestPort.TabIndex = 8
	$button_networkTestPort.Text = "Compruobar Puerto"
	$tooltipinfo.SetToolTip($button_networkTestPort, "Comprueba un puerto (por defecto = 80)")
	$button_networkTestPort.UseVisualStyleBackColor = $True
	$button_networkTestPort.add_Click($button_networkTestPort_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Route Table"
	#=================================================================
	$button_networkRouteTable.Location = '8, 32'
	$button_networkRouteTable.Name = "button_networkRouteTable"
	$button_networkRouteTable.Size = '123, 23'
	$button_networkRouteTable.TabIndex = 10
	$button_networkRouteTable.Text = "Route Table"
	$tooltipinfo.SetToolTip($button_networkRouteTable, "Route Table")
	$button_networkRouteTable.UseVisualStyleBackColor = $True
	$button_networkRouteTable.add_Click($button_networkRouteTable_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Procesos"
	#=================================================================
	$tabpage_processes.Controls.Add($buttonCommandLineGridView)
	$tabpage_processes.Controls.Add($button_processAll)
	$tabpage_processes.Controls.Add($buttonCommandLine)
	$tabpage_processes.Controls.Add($groupbox1)
	$tabpage_processes.Controls.Add($button_process100MB)
	$tabpage_processes.Controls.Add($button_ProcessGrid)
	$tabpage_processes.Controls.Add($button_processOwners)
	$tabpage_processes.Controls.Add($button_processLastHour)
	$tabpage_processes.Location = '4, 22'
	$tabpage_processes.Name = "tabpage_processes"
	$tabpage_processes.Size = '1162, 111'
	$tabpage_processes.TabIndex = 3
	$tabpage_processes.Text = "Procesos"
	$tabpage_processes.UseVisualStyleBackColor = $True
	#==================================================================
	# SUBBLOQUE: Botón "CommandLine - Procesos"
	#=================================================================
	$buttonCommandLineGridView.Location = '284, 3'
	$buttonCommandLineGridView.Name = "buttonCommandLineGridView"
	$buttonCommandLineGridView.Size = '159, 23'
	$buttonCommandLineGridView.TabIndex = 17
	$buttonCommandLineGridView.Text = "CommandLine - Procesos GUI"
	$buttonCommandLineGridView.UseVisualStyleBackColor = $True
	$buttonCommandLineGridView.add_Click($buttonCommandLineGridView_Click)
	#==================================================================
	# SUBBLOQUE: Botón "Obtiene todos los procesos"
	#=================================================================
	$button_processAll.Location = '8, 3'
	$button_processAll.Name = "button_processAll"
	$button_processAll.Size = '132, 23'
	$button_processAll.TabIndex = 5
	$button_processAll.Text = "Procesos"
	$tooltipinfo.SetToolTip($button_processAll, "Obtiene todos los procesos")
	$button_processAll.UseVisualStyleBackColor = $True
	$button_processAll.add_Click($button_processAll_Click)
	#==================================================================
	# SUBBLOQUE: Botón Obtiene los procesos CommandLine
	#=================================================================
	$buttonCommandLine.Location = '146, 61'
	$buttonCommandLine.Name = "buttonCommandLine"
	$buttonCommandLine.Size = '132, 23'
	$buttonCommandLine.TabIndex = 15
	$buttonCommandLine.Text = "Procesos CommandLine"
	$tooltipinfo.SetToolTip($buttonCommandLine, "Obtiene los procesos CommandLine")
	$buttonCommandLine.UseVisualStyleBackColor = $True
	$buttonCommandLine.add_Click($buttonCommandLine_Click)
	#==================================================================
	# SUBBLOQUE: Etiqueta Cierra un Proceso 
	#=================================================================
	$groupbox1.Controls.Add($textbox_processName)
	$groupbox1.Controls.Add($label_processEnterAProcessName)
	$groupbox1.Controls.Add($button_processTerminate)
	$groupbox1.Location = '870, 17'
	$groupbox1.Name = "groupbox1"
	$groupbox1.Size = '231, 83'
	$groupbox1.TabIndex = 16
	$groupbox1.TabStop = $False
	$groupbox1.Text = "Cierra un Proceso"
	#==================================================================
	# SUBBLOQUE: Etiqueta "Introduce un nombre de proceso, ""msedge.exe" 
	#=================================================================
	[void]$textbox_processName.AutoCompleteCustomSource.Add("dhcp")
	[void]$textbox_processName.AutoCompleteCustomSource.Add("iisadmin")
	[void]$textbox_processName.AutoCompleteCustomSource.Add("msftpsvc")
	[void]$textbox_processName.AutoCompleteCustomSource.Add("nntpsvc")
	[void]$textbox_processName.AutoCompleteCustomSource.Add("omniinet")
	[void]$textbox_processName.AutoCompleteCustomSource.Add("smtpsvc")
	[void]$textbox_processName.AutoCompleteCustomSource.Add("spooler")
	[void]$textbox_processName.AutoCompleteCustomSource.Add("sql")
	[void]$textbox_processName.AutoCompleteCustomSource.Add("w3svc")
	$textbox_processName.AutoCompleteMode = 'Suggest'
	$textbox_processName.Location = '6, 19'
	$textbox_processName.Name = "textbox_processName"
	$textbox_processName.Size = '116, 20'
	$textbox_processName.TabIndex = 12
	$textbox_processName.Text = "<NombreProceso>"
	$tooltipinfo.SetToolTip($textbox_processName, "Introduce un nombre de proceso, ""msedge.exe""")
	#==================================================================
	# SUBBLOQUE: Etiqueta Introduce el nombre de un proceso (ex:msedge.exe) 
	#=================================================================
	$label_processEnterAProcessName.Font = "Trebuchet MS, 6.75pt, style=Italic"
	$label_processEnterAProcessName.Location = '6, 41'
	$label_processEnterAProcessName.Name = "label_processEnterAProcessName"
	$label_processEnterAProcessName.Size = '206, 23'
	$label_processEnterAProcessName.TabIndex = 14
	$label_processEnterAProcessName.Text = "Introduce el nombre de un proceso (ex:msedge.exe)"
	$label_processEnterAProcessName.TextAlign = 'MiddleCenter'
	#==================================================================
	# SUBBLOQUE: Botón  Cerrar el proceso
	#=================================================================
	$button_processTerminate.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$button_processTerminate.ForeColor = 'Red'
	$button_processTerminate.Location = '128, 18'
	$button_processTerminate.Name = "button_processTerminate"
	$button_processTerminate.Size = '84, 23'
	$button_processTerminate.TabIndex = 13
	$button_processTerminate.Text = "Cerar"
	$tooltipinfo.SetToolTip($button_processTerminate, "Cerrar el proceso")
	$button_processTerminate.UseVisualStyleBackColor = $True
	$button_processTerminate.add_Click($button_processTerminate_Click)
	#==================================================================
	# SUBBLOQUE: Etiqueta Obtiene todos los procesos que están consumiento más de 100MB de Memoria 
	#=================================================================
	$button_process100MB.Location = '8, 32'
	$button_process100MB.Name = "button_process100MB"
	$button_process100MB.Size = '132, 23'
	$button_process100MB.TabIndex = 0
	$button_process100MB.Text = "+100MB Memoria"
	$tooltipinfo.SetToolTip($button_process100MB, "Obtiene todos los procesos que están consumiento más de 100MB de Memoria")
	$button_process100MB.UseVisualStyleBackColor = $True
	$button_process100MB.add_Click($button_process100MB_Click)
	#==================================================================
	# SUBBLOQUE: Etiqueta Obtiene todos los procesos en un grid 
	#=================================================================
	$button_ProcessGrid.Location = '146, 32'
	$button_ProcessGrid.Name = "button_ProcessGrid"
	$button_ProcessGrid.Size = '132, 23'
	$button_ProcessGrid.TabIndex = 6
	$button_ProcessGrid.Text = "Procesos - Grid Vista"
	$tooltipinfo.SetToolTip($button_ProcessGrid, "Obtiene todos los procesos en un grid")
	$button_ProcessGrid.UseVisualStyleBackColor = $True
	$button_ProcessGrid.add_Click($button_ProcessGrid_Click)
	#==================================================================
	# SUBBLOQUE: Etiqueta Obtiene el propietario de cada proceso 
	#=================================================================
	$button_processOwners.Location = '8, 61'
	$button_processOwners.Name = "button_processOwners"
	$button_processOwners.Size = '132, 23'
	$button_processOwners.TabIndex = 4
	$button_processOwners.Text = "Propietarios"
	$tooltipinfo.SetToolTip($button_processOwners, "Obtiene el propietario de cada proceso")
	$button_processOwners.UseVisualStyleBackColor = $True
	$button_processOwners.add_Click($button_processOwners_Click)
	#==================================================================
	# SUBBLOQUE: Etiqueta Obtiene los procesos que se ha iniciado en la última hora 
	#=================================================================
	$button_processLastHour.Location = '146, 3'
	$button_processLastHour.Name = "button_processLastHour"
	$button_processLastHour.Size = '132, 23'
	$button_processLastHour.TabIndex = 7
	$button_processLastHour.Text = "Empezados < 1h"
	$tooltipinfo.SetToolTip($button_processLastHour, "Obtiene los procesos que se ha iniciado en la última hora")
	$button_processLastHour.UseVisualStyleBackColor = $True
	$button_processLastHour.add_Click($button_processLastHour_Click)
	#==================================================================
	# SUBBLOQUE: Botón Servicios
	#=================================================================
	$tabpage_services.Controls.Add($button_servicesNonStandardUser)
	$tabpage_services.Controls.Add($button_mmcServices)
	$tabpage_services.Controls.Add($button_servicesAutoNotStarted)
	$tabpage_services.Controls.Add($groupbox_Service_QueryStartStop)
	$tabpage_services.Controls.Add($button_servicesRunning)
	$tabpage_services.Controls.Add($button_servicesAll)
	$tabpage_services.Controls.Add($button_servicesGridView)
	$tabpage_services.Controls.Add($button_servicesAutomatic)
	$tabpage_services.Location = '4, 22'
	$tabpage_services.Name = "tabpage_services"
	$tabpage_services.Size = '1162, 111'
	$tabpage_services.TabIndex = 2
	$tabpage_services.Text = "Servicios"
	$tooltipinfo.SetToolTip($tabpage_services, "Servicios")
	$tabpage_services.UseVisualStyleBackColor = $True
	#==================================================================
	# SUBBLOQUE: Etiqueta Obtiene servicios configurados con un usuario no estandar
	#=================================================================
	$button_servicesNonStandardUser.Location = '285, 3'
	$button_servicesNonStandardUser.Name = "button_servicesNonStandardUser"
	$button_servicesNonStandardUser.Size = '133, 23'
	$button_servicesNonStandardUser.TabIndex = 8
	$button_servicesNonStandardUser.Text = "Usuario No Estandar"
	$tooltipinfo.SetToolTip($button_servicesNonStandardUser, "Obtiene servicios configurados con un usuario no estandar")
	$button_servicesNonStandardUser.UseVisualStyleBackColor = $True
	$button_servicesNonStandardUser.add_Click($button_servicesNonStandardUser_Click)
	#==================================================================
	# SUBBLOQUE: Botón  Abre Consola Servicios
	#=================================================================
	$button_mmcServices.ForeColor = 'ForestGreen'
	$button_mmcServices.Location = '8, 3'
	$button_mmcServices.Name = "button_mmcServices"
	$button_mmcServices.Size = '132, 23'
	$button_mmcServices.TabIndex = 0
	$button_mmcServices.Text = "MMC: Consola Servicios"
	$tooltipinfo.SetToolTip($button_mmcServices, "Abre Consola Servicios")
	$button_mmcServices.UseVisualStyleBackColor = $True
	$button_mmcServices.add_Click($button_mmcServices_Click)
	#==================================================================
	# SUBBLOQUE: Etiqueta  Servicios en ""Automatico"" y Estado distinto a ""Arrancados
	#=================================================================
	$button_servicesAutoNotStarted.Location = '146, 61'
	$button_servicesAutoNotStarted.Name = "button_servicesAutoNotStarted"
	$button_servicesAutoNotStarted.Size = '133, 23'
	$button_servicesAutoNotStarted.TabIndex = 9
	$button_servicesAutoNotStarted.Text = "Auto + No Ejecucción"
	$tooltipinfo.SetToolTip($button_servicesAutoNotStarted, "Servicios en ""Automatico"" y Estado distinto a ""Arrancados""")
	$button_servicesAutoNotStarted.UseVisualStyleBackColor = $True
	$button_servicesAutoNotStarted.add_Click($button_servicesAutoNotStarted_Click)
	#==================================================================
	# SUBBLOQUE: Botón Consulta/Arranca/Para
	#=================================================================
	$groupbox_Service_QueryStartStop.Controls.Add($textbox_servicesAction)
	$groupbox_Service_QueryStartStop.Controls.Add($button_servicesRestart)
	$groupbox_Service_QueryStartStop.Controls.Add($label_servicesEnterAServiceName)
	$groupbox_Service_QueryStartStop.Controls.Add($button_servicesQuery)
	$groupbox_Service_QueryStartStop.Controls.Add($button_servicesStart)
	$groupbox_Service_QueryStartStop.Controls.Add($button_servicesStop)
	$groupbox_Service_QueryStartStop.Location = '872, 10'
	$groupbox_Service_QueryStartStop.Name = "groupbox_Service_QueryStartStop"
	$groupbox_Service_QueryStartStop.Size = '274, 90'
	$groupbox_Service_QueryStartStop.TabIndex = 13
	$groupbox_Service_QueryStartStop.TabStop = $False
	$groupbox_Service_QueryStartStop.Text = "Consulta/Arranca/Para"
	#==================================================================
	# SUBBLOQUE: Eiqueta Introduce el nombre del servicio 
	#=================================================================
	[void]$textbox_servicesAction.AutoCompleteCustomSource.Add("dhcp")
	[void]$textbox_servicesAction.AutoCompleteCustomSource.Add("iisadmin")
	[void]$textbox_servicesAction.AutoCompleteCustomSource.Add("msftpsvc")
	[void]$textbox_servicesAction.AutoCompleteCustomSource.Add("nntpsvc")
	[void]$textbox_servicesAction.AutoCompleteCustomSource.Add("omniinet")
	[void]$textbox_servicesAction.AutoCompleteCustomSource.Add("smtpsvc")
	[void]$textbox_servicesAction.AutoCompleteCustomSource.Add("spooler")
	[void]$textbox_servicesAction.AutoCompleteCustomSource.Add("sql")
	[void]$textbox_servicesAction.AutoCompleteCustomSource.Add("w3svc")
	$textbox_servicesAction.AutoCompleteMode = 'Suggest'
	$textbox_servicesAction.Location = '6, 60'
	$textbox_servicesAction.Name = "textbox_servicesAction"
	$textbox_servicesAction.Size = '116, 120'
	$textbox_servicesAction.TabIndex = 100
	$textbox_servicesAction.Text = "<NombredelServicio>"
	$tooltipinfo.SetToolTip($textbox_servicesAction, "Introduce el nombre del servicio")
	$textbox_servicesAction.add_Click($textbox_servicesAction_Click)
	#==================================================================
	# SUBBLOQUE: Botón Reinicia el servicio
	#=================================================================
	$button_servicesRestart.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$button_servicesRestart.Location = '200, 15'
	$button_servicesRestart.Name = "button_servicesRestart"
	$button_servicesRestart.Size = '68, 23'
	$button_servicesRestart.TabIndex = 10
	$button_servicesRestart.Text = "Reinicia"
	$tooltipinfo.SetToolTip($button_servicesRestart, "Reinicia el servicio")
	$button_servicesRestart.UseVisualStyleBackColor = $True
	$button_servicesRestart.add_Click($button_servicesRestart_Click)
	#==================================================================
	# label_servicesEnterAServiceName
	#==================================================================
	$label_servicesEnterAServiceName.Font = "Trebuchet MS, 6.75pt, style=Italic"
	$label_servicesEnterAServiceName.Location = '6, 37'
	$label_servicesEnterAServiceName.Name = "label_servicesEnterAServiceName"
	$label_servicesEnterAServiceName.Size = '116, 15'
	$label_servicesEnterAServiceName.TabIndex = 12
	$label_servicesEnterAServiceName.Text = "Introduce el nombre del Servicio"
	$label_servicesEnterAServiceName.TextAlign = 'MiddleCenter'
	#==================================================================
	# button_servicesQuery
	#==================================================================
	$button_servicesQuery.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$button_servicesQuery.Location = '128, 19'
	$button_servicesQuery.Name = "button_servicesQuery"
	$button_servicesQuery.Size = '59, 23'
	$button_servicesQuery.TabIndex = 4
	$button_servicesQuery.Text = "Consulta"
	$tooltipinfo.SetToolTip($button_servicesQuery, "Obtiene informacion del servicio indicado")
	$button_servicesQuery.UseVisualStyleBackColor = $True
	$button_servicesQuery.add_Click($button_servicesQuery_Click)
	#==================================================================
	# button_servicesStart
	#==================================================================
	$button_servicesStart.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$button_servicesStart.ForeColor = 'DarkBlue'
	$button_servicesStart.Location = '200, 37'
	$button_servicesStart.Name = "button_servicesStart"
	$button_servicesStart.Size = '68, 23'
	$button_servicesStart.TabIndex = 6
	$button_servicesStart.Text = "Arranca"
	$tooltipinfo.SetToolTip($button_servicesStart, "Arranca el servicio indicado")
	$button_servicesStart.UseVisualStyleBackColor = $True
	$button_servicesStart.add_Click($button_servicesStart_Click)
	#==================================================================
	# button_servicesStop
	#==================================================================
	$button_servicesStop.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	$button_servicesStop.ForeColor = 'Red'
	$button_servicesStop.Location = '200, 59'
	$button_servicesStop.Name = "button_servicesStop"
	$button_servicesStop.Size = '68, 23'
	$button_servicesStop.TabIndex = 5
	$button_servicesStop.Text = "Para"
	$tooltipinfo.SetToolTip($button_servicesStop, "Detiene el servicio indicado")
	$button_servicesStop.UseVisualStyleBackColor = $True
	$button_servicesStop.add_Click($button_servicesStop_Click)
	#==================================================================
	# button_servicesRunning
	#==================================================================
	$button_servicesRunning.Location = '146, 3'
	$button_servicesRunning.Name = "button_servicesRunning"
	$button_servicesRunning.Size = '133, 23'
	$button_servicesRunning.TabIndex = 1
	$button_servicesRunning.Text = "En Ejecución"
	$tooltipinfo.SetToolTip($button_servicesRunning, "Servicios en ejecución")
	$button_servicesRunning.UseVisualStyleBackColor = $True
	$button_servicesRunning.add_Click($button_servicesRunning_Click)
	#==================================================================
	# button_servicesAll
	#==================================================================
	$button_servicesAll.Location = '8, 32'
	$button_servicesAll.Name = "button_servicesAll"
	$button_servicesAll.Size = '132, 23'
	$button_servicesAll.TabIndex = 3
	$button_servicesAll.Text = "Servicios"
	$tooltipinfo.SetToolTip($button_servicesAll, "Obtiene todos los servicios")
	$button_servicesAll.UseVisualStyleBackColor = $True
	$button_servicesAll.add_Click($button_servicesAll_Click)
	#==================================================================
	# button_servicesGridView
	#==================================================================
	$button_servicesGridView.Location = '8, 61'
	$button_servicesGridView.Name = "button_servicesGridView"
	$button_servicesGridView.Size = '132, 23'
	$button_servicesGridView.TabIndex = 7
	$button_servicesGridView.Text = "Services - GridView"
	$tooltipinfo.SetToolTip($button_servicesGridView, "Get all the services in a Grid View form")
	$button_servicesGridView.UseVisualStyleBackColor = $True
	$button_servicesGridView.add_Click($button_servicesGridView_Click)
	#==================================================================
	# button_servicesAutomatic
	#==================================================================
	$button_servicesAutomatic.Location = '146, 32'
	$button_servicesAutomatic.Name = "button_servicesAutomatic"
	$button_servicesAutomatic.Size = '133, 23'
	$button_servicesAutomatic.TabIndex = 2
	$button_servicesAutomatic.Text = "Automatic"
	$tooltipinfo.SetToolTip($button_servicesAutomatic, "Services with StartupType = Automatic")
	$button_servicesAutomatic.UseVisualStyleBackColor = $True
	$button_servicesAutomatic.add_Click($button_servicesAutomatic_Click)
	#==================================================================
	# tabpage_diskdrives
	#==================================================================
	$tabpage_diskdrives.Controls.Add($button_DiskUsage)
	$tabpage_diskdrives.Controls.Add($button_DiskPartition)
	$tabpage_diskdrives.Controls.Add($button_DiskLogical)
	$tabpage_diskdrives.Controls.Add($button_DiskMountPoint)
	$tabpage_diskdrives.Controls.Add($button_DiskRelationship)
	$tabpage_diskdrives.Controls.Add($button_DiskMappedDrive)
	$tabpage_diskdrives.Location = '4, 22'
	$tabpage_diskdrives.Name = "tabpage_diskdrives"
	$tabpage_diskdrives.Size = '1162, 111'
	$tabpage_diskdrives.TabIndex = 10
	$tabpage_diskdrives.Text = "Disk Drives"
	$tabpage_diskdrives.UseVisualStyleBackColor = $True
	#==================================================================
	# button_DiskUsage
	#==================================================================
	$button_DiskUsage.Location = '8, 3'
	$button_DiskUsage.Name = "button_DiskUsage"
	$button_DiskUsage.Size = '112, 23'
	$button_DiskUsage.TabIndex = 22
	$button_DiskUsage.Text = "Capacidad Discos"
	$tooltipinfo.SetToolTip($button_DiskUsage, "Tamaño dispoñible")
	$button_DiskUsage.UseVisualStyleBackColor = $True
	$button_DiskUsage.add_Click($button_DiskUsage_Click)
	#==================================================================
	# button_DiskPartition
	#==================================================================
	$button_DiskPartition.Location = '126, 61'
	$button_DiskPartition.Name = "button_DiskPartition"
	$button_DiskPartition.Size = '112, 23'
	$button_DiskPartition.TabIndex = 17
	$button_DiskPartition.Text = "Particións"
	$tooltipinfo.SetToolTip($button_DiskPartition, "Particións do equipo")
	$button_DiskPartition.UseVisualStyleBackColor = $True
	$button_DiskPartition.add_Click($button_DiskPartition_Click)
	#==================================================================
	# button_DiskLogical
	#==================================================================
	$button_DiskLogical.Location = '126, 3'
	$button_DiskLogical.Name = "button_DiskLogical"
	$button_DiskLogical.Size = '112, 23'
	$button_DiskLogical.TabIndex = 13
	$button_DiskLogical.Text = "Discos Loxicos"
	$tooltipinfo.SetToolTip($button_DiskLogical, "Visualizar discos loxicos")
	$button_DiskLogical.UseVisualStyleBackColor = $True
	$button_DiskLogical.add_Click($button_DiskLogical_Click)
	#==================================================================
	# button_DiskMountPoint
	#==================================================================
	$button_DiskMountPoint.Location = '126, 32'
	$button_DiskMountPoint.Name = "button_DiskMountPoint"
	$button_DiskMountPoint.Size = '112, 23'
	$button_DiskMountPoint.TabIndex = 20
	$button_DiskMountPoint.Text = "Discos Fisicos"
	$tooltipinfo.SetToolTip($button_DiskMountPoint, "Punto montaxe discos fisicos")
	$button_DiskMountPoint.UseVisualStyleBackColor = $True
	$button_DiskMountPoint.add_Click($button_DiskMountPoint_Click)
	#==================================================================
	# button_DiskRelationship
	#==================================================================
	$button_DiskRelationship.Location = '8, 61'
	$button_DiskRelationship.Name = "button_DiskRelationship"
	$button_DiskRelationship.Size = '112, 23'
	$button_DiskRelationship.TabIndex = 19
	$button_DiskRelationship.Text = "Disk Relationship"
	$tooltipinfo.SetToolTip($button_DiskRelationship, "Get the disk(s) relationship")
	$button_DiskRelationship.UseVisualStyleBackColor = $True
	$button_DiskRelationship.add_Click($button_DiskRelationship_Click)
	#==================================================================
	# button_DiskMappedDrive
	#==================================================================
	$button_DiskMappedDrive.Location = '8, 32'
	$button_DiskMappedDrive.Name = "button_DiskMappedDrive"
	$button_DiskMappedDrive.Size = '112, 23'
	$button_DiskMappedDrive.TabIndex = 21
	$button_DiskMappedDrive.Text = "Unidades Mapeadas"
	$tooltipinfo.SetToolTip($button_DiskMappedDrive, "Obten as unidades mapeadas")
	$button_DiskMappedDrive.UseVisualStyleBackColor = $True
	$button_DiskMappedDrive.add_Click($button_DiskMappedDrive_Click)
	#==================================================================
	# tabpage_shares
	#==================================================================
	$tabpage_shares.Controls.Add($button_mmcShares)
	$tabpage_shares.Controls.Add($button_SharesGrid)
	$tabpage_shares.Controls.Add($button_Shares)
	$tabpage_shares.Location = '4, 22'
	$tabpage_shares.Name = "tabpage_shares"
	$tabpage_shares.Size = '1162, 111'
	$tabpage_shares.TabIndex = 7
	$tabpage_shares.Text = "Shares"
	$tabpage_shares.UseVisualStyleBackColor = $True
	#==================================================================
	# button_mmcShares
	#==================================================================
	$button_mmcShares.ForeColor = 'ForestGreen'
	$button_mmcShares.Location = '8, 3'
	$button_mmcShares.Name = "button_mmcShares"
	$button_mmcShares.Size = '140, 23'
	$button_mmcShares.TabIndex = 1
	$button_mmcShares.Text = "MMC: Shares"
	$tooltipinfo.SetToolTip($button_mmcShares, "Launch the shared folders console (fsmgmt.msc)")
	$button_mmcShares.UseVisualStyleBackColor = $True
	$button_mmcShares.add_Click($button_mmcShares_Click)
	#==================================================================
	# button_SharesGrid
	#==================================================================
	$button_SharesGrid.Location = '8, 61'
	$button_SharesGrid.Name = "button_SharesGrid"
	$button_SharesGrid.Size = '140, 23'
	$button_SharesGrid.TabIndex = 16
	$button_SharesGrid.Text = "Shares - GridView"
	$tooltipinfo.SetToolTip($button_SharesGrid, "Get a list of all the shares in a Grid View form")
	$button_SharesGrid.UseVisualStyleBackColor = $True
	$button_SharesGrid.add_Click($button_SharesGrid_Click)
	#==================================================================
	# button_Shares
	#==================================================================
	$button_Shares.Location = '8, 32'
	$button_Shares.Name = "button_Shares"
	$button_Shares.Size = '140, 23'
	$button_Shares.TabIndex = 0
	$button_Shares.Text = "Shares"
	$tooltipinfo.SetToolTip($button_Shares, "Get a list of all the shares with a local path")
	$button_Shares.UseVisualStyleBackColor = $True
	$button_Shares.add_Click($button_Shares_Click)
	#==================================================================
	# tabpage_eventlog
	#==================================================================
	$tabpage_eventlog.Controls.Add($button_RebootHistory)
	$tabpage_eventlog.Controls.Add($button_mmcEvents)
	$tabpage_eventlog.Controls.Add($button_EventsLogNames)
	$tabpage_eventlog.Location = '4, 22'
	$tabpage_eventlog.Name = "tabpage_eventlog"
	$tabpage_eventlog.Size = '1162, 111'
	$tabpage_eventlog.TabIndex = 4
	$tabpage_eventlog.Text = "Event Log"
	$tabpage_eventlog.UseVisualStyleBackColor = $True
	#==================================================================
	# button_RebootHistory
	#==================================================================
	$button_RebootHistory.Location = '163, 3'
	$button_RebootHistory.Name = "button_RebootHistory"
	$button_RebootHistory.Size = '149, 23'
	$button_RebootHistory.TabIndex = 5
	$button_RebootHistory.Text = "Historial Reinicios (slow)"
	$button_RebootHistory.UseVisualStyleBackColor = $True
	$button_RebootHistory.add_Click($button_RebootHistory_Click)
	#==================================================================
	# button_mmcEvents
	#==================================================================
	$button_mmcEvents.ForeColor = 'ForestGreen'
	$button_mmcEvents.Location = '8, 3'
	$button_mmcEvents.Name = "button_mmcEvents"
	$button_mmcEvents.Size = '149, 23'
	$button_mmcEvents.TabIndex = 0
	$button_mmcEvents.Text = "MMC: Event Viewer"
	$button_mmcEvents.UseVisualStyleBackColor = $True
	$button_mmcEvents.add_Click($button_mmcEvents_Click)
	#==================================================================
	# button_EventsLogNames
	#==================================================================
	$button_EventsLogNames.Location = '8, 32'
	$button_EventsLogNames.Name = "button_EventsLogNames"
	$button_EventsLogNames.Size = '148, 23'
	$button_EventsLogNames.TabIndex = 4
	$button_EventsLogNames.Text = "LogNames"
	$button_EventsLogNames.UseVisualStyleBackColor = $True
	$button_EventsLogNames.add_Click($button_EventsLogNames_Click)
	#==================================================================
	# tabpage_ExternalTools
	#==================================================================
	$tabpage_ExternalTools.Controls.Add($button_Rwinsta)
	$tabpage_ExternalTools.Controls.Add($button_Qwinsta)
	$tabpage_ExternalTools.Controls.Add($button_MsInfo32)
	$tabpage_ExternalTools.Controls.Add($button_DriverQuery)
	$tabpage_ExternalTools.Controls.Add($button_SystemInfoexe)
	$tabpage_ExternalTools.Controls.Add($button_PAExec)
	$tabpage_ExternalTools.Controls.Add($button_psexec)
	$tabpage_ExternalTools.Controls.Add($textbox_networktracertparam)
	$tabpage_ExternalTools.Controls.Add($button_networkTracert)
	$tabpage_ExternalTools.Controls.Add($button_networkNsLookup)
	$tabpage_ExternalTools.Controls.Add($button_networkPing)
	$tabpage_ExternalTools.Controls.Add($textbox_networkpathpingparam)
	$tabpage_ExternalTools.Controls.Add($textbox_pingparam)
	$tabpage_ExternalTools.Controls.Add($button_networkPathPing)
	$tabpage_ExternalTools.Location = '4, 22'
	$tabpage_ExternalTools.Name = "tabpage_ExternalTools"
	$tabpage_ExternalTools.Size = '1162, 111'
	$tabpage_ExternalTools.TabIndex = 9
	$tabpage_ExternalTools.Text = "ExternalTools"
	$tabpage_ExternalTools.UseVisualStyleBackColor = $True
	#==================================================================

	#==================================================================
	# button_AD_ShowGroups
	#==================================================================
	$button_AD_ShowGroups.Location = '14, 21'
	$button_AD_ShowGroups.Name = "button_AD_ShowGroups"
	$button_AD_ShowGroups.Size = '94, 23'
	$button_AD_ShowGroups.TabIndex = 1
	$button_AD_ShowGroups.Text = "Mostrar Grupos"
	$button_AD_ShowGroups.UseVisualStyleBackColor = $True
	$button_AD_ShowGroups.add_Click($button_AD_ShowGroups_Click)
	#==================================================================
	# button_AD_AddToGroup
	#==================================================================
	$button_AD_AddToGroup.Location = '14, 50'
	$button_AD_AddToGroup.Name = "button_AD_AddToGroup"
	$button_AD_AddToGroup.Size = '94, 23'
	$button_AD_AddToGroup.TabIndex = 2
	$button_AD_AddToGroup.Text = "Agregar a Grupo"
	$button_AD_AddToGroup.UseVisualStyleBackColor = $True
	$button_AD_AddToGroup.add_Click($button_AD_AddToGroup_Click)
	#==================================================================
	# button_Rwinsta
	#==================================================================
	$button_Rwinsta.Location = '289, 27'
	$button_Rwinsta.Name = "button_Rwinsta"
	$button_Rwinsta.Size = '75, 23'
	$button_Rwinsta.TabIndex = 48
	$button_Rwinsta.Text = "Rwinsta"
	$button_Rwinsta.UseVisualStyleBackColor = $True
	$button_Rwinsta.add_Click($button_Rwinsta_Click)
	#==================================================================
	# button_Qwinsta
	#==================================================================
	$button_Qwinsta.Location = '289, 3'
	$button_Qwinsta.Name = "button_Qwinsta"
	$button_Qwinsta.Size = '75, 23'
	$button_Qwinsta.TabIndex = 47
	$button_Qwinsta.Text = "Qwinsta"
	$button_Qwinsta.UseVisualStyleBackColor = $True
	$button_Qwinsta.add_Click($button_Qwinsta_Click)
	#==================================================================
	# button_MsInfo32
	#==================================================================
	$button_MsInfo32.Location = '122, 52'
	$button_MsInfo32.Name = "button_MsInfo32"
	$button_MsInfo32.Size = '75, 23'
	$button_MsInfo32.TabIndex = 46
	$button_MsInfo32.Text = "MsInfo32"
	$button_MsInfo32.UseVisualStyleBackColor = $True
	$button_MsInfo32.add_Click($button_MsInfo32_Click)
	#==================================================================
	# button_DriverQuery
	#==================================================================
	$button_DriverQuery.Location = '122, 27'
	$button_DriverQuery.Name = "button_DriverQuery"
	$button_DriverQuery.Size = '75, 23'
	$button_DriverQuery.TabIndex = 45
	$button_DriverQuery.Text = "DriverQuery"
	$button_DriverQuery.UseVisualStyleBackColor = $True
	$button_DriverQuery.add_Click($button_DriverQuery_Click)
	#==================================================================
	# button_SystemInfoexe
	#==================================================================
	$button_SystemInfoexe.Location = '122, 3'
	$button_SystemInfoexe.Name = "button_SystemInfoexe"
	$button_SystemInfoexe.Size = '75, 23'
	$button_SystemInfoexe.TabIndex = 1
	$button_SystemInfoexe.Text = "SystemInfo"
	$button_SystemInfoexe.UseVisualStyleBackColor = $True
	$button_SystemInfoexe.add_Click($button_SystemInfoexe_Click)
	#==================================================================
	# button_PAExec
	#==================================================================
	$button_PAExec.Location = '203, 27'
	$button_PAExec.Name = "button_PAExec"
	$button_PAExec.Size = '75, 23'
	$button_PAExec.TabIndex = 44
	$button_PAExec.Text = "PAExec"
	$button_PAExec.UseVisualStyleBackColor = $True
	$button_PAExec.add_Click($button_PAExec_Click)
	#==================================================================
	# button_psexec
	#==================================================================
	$button_psexec.Location = '203, 3'
	$button_psexec.Name = "button_psexec"
	$button_psexec.Size = '75, 23'
	$button_psexec.TabIndex = 43
	$button_psexec.Text = "PsExec"
	$button_psexec.UseVisualStyleBackColor = $True
	$button_psexec.add_Click($button_psexec_Click)
	#==================================================================
	# textbox_networktracertparam
	#==================================================================
	$textbox_networktracertparam.Location = '84, 78'
	$textbox_networktracertparam.Name = "textbox_networktracertparam"
	$textbox_networktracertparam.Size = '34, 20'
	$textbox_networktracertparam.TabIndex = 7
	$textbox_networktracertparam.Text = "-d"
	#==================================================================
	# button_networkTracert
	#==================================================================
	$button_networkTracert.Location = '3, 76'
	$button_networkTracert.Name = "button_networkTracert"
	$button_networkTracert.Size = '75, 23'
	$button_networkTracert.TabIndex = 6
	$button_networkTracert.Text = "Tracert"
	$button_networkTracert.UseVisualStyleBackColor = $True
	$button_networkTracert.add_Click($button_networkTracert_Click)
	#==================================================================
	# button_networkNsLookup
	#==================================================================
	$button_networkNsLookup.Location = '3, 3'
	$button_networkNsLookup.Name = "button_networkNsLookup"
	$button_networkNsLookup.Size = '75, 23'
	$button_networkNsLookup.TabIndex = 5
	$button_networkNsLookup.Text = "NsLookup"
	$button_networkNsLookup.UseVisualStyleBackColor = $True
	$button_networkNsLookup.add_Click($button_networkNsLookup_Click)
	#==================================================================
	# button_networkPing
	#==================================================================
	$button_networkPing.Location = '3, 27'
	$button_networkPing.Name = "button_networkPing"
	$button_networkPing.Size = '75, 23'
	$button_networkPing.TabIndex = 0
	$button_networkPing.Text = "Ping"
	$button_networkPing.UseVisualStyleBackColor = $True
	$button_networkPing.add_Click($button_networkPing_Click)
	#==================================================================
	# textbox_networkpathpingparam
	#==================================================================
	$textbox_networkpathpingparam.Location = '84, 54'
	$textbox_networkpathpingparam.Name = "textbox_networkpathpingparam"
	$textbox_networkpathpingparam.Size = '34, 20'
	$textbox_networkpathpingparam.TabIndex = 3
	$textbox_networkpathpingparam.Text = "-n"
	#==================================================================
	# textbox_pingparam
	#==================================================================
	$textbox_pingparam.Location = '84, 29'
	$textbox_pingparam.Name = "textbox_pingparam"
	$textbox_pingparam.Size = '34, 20'
	$textbox_pingparam.TabIndex = 1
	$textbox_pingparam.Text = "-t"
	#==================================================================
	# button_networkPathPing
	#==================================================================
	$button_networkPathPing.Location = '3, 52'
	$button_networkPathPing.Name = "button_networkPathPing"
	$button_networkPathPing.Size = '75, 23'
	$button_networkPathPing.TabIndex = 2
	$button_networkPathPing.Text = "PathPing"
	$button_networkPathPing.UseVisualStyleBackColor = $True
	$button_networkPathPing.add_Click($button_networkPathPing_Click)
	#==================================================================
	# groupbox_ComputerName
	#==================================================================
	$groupbox_ComputerName.Controls.Add($label_UptimeStatus)
	$groupbox_ComputerName.Controls.Add($textbox_computername)
	$groupbox_ComputerName.Controls.Add($label_OSStatus)
	$groupbox_ComputerName.Controls.Add($button_Check)
	$groupbox_ComputerName.Controls.Add($label_PingStatus)
	$groupbox_ComputerName.Controls.Add($label_Ping)
	$groupbox_ComputerName.Controls.Add($label_PSRemotingStatus)
	$groupbox_ComputerName.Controls.Add($label_Uptime)
	$groupbox_ComputerName.Controls.Add($label_RDPStatus)
	$groupbox_ComputerName.Controls.Add($label_OS)
	$groupbox_ComputerName.Controls.Add($label_PermissionStatus)
	$groupbox_ComputerName.Controls.Add($label_Permission)
	$groupbox_ComputerName.Controls.Add($label_PSRemoting)
	$groupbox_ComputerName.Controls.Add($label_RDP)
	$groupbox_ComputerName.Controls.Add($label_WinRM)
	$groupbox_ComputerName.Controls.Add($label_WinRMStatus)
	$groupbox_ComputerName.Dock = 'Top'
	$groupbox_ComputerName.Location = '0, 26'
	$groupbox_ComputerName.Name = "groupbox_ComputerName"
	$groupbox_ComputerName.Size = '1170, 61'
	$groupbox_ComputerName.TabIndex = 62
	$groupbox_ComputerName.TabStop = $False
	$groupbox_ComputerName.Text = "Nome de Equipo"
	#==================================================================
	# label_UptimeStatus
	#==================================================================
	$label_UptimeStatus.Location = '614, 33'
	$label_UptimeStatus.Name = "label_UptimeStatus"
	$label_UptimeStatus.Size = '539, 19'
	$label_UptimeStatus.TabIndex = 61
	#==================================================================
	# textbox_computername
	#==================================================================
	$textbox_computername.AutoCompleteMode = 'SuggestAppend'
	$textbox_computername.AutoCompleteSource = 'CustomSource'
	$textbox_computername.BackColor = 'LemonChiffon'
	$textbox_computername.BorderStyle = 'FixedSingle'
	$textbox_computername.CharacterCasing = 'Upper'
	$textbox_computername.Font = "Consolas, 18pt"
	$textbox_computername.ForeColor = 'WindowText'
	$textbox_computername.Location = '6, 14'
	$textbox_computername.Name = "textbox_computername"
	$textbox_computername.Size = '209, 36'
	$textbox_computername.TabIndex = 2
	$textbox_computername.Text = "LOCALHOST"
	$textbox_computername.TextAlign = 'Center'
	$tooltipinfo.SetToolTip($textbox_computername, "Por favor, introduce un nome de equipo")
	$textbox_computername.add_TextChanged($textbox_computername_TextChanged)
	$textbox_computername.add_KeyPress($textbox_computername_KeyPress)
	$textbox_computername.Add_KeyDown({
		if ($_.KeyCode -eq "Enter") {
			# Solo ejecutar si no hay una recolección en curso
			if ($button_Check.Enabled -and -not $global:StreamRunning) {
				$button_Check.PerformClick()
			}
		}
	})
	#==================================================================
	# label_OSStatus
	#==================================================================
	$label_OSStatus.Location = '614, 16'
	$label_OSStatus.Name = "label_OSStatus"
	$label_OSStatus.Size = '60, 16'
	$label_OSStatus.TabIndex = 60
	#==================================================================
	# button_Check
	#==================================================================
	$button_Check.Font = "Microsoft Sans Serif, 8.25pt, style=Bold"
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\button_Check.ps1')
	$button_Check.ImageAlign = 'MiddleLeft'
	$button_Check.Location = '220, 15'
	$button_Check.Name = "button_Check"
	$button_Check.Size = '105, 35'
	$button_Check.TabIndex = 51
	$button_Check.Text = "Recolle Datos"
	$button_Check.TextAlign = "MiddleRight"
	$tooltipinfo.SetToolTip($button_Check, "Comproba a conectividade e a información básica")
	$button_Check.UseVisualStyleBackColor = $True
	$button_Check.add_Click($button_Check_Click)
	#==================================================================
	# label_PingStatus
	#==================================================================
	$label_PingStatus.Location = '399, 17'
	$label_PingStatus.Name = "label_PingStatus"
	$label_PingStatus.Size = '33, 16'
	$label_PingStatus.TabIndex = 50
	#==================================================================
	# label_Ping
	#==================================================================
	$label_Ping.Font = "Trebuchet MS, 8.25pt, style=Underline"
	$label_Ping.Location = '333, 16'
	$label_Ping.Name = "label_Ping"
	$label_Ping.Size = '33, 16'
	$label_Ping.TabIndex = 49
	$label_Ping.Text = "Ping:"
	#==================================================================
	# label_PSRemotingStatus
	#==================================================================
	$label_PSRemotingStatus.Location = '496, 33'
	$label_PSRemotingStatus.Name = "label_PSRemotingStatus"
	$label_PSRemotingStatus.Size = '69, 14'
	$label_PSRemotingStatus.TabIndex = 57
	#==================================================================
	# label_Uptime
	#==================================================================
	$label_Uptime.Font = "Trebuchet MS, 8.25pt, style=Underline"
	$label_Uptime.Location = '571, 32'
	$label_Uptime.Name = "label_Uptime"
	$label_Uptime.Size = '50, 20'
	$label_Uptime.TabIndex = 59
	$label_Uptime.Text = "ON:"
	#==================================================================
	# label_RDPStatus
	#==================================================================
	$label_RDPStatus.Location = '496, 16'
	$label_RDPStatus.Name = "label_RDPStatus"
	$label_RDPStatus.Size = '69, 19'
	$label_RDPStatus.TabIndex = 56
	#==================================================================
	# label_OS / WMI
	#==================================================================
	$label_OS.Font = "Trebuchet MS, 8.25pt, style=Underline"
	$label_OS.Location = '571, 16'
	$label_OS.Name = "label_OS"
	$label_OS.Size = '37, 20'
	$label_OS.TabIndex = 58
	$label_OS.Text = "WMI:"
	#==================================================================
	# label_PermissionStatus
	#==================================================================
	$label_PermissionStatus.Location = '399, 33'
	$label_PermissionStatus.Name = "label_PermissionStatus"
	$label_PermissionStatus.Size = '33, 16' #'33, 16'
	$label_PermissionStatus.TabIndex = 53
	#==================================================================
	# label_Permission
	#==================================================================
	$label_Permission.Font = "Trebuchet MS, 8.25pt, style=Underline"
	$label_Permission.Location = '333, 32'
	$label_Permission.Name = "label_Permission"
	$label_Permission.Size = '72, 20'
	$label_Permission.TabIndex = 52
	$label_Permission.Text = "Permisos:"
	#==================================================================
	# label_PSRemoting #VNC
	#==================================================================
	$label_PSRemoting.Font = "Trebuchet MS, 8.25pt, style=Underline"
	$label_PSRemoting.Location = '431, 32'
	$label_PSRemoting.Name = "label_PSRemoting"
	$label_PSRemoting.Size = '75, 20'
	$label_PSRemoting.TabIndex = 55
	$label_PSRemoting.Text = "Estado VNC:"
	#==================================================================
	# label_RDP
	#==================================================================
	$label_RDP.Font = "Trebuchet MS, 8.25pt, style=Underline"
	$label_RDP.Location = '431, 16'
	$label_RDP.Name = "label_RDP"
	$label_RDP.Size = '37, 20'
	$label_RDP.TabIndex = 54
	$label_RDP.Text = "RDP:"
	#==================================================================
	# label_WinRM
	#==================================================================
	$label_WinRM.Font = "Trebuchet MS, 8.25pt, style=Underline"
	$label_WinRM.Location = '700, 16'
	$label_WinRM.Name = "label_WinRM"
	$label_WinRM.Size = '50, 20'
	$label_WinRM.TabIndex = 61
	$label_WinRM.Text = "WinRM:"
	#==================================================================
	# label_WinRMStatus
	#==================================================================
	$label_WinRMStatus.Location = '755, 16'
	$label_WinRMStatus.Name = "label_WinRMStatus"
	$label_WinRMStatus.Size = '150, 16'
	$label_WinRMStatus.TabIndex = 62
	$label_WinRMStatus.Text = ""
	#==================================================================
	# richtextbox_Logs
	#==================================================================
	$richtextbox_Logs.BackColor = 'InactiveBorder'
	$richtextbox_Logs.Dock = 'Bottom'
	$richtextbox_Logs.Font = "Consolas, 8.25pt"
	$richtextbox_Logs.ForeColor = 'Green'
	$richtextbox_Logs.Location = '0, 623'
	$richtextbox_Logs.Name = "richtextbox_Logs"
	$richtextbox_Logs.ReadOnly = $True
	$richtextbox_Logs.Size = '1170, 70'
	$richtextbox_Logs.TabIndex = 35
	$richtextbox_Logs.Text = ""
	$richtextbox_Logs.add_TextChanged($richtextbox_Logs_TextChanged)
	#==================================================================
	# statusbar1
	#==================================================================
	$statusbar1.Location = '0, 693'
	$statusbar1.Name = "statusbar1"
	$statusbar1.Size = '1170, 26'
	$statusbar1.TabIndex = 16
	#==================================================================
	# menustrip_principal - MENU1_PRINCIPAL
	#==================================================================
	$menustrip_principal.Font = "Trebuchet MS, 9pt"
	[void]$menustrip_principal.Items.Add($ToolStripMenuItem_AdminArsenal)
	[void]$menustrip_principal.Items.Add($ToolStripMenuItem_localhost)
	[void]$menustrip_principal.Items.Add($ToolStripMenuItem_scripts)
	[void]$menustrip_principal.Items.Add($ToolStripMenuItem_apps)
	[void]$menustrip_principal.Items.Add($ToolStripMenuItem_configuracion)
	[void]$menustrip_principal.Items.Add($ToolStripMenuItem_about)
	$menustrip_principal.Location = '0, 0'
	$menustrip_principal.Name = "menustrip_principal"
	$menustrip_principal.Size = '1170, 26'
	$menustrip_principal.TabIndex = 1
	$menustrip_principal.Text = "menustrip1"
	#==================================================================
	# ToolStripMenuItem_AdminArsenal - Desplegable Ferramentas Admin
	#==================================================================
	# grupo 1: administración
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_ADPrinters)        #ADMINISTRACIÓN IMPRESORAS
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_ADSearchDialog)     #DIRECTORIO ACTIVO
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_GeneratePassword)   #EXCHANGE
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_DHCP)              #DHCP
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_TerminalAdmin)     #WSUS
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_InternetExplorer)  #VDIS
	# separador
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($toolstripseparator4)
	# grupo 2: consolas
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_CommandPrompt)     #CMD
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_Powershell)        #POWERSHELL
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_PowershellISE)     #EDITOR POWERSHELL
	# separador
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($toolstripseparator5)
	# grupo 3: utilidades finales
	[void]$ToolStripMenuItem_AdminArsenal.DropDownItems.Add($ToolStripMenuItem_sysInternals)     #SYSINTERNALS (opcional)
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_AdminArsenal.ps1')
	$ToolStripMenuItem_AdminArsenal.Name = "ToolStripMenuItem_AdminArsenal"
	$ToolStripMenuItem_AdminArsenal.Size = '109, 22'
	$ToolStripMenuItem_AdminArsenal.Text = "Administrador"
	#==================================================================
	# ToolStripMenuItem_CommandPrompt - Consola de Comandos
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_CommandPrompt.ps1')
	$ToolStripMenuItem_CommandPrompt.Name = "ToolStripMenuItem_CommandPrompt"
	$ToolStripMenuItem_CommandPrompt.Size = '290, 22'
	$ToolStripMenuItem_CommandPrompt.Text = "Consola de Comandos"
	$ToolStripMenuItem_CommandPrompt.add_Click($ToolStripMenuItem_CommandPrompt_Click)
	#==================================================================
	# ToolStripMenuItem_Powershell - Botón Poweshell
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_Powershell.ps1')
	$ToolStripMenuItem_Powershell.Name = "ToolStripMenuItem_Powershell"
	$ToolStripMenuItem_Powershell.Size = '290, 22'
	$ToolStripMenuItem_Powershell.Text = "Consola Powershell"
	$ToolStripMenuItem_Powershell.add_Click($ToolStripMenuItem_Powershell_Click)
	#==================================================================
	# ToolStripMenuItem_localhost
	#==================================================================
	# Menú Localhost reorganizado según petición
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_systemInformationMSinfo32exe)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_systemproperties)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($toolstripseparator3)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_compmgmt)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_devicemanager)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_taskManager)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_services)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_regedit)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_mmc)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($toolstripseparator1)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_netstatsListening)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_registeredSnappins)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_resetCredenciaisVNC)
	[void]$ToolStripMenuItem_localhost.DropDownItems.Add($ToolStripMenuItem_otherLocalTools)
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_localhost.ps1')
	$ToolStripMenuItem_localhost.Name = "ToolStripMenuItem_localhost"
	$ToolStripMenuItem_localhost.Size = '90, 22'
	$ToolStripMenuItem_localhost.Text = "LocalHost"
	#==================================================================
	# ToolStripMenuItem_compmgmt
	#==================================================================
	$ToolStripMenuItem_compmgmt.Name = "ToolStripMenuItem_compmgmt"
	$ToolStripMenuItem_compmgmt.Size = '278, 22'
	$ToolStripMenuItem_compmgmt.Text = "Administrador De Equipo"
	$ToolStripMenuItem_compmgmt.add_Click($ToolStripMenuItem_compmgmt_Click)
	#==================================================================
	# ToolStripMenuItem_taskManager
	#==================================================================
	$ToolStripMenuItem_taskManager.Name = "ToolStripMenuItem_taskManager"
	$ToolStripMenuItem_taskManager.Size = '278, 22'
	$ToolStripMenuItem_taskManager.Text = "Administrador De Tareas"
	$ToolStripMenuItem_taskManager.add_Click($ToolStripMenuItem_taskManager_Click)
	#==================================================================
	# ToolStripMenuItem_services
	#==================================================================
	$ToolStripMenuItem_services.Name = "ToolStripMenuItem_services"
	$ToolStripMenuItem_services.Size = '278, 22'
	$ToolStripMenuItem_services.Text = "Servicios"
	$ToolStripMenuItem_services.add_Click($ToolStripMenuItem_services_Click)
	#==================================================================
	# ToolStripMenuItem_regedit
	#==================================================================
	$ToolStripMenuItem_regedit.Name = "ToolStripMenuItem_regedit"
	$ToolStripMenuItem_regedit.Size = '278, 22'
	$ToolStripMenuItem_regedit.Text = "Registro"
	$ToolStripMenuItem_regedit.add_Click($ToolStripMenuItem_regedit_Click)
	#==================================================================
	# ToolStripMenuItem_mmc
	#==================================================================
	$ToolStripMenuItem_mmc.Name = "ToolStripMenuItem_mmc"
	$ToolStripMenuItem_mmc.Size = '278, 22'
	$ToolStripMenuItem_mmc.Text = "Microsoft Management Console"
	$ToolStripMenuItem_mmc.add_Click($ToolStripMenuItem_mmc_Click)
	#==================================================================
	# ToolStripMenuItem_about
	#==================================================================
	[void]$ToolStripMenuItem_about.DropDownItems.Add($ToolStripMenuItem_AboutInfo)
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_about.ps1')
	$ToolStripMenuItem_about.Name = "ToolStripMenuItem_about"
	$ToolStripMenuItem_about.Size = '69, 22'
	$ToolStripMenuItem_about.Text = "Acerca de" #"About"
	#==================================================================
	# ToolStripMenuItem_AboutInfo
	#==================================================================
	$ToolStripMenuItem_AboutInfo.Name = "ToolStripMenuItem_AboutInfo"
	$ToolStripMenuItem_AboutInfo.Size = '211, 22'
	$ToolStripMenuItem_AboutInfo.Text = "About $ApplicationName"
	$ToolStripMenuItem_AboutInfo.add_Click($ToolStripMenuItem_AboutInfo_Click)
	#==================================================================
	# contextmenustripServer
	#==================================================================
	[void]$contextmenustripServer.Items.Add($ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools)
	[void]$contextmenustripServer.Items.Add($ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC)
	$contextmenustripServer.Name = "contextmenustripServer"
	$contextmenustripServer.ShowImageMargin = $False
	$contextmenustripServer.Size = '79, 26'
	#==================================================================
	# ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools
	#==================================================================
	[void]$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools.DropDownItems.Add($ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Ping)
	[void]$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools.DropDownItems.Add($ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_RDP)
	[void]$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools.DropDownItems.Add($ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Qwinsta)
	[void]$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools.DropDownItems.Add($ToolStripMenuItem_rwinsta)
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools.Name = "ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools.Size = '78, 22'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Tools.Text = "Ferramentas" #Tools
	#==================================================================
	# ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC
	#==================================================================
	[void]$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC.DropDownItems.Add($ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_compmgmt)
	[void]$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC.DropDownItems.Add($ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_services)
	[void]$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC.DropDownItems.Add($ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_eventvwr)
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC.Name = "ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC.Size = '130, 22'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_ConsolesMMC.Text = "Consolas MMC" #Consoles MMC
	#==================================================================
	# ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Ping
	#==================================================================
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Ping.Name = "ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Ping"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Ping.Size = '117, 22'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Ping.Text = "Ping"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Ping.add_Click($button_ping_Click)
	#==================================================================
	# ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_RDP
	#==================================================================
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_RDP.Name = "ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_RDP"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_RDP.Size = '117, 22'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_RDP.Text = "RDP"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_RDP.add_Click($button_remot_Click)
	#==================================================================
	# ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_compmgmt
	#==================================================================
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_compmgmt.Name = "ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_compmgmt"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_compmgmt.Size = '202, 22'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_compmgmt.Text = "Administración de Equipos" #"Computer Management"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_compmgmt.add_Click($button_mmcCompmgmt_Click)
	#==================================================================
	# ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_services
	#==================================================================
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_services.Name = "ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_services"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_services.Size = '202, 22'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_services.Text = "Servicios" #"Services"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_services.add_Click($button_mmcServices_Click)
	#==================================================================
	# ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_eventvwr
	#==================================================================
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_eventvwr.Name = "ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_eventvwr"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_eventvwr.Size = '197, 22'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_eventvwr.Text = "Visor de Eventos" #"Events Viewer"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_eventvwr.add_Click($button_mmcEvents_Click)
	#==================================================================
	# ToolStripMenuItem_InternetExplorer
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_InternetExplorer.ps1')
	$ToolStripMenuItem_InternetExplorer.Name = "ToolStripMenuItem_InternetExplorer"
	$ToolStripMenuItem_InternetExplorer.Size = '290, 22'
	$ToolStripMenuItem_InternetExplorer.Text = "Portal Web"
	$ToolStripMenuItem_InternetExplorer.add_Click($ToolStripMenuItem_InternetExplorer_Click)
	#==================================================================
	# ToolStripMenuItem_TerminalAdmin
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_TerminalAdmin.ps1')
	$ToolStripMenuItem_TerminalAdmin.Name = "ToolStripMenuItem_TerminalAdmin"
	$ToolStripMenuItem_TerminalAdmin.Size = '290, 22'
	$ToolStripMenuItem_TerminalAdmin.Text = "WSUS" #Terminal Admin (TsAdmin)
	$ToolStripMenuItem_TerminalAdmin.add_Click($ToolStripMenuItem_TerminalAdmin_Click)
	#==================================================================
	# ToolStripMenuItem_ADSearchDialog
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_ADSearchDialog.ps1')
	$ToolStripMenuItem_ADSearchDialog.Name = "ToolStripMenuItem_ADSearchDialog"
	$ToolStripMenuItem_ADSearchDialog.Size = '290, 22'
	$ToolStripMenuItem_ADSearchDialog.Text = "Directorio Activo"
	$ToolStripMenuItem_ADSearchDialog.add_Click($ToolStripMenuItem_ADSearchDialog_Click)
	#==================================================================
	# ToolStripMenuItem_ADPrinters
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_ADPrinters.ps1')
	$ToolStripMenuItem_ADPrinters.Name = "ToolStripMenuItem_ADPrinters"
	$ToolStripMenuItem_ADPrinters.Size = '290, 22'
	$ToolStripMenuItem_ADPrinters.Text = "Administración Impresoras"
	$ToolStripMenuItem_ADPrinters.add_Click($ToolStripMenuItem_ADPrinters_Click)

	#==================================================================
	# ToolStripMenuItem_DHCP
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_DHCP.ps1')
	$ToolStripMenuItem_DHCP.Name = "ToolStripMenuItem_DHCP"
	$ToolStripMenuItem_DHCP.Size = '290, 22'
	$ToolStripMenuItem_DHCP.Text = "DHCP"
	$ToolStripMenuItem_DHCP.add_Click($ToolStripMenuItem_DHCP_Click)
	#==================================================================
	#==================================================================
	#==================================================================
	# ToolStripMenuItem_systemInformationMSinfo32exe
	#==================================================================
	$ToolStripMenuItem_systemInformationMSinfo32exe.Name = "ToolStripMenuItem_systemInformationMSinfo32exe"
	$ToolStripMenuItem_systemInformationMSinfo32exe.Size = '278, 22'
	$ToolStripMenuItem_systemInformationMSinfo32exe.Text = "Información do Sistema" #"System Information (MSinfo32.exe)"
	$ToolStripMenuItem_systemInformationMSinfo32exe.add_Click($ToolStripMenuItem_systemInformationMSinfo32exe_Click)
	#==================================================================
	# ToolStripMenuItem_netstatsListening
	#==================================================================
	$ToolStripMenuItem_netstatsListening.Name = "ToolStripMenuItem_netstatsListening"
	$ToolStripMenuItem_netstatsListening.Size = '278, 22'
	$ToolStripMenuItem_netstatsListening.Text = "Estadísticas de Rede | Portos Escoitando"
	$ToolStripMenuItem_netstatsListening.add_Click($ToolStripMenuItem_netstatsListening_Click)
	#==================================================================
	# ToolStripMenuItem_registeredSnappins
	#==================================================================
	$ToolStripMenuItem_registeredSnappins.Name = "ToolStripMenuItem_registeredSnappins"
	$ToolStripMenuItem_registeredSnappins.Size = '278, 22'
	$ToolStripMenuItem_registeredSnappins.Text = "Snapins - Modulos Disponibles"
	$ToolStripMenuItem_registeredSnappins.add_Click($ToolStripMenuItem_registeredSnappins_Click)
	#==================================================================
	# ToolStripMenuItem_otherLocalTools
	#==================================================================
	# orden personalizado según especificaciones recientes
	[void]$ToolStripMenuItem_otherLocalTools.DropDownItems.Add($ToolStripMenuItem_diskManagement)
	[void]$ToolStripMenuItem_otherLocalTools.DropDownItems.Add($ToolStripMenuItem_localSecuritySettings)
	[void]$ToolStripMenuItem_otherLocalTools.DropDownItems.Add($ToolStripMenuItem_groupPolicyEditor)
	[void]$ToolStripMenuItem_otherLocalTools.DropDownItems.Add($ToolStripMenuItem_performanceMonitor)
	[void]$ToolStripMenuItem_otherLocalTools.DropDownItems.Add($ToolStripMenuItem_scheduledTasks)
	[void]$ToolStripMenuItem_otherLocalTools.DropDownItems.Add($ToolStripMenuItem_certificateManager)
	[void]$ToolStripMenuItem_otherLocalTools.DropDownItems.Add($ToolStripMenuItem_localUsersAndGroups)
	[void]$ToolStripMenuItem_otherLocalTools.DropDownItems.Add($ToolStripMenuItem_sharedFolders)
	$ToolStripMenuItem_otherLocalTools.Name = "ToolStripMenuItem_otherLocalTools"
	$ToolStripMenuItem_otherLocalTools.Size = '278, 22'
	$ToolStripMenuItem_otherLocalTools.Text = "Outras Aplicacións de Windows"
	#==================================================================
	#==================================================================
	#==================================================================
	# ToolStripMenuItem_certificateManager
	#==================================================================
	$ToolStripMenuItem_certificateManager.Name = "ToolStripMenuItem_certificateManager"
	$ToolStripMenuItem_certificateManager.Size = '311, 22'
	$ToolStripMenuItem_certificateManager.Text = "Almacén de Certificados" #"Device Manager"
	$ToolStripMenuItem_certificateManager.add_Click($ToolStripMenuItem_certificateManager_Click)
	#==================================================================
	# ToolStripMenuItem_devicemanager
	#==================================================================
	$ToolStripMenuItem_devicemanager.Name = "ToolStripMenuItem_devicemanager"
	$ToolStripMenuItem_devicemanager.Size = '278, 22'
	$ToolStripMenuItem_devicemanager.Text = "Administrador De Dispositivos" #"Device Manager"
	$ToolStripMenuItem_devicemanager.add_Click($ToolStripMenuItem_devicemanager_Click)
	#==================================================================
	# ToolStripMenuItem_systemproperties
	#==================================================================
	$ToolStripMenuItem_systemproperties.Name = "ToolStripMenuItem_systemproperties"
	$ToolStripMenuItem_systemproperties.Size = '278, 22'
	$ToolStripMenuItem_systemproperties.Text = "Propiedades do Sistema" #"System Properties"
	$ToolStripMenuItem_systemproperties.add_Click($ToolStripMenuItem_systemproperties_Click)
	#==================================================================
	# ToolStripMenuItem_sharedFolders
	#==================================================================
	$ToolStripMenuItem_sharedFolders.Name = "ToolStripMenuItem_sharedFolders"
	$ToolStripMenuItem_sharedFolders.Size = '311, 22'
	$ToolStripMenuItem_sharedFolders.Text = "Recursos Compartidos" #"Shared Folders"
	$ToolStripMenuItem_sharedFolders.add_Click($ToolStripMenuItem_sharedFolders_Click)
	#==================================================================
	# ToolStripMenuItem_performanceMonitor
	#==================================================================
	$ToolStripMenuItem_performanceMonitor.Name = "ToolStripMenuItem_performanceMonitor"
	$ToolStripMenuItem_performanceMonitor.Size = '311, 22'
	$ToolStripMenuItem_performanceMonitor.Text = "Monitor de Rendemento" #"Performance Monitor"
	$ToolStripMenuItem_performanceMonitor.add_Click($ToolStripMenuItem_performanceMonitor_Click)
	#==================================================================
	#==================================================================
	# ToolStripMenuItem_groupPolicyEditor
	#==================================================================
	$ToolStripMenuItem_groupPolicyEditor.Name = "ToolStripMenuItem_groupPolicyEditor"
	$ToolStripMenuItem_groupPolicyEditor.Size = '311, 22'
	$ToolStripMenuItem_groupPolicyEditor.Text = "Directivas Locales"
	$ToolStripMenuItem_groupPolicyEditor.add_Click($ToolStripMenuItem_groupPolicyEditor_Click)
	#==================================================================
	# ToolStripMenuItem_localUsersAndGroups
	#==================================================================
	$ToolStripMenuItem_localUsersAndGroups.Name = "ToolStripMenuItem_localUsersAndGroups"
	$ToolStripMenuItem_localUsersAndGroups.Size = '311, 22'
	$ToolStripMenuItem_localUsersAndGroups.Text = "Usuarios & Grupos Locales"
	$ToolStripMenuItem_localUsersAndGroups.add_Click($ToolStripMenuItem_localUsersAndGroups_Click)
	#==================================================================
	# ToolStripMenuItem_diskManagement
	#==================================================================
	$ToolStripMenuItem_diskManagement.Name = "ToolStripMenuItem_diskManagement"
	$ToolStripMenuItem_diskManagement.Size = '311, 22'
	$ToolStripMenuItem_diskManagement.Text = "Administración De Discos" #nuevo nombre según petición
	$ToolStripMenuItem_diskManagement.add_Click($ToolStripMenuItem_diskManagement_Click)
	#==================================================================
	# ToolStripMenuItem_localSecuritySettings
	#==================================================================
	$ToolStripMenuItem_localSecuritySettings.Name = "ToolStripMenuItem_localSecuritySettings"
	$ToolStripMenuItem_localSecuritySettings.Size = '311, 22'
	$ToolStripMenuItem_localSecuritySettings.Text = "Directivas De Seguridad"
	$ToolStripMenuItem_localSecuritySettings.add_Click($ToolStripMenuItem_localSecuritySettings_Click)
	#==================================================================
	#==================================================================
	# ToolStripMenuItem_scheduledTasks
	#==================================================================
	$ToolStripMenuItem_scheduledTasks.Name = "ToolStripMenuItem_scheduledTasks"
	$ToolStripMenuItem_scheduledTasks.Size = '311, 22'
	$ToolStripMenuItem_scheduledTasks.Text = "Programador De Tareas"
	$ToolStripMenuItem_scheduledTasks.add_Click($ToolStripMenuItem_scheduledTasks_Click)
	#==================================================================
	# ToolStripMenuItem_PowershellISE
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_PowershellISE.ps1')
	$ToolStripMenuItem_PowershellISE.Name = "ToolStripMenuItem_PowershellISE"
	$ToolStripMenuItem_PowershellISE.Size = '290, 22'
	$ToolStripMenuItem_PowershellISE.Text = "Editor Poweshell"
	$ToolStripMenuItem_PowershellISE.add_Click($ToolStripMenuItem_PowershellISE_Click)
	#==================================================================
	#==================================================================
	# errorprovider1
	#==================================================================
	$errorprovider1.ContainerControl = $form_MainForm
	#==================================================================
	# tooltipinfo
	#==================================================================   
	[void]$ToolStripMenuItem_sysInternals.DropDownItems.Add($ToolStripMenuItem_adExplorer)
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_sysInternals.ps1')
	$ToolStripMenuItem_sysInternals.Name = "ToolStripMenuItem_sysInternals"
	$ToolStripMenuItem_sysInternals.Size = '290, 22'
	$ToolStripMenuItem_sysInternals.Text = "SysInternals"
	#==================================================================
	# ToolStripMenuItem_adExplorer
	#==================================================================
	$ToolStripMenuItem_adExplorer.Name = "ToolStripMenuItem_adExplorer"
	$ToolStripMenuItem_adExplorer.Size = '152, 22'
	$ToolStripMenuItem_adExplorer.Text = "AdExplorer"
	$ToolStripMenuItem_adExplorer.add_Click($ToolStripMenuItem_adExplorer_Click)
	#==================================================================
	# ToolStripMenuItem_resetCredenciaisVNC
	#==================================================================
	$ToolStripMenuItem_resetCredenciaisVNC.Name = "ToolStripMenuItem_resetCredenciaisVNC"
	$ToolStripMenuItem_resetCredenciaisVNC.Size = '278, 22'
	$ToolStripMenuItem_resetCredenciaisVNC.Text = "Reintroducir Credenciais VNC"
	$ToolStripMenuItem_resetCredenciaisVNC.add_Click($ToolStripMenuItem_resetCredenciaisVNC_Click)
	#==================================================================
	# ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Qwinsta
	#==================================================================
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Qwinsta.Name = "ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Qwinsta"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Qwinsta.Size = '117, 22'
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Qwinsta.Text = "Qwinsta"
	$ContextMenuStripItem_consoleToolStripMenuItem_ComputerName_Qwinsta.add_Click($button_Qwinsta_Click)
	#==================================================================
	# ToolStripMenuItem_rwinsta
	#==================================================================
	$ToolStripMenuItem_rwinsta.Name = "ToolStripMenuItem_rwinsta"
	$ToolStripMenuItem_rwinsta.Size = '152, 22'
	$ToolStripMenuItem_rwinsta.Text = "Rwinsta"
	$ToolStripMenuItem_rwinsta.add_Click($button_Rwinsta_Click)
	#==================================================================
	# ToolStripMenuItem_GeneratePassword
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_GeneratePassword.ps1')
	$ToolStripMenuItem_GeneratePassword.Name = "ToolStripMenuItem_GeneratePassword"
	$ToolStripMenuItem_GeneratePassword.Size = '290, 22'
	$ToolStripMenuItem_GeneratePassword.Text = "Exchange" #"Generate a password"
	$ToolStripMenuItem_GeneratePassword.add_Click($button_PasswordGen_Click)
	#==================================================================
	# ToolStripMenuItem_scripts
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_scripts.ps1')
	$ToolStripMenuItem_scripts.Name = "ToolStripMenuItem_scripts"
	$ToolStripMenuItem_scripts.Size = '74, 22'
	$ToolStripMenuItem_scripts.Text = "Scripts"
	#==================================================================
	# ToolStripMenuItem_apps (Aplicacións) - contenido gestionado por sections\Aplicacions.psm1
	#==================================================================
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\ToolStripMenuItem_apps.ps1')
	$ToolStripMenuItem_apps.Name = "ToolStripMenuItem_apps"
	$ToolStripMenuItem_apps.Size = '74, 22'
	$ToolStripMenuItem_apps.Text = "Aplicacións"
	#==================================================================
	# ToolStripMenuItem_configuracion - contenido gestionado por sections\Configuracion.psm1
	#==================================================================
	$ToolStripMenuItem_configuracion.Name = "ToolStripMenuItem_configuracion"
	$ToolStripMenuItem_configuracion.Size = '95, 22'
	$ToolStripMenuItem_configuracion.Text = "⚙️  Configuración"
	#==================================================================	
	#==================================================================
	# imagelistAnimation
	#==================================================================
	$Formatter_binaryFomatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
	. (Join-Path $Global:ScriptRoot 'icos\BinaryData\imagelistAnimation.ps1')
	$imagelistAnimation.ImageStream = $Formatter_binaryFomatter.Deserialize($System_IO_MemoryStream)
	$Formatter_binaryFomatter = $null
	$System_IO_MemoryStream = $null
	$imagelistAnimation.TransparentColor = 'Transparent'
	#==================================================================
	# timerCheckJob
	#==================================================================
	$timerCheckJob.add_Tick($timerCheckJob_Tick2)
	#endregion Generated Form Code
	#==================================================================
	#Save the initial state of the form
	$InitialFormWindowState = $form_MainForm.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$form_MainForm.add_Load($Form_StateCorrection_Load)
	# Reconstruir el menú Scripts con orden correcto, categorías y soporte de scripts en caliente
	# Nota: Initialize-ScriptsMenu limpia los items inline y reconstruye el menú completo.
	if (Get-Command 'Initialize-ScriptsMenu' -ErrorAction SilentlyContinue) {
		Initialize-ScriptsMenu -Menu $ToolStripMenuItem_scripts
	}

	# Inicializar el menú Aplicacións con todas las apps y ferramentas
	if (Get-Command 'Initialize-AplicacionsMenu' -ErrorAction SilentlyContinue) {
		Initialize-AplicacionsMenu -Menu $ToolStripMenuItem_apps
	}

	# Inicializar el menú Configuración con las opciones de gestión de BD
	if (Get-Command 'Initialize-ConfiguracionMenu' -ErrorAction SilentlyContinue) {
		Initialize-ConfiguracionMenu -Menu $ToolStripMenuItem_configuracion
	}

	# Registrar callback de logging para SharedDataManager (antes de mostrar el form)
	if (Get-Command 'Register-ServerLogCallback' -ErrorAction SilentlyContinue) {
		Register-ServerLogCallback -Callback { param($m) Add-Logs -text $m -ErrorAction SilentlyContinue }
	}

	#Clean up the control events
	$form_MainForm.add_FormClosed($Form_Cleanup_FormClosed)
	#Store the control values when form is closing
	$form_MainForm.add_Closing($Form_StoreValues_Closing)
	#Show the Form
	return $form_MainForm.ShowDialog()
}
#Start the application
Main ($CommandLine)

