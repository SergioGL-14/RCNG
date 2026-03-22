#requires -version 5.1

# Launcher de despliegue.
# Prepara la copia local, compara versiones contra el servidor y deja lista
# la contraseña temporal de sesión antes de arrancar NRC_APP.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configuración de rutas y patrón del ejecutable.
$LOCAL_PATH  = "C:\NRC_APP"
$SERVER_PATH = "\\server\share\NRC_APP"
$EXE_PATTERN = "launcher_v*.exe"
$script:LauncherDefaults = @{
    SharedServerBase   = $SERVER_PATH
    SupportDisplayName = 'NRC_APP Support'
}
$script:LauncherSettings = $null

function Get-LauncherSettings {
    if ($null -ne $script:LauncherSettings) {
        return $script:LauncherSettings
    }

    $settings = @{}
    foreach ($key in $script:LauncherDefaults.Keys) {
        $settings[$key] = [string]$script:LauncherDefaults[$key]
    }

    $settingsPath = Join-Path $LOCAL_PATH 'database\appsettings.json'
    if (Test-Path $settingsPath) {
        try {
            $raw = Get-Content -Path $settingsPath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $json = $raw | ConvertFrom-Json -ErrorAction Stop
                foreach ($key in $script:LauncherDefaults.Keys) {
                    if ($json.PSObject.Properties[$key] -and -not [string]::IsNullOrWhiteSpace([string]$json.$key)) {
                        $settings[$key] = [string]$json.$key
                    }
                }
            }
        } catch {
        }
    }

    $script:LauncherSettings = $settings
    return $script:LauncherSettings
}

function Get-LauncherSettingValue {
    param(
        [Parameter(Mandatory=$true)][string]$Key,
        [string]$DefaultValue = ''
    )

    $settings = Get-LauncherSettings
    if ($settings.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace([string]$settings[$Key])) {
        return [string]$settings[$Key]
    }
    if (-not [string]::IsNullOrWhiteSpace($DefaultValue)) {
        return $DefaultValue
    }
    return ''
}

$SERVER_PATH = Get-LauncherSettingValue -Key 'SharedServerBase' -DefaultValue $SERVER_PATH

# La clave sale del MachineGuid para que el launcher y la app compartan el mismo secreto local.
# Asi temp.pass solo puede reutilizarse en este mismo equipo.
function Get-PassKey {
    $guid   = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Cryptography' -Name MachineGuid).MachineGuid
    $salt   = [System.Text.Encoding]::UTF8.GetBytes('NRC_Pass_v6_Salt_2026')
    $derive = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($guid, $salt, 10000)
    $key    = $derive.GetBytes(32)
    $derive.Dispose()
    return $key
}

# Diálogo de autenticación para acceder al recurso de despliegue.
function Show-CredentialsDialog {
    param()
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Network Remote Control - Autenticación"
    $form.Size = New-Object System.Drawing.Size(400,300)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(0x2E,0x3B,0x4E)

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Text = Get-LauncherSettingValue -Key 'SupportDisplayName' -DefaultValue 'NRC_APP Support'
    $lblHeader.Font = New-Object System.Drawing.Font('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
    $lblHeader.ForeColor = [System.Drawing.Color]::White
    $lblHeader.AutoSize = $true
    $lblHeader.Location = New-Object System.Drawing.Point(120,15)
    $form.Controls.Add($lblHeader)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = "Usuario:"
    $lblUser.Font = New-Object System.Drawing.Font('Segoe UI',12)
    $lblUser.ForeColor = [System.Drawing.Color]::White
    $lblUser.Location = New-Object System.Drawing.Point(8,68)
    $form.Controls.Add($lblUser)

    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Font = New-Object System.Drawing.Font('Segoe UI',12)
    $txtUser.Size = New-Object System.Drawing.Size(240,25)
    $txtUser.Location = New-Object System.Drawing.Point(110,68)
    $form.Controls.Add($txtUser)

    $lblPass = New-Object System.Windows.Forms.Label
    $lblPass.Text = [Text.Encoding]::UTF8.GetString([Text.Encoding]::UTF8.GetBytes("Contraseña:"))
    $lblPass.Font = New-Object System.Drawing.Font('Segoe UI',12)
    $lblPass.ForeColor = [System.Drawing.Color]::White
    $lblPass.Location = New-Object System.Drawing.Point(8,115)
    $form.Controls.Add($lblPass)

    $txtPass = New-Object System.Windows.Forms.TextBox
    $txtPass.Font = New-Object System.Drawing.Font('Segoe UI',12)
    $txtPass.Size = New-Object System.Drawing.Size(240,25)
    $txtPass.Location = New-Object System.Drawing.Point(110,113)
    $txtPass.UseSystemPasswordChar = $true
    $form.Controls.Add($txtPass)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Aceptar"
    $btnOk.Font = New-Object System.Drawing.Font('Segoe UI',12,[System.Drawing.FontStyle]::Bold)
    $btnOk.Size = New-Object System.Drawing.Size(100,35)
    $btnOk.Location = New-Object System.Drawing.Point(150,190)
    $btnOk.FlatStyle = 'Flat'
    $btnOk.BackColor = [System.Drawing.Color]::FromArgb(0x3B,0x8B,0xFF)
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
    $form.Controls.Add($btnOk)
    $form.AcceptButton = $btnOk

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        [System.Windows.Forms.MessageBox]::Show('Se canceló la autenticación. El programa finalizará.','Información',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        Exit 1
    }

    $user = $txtUser.Text.Trim()
    $pass = $txtPass.Text
    if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) {
        [System.Windows.Forms.MessageBox]::Show('Se canceló la autenticación. El programa finalizará.','Información',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        Exit 1
    }
    return @{ User = $user; Password = $pass }
}

# ========================================
# DIÁLOGO DE ACTUALIZACIÓN
# ========================================
function Show-UpdateDialog {
    param([string]$Version)
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Actualizando NRC_APP"
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.ControlBox = $false
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(0x2E,0x3B,0x4E)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Actualizando a versión $Version..."
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI',12,[System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = [System.Drawing.Color]::White
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point(20,40)
    $form.Controls.Add($lbl)

    $form.Show()
    return $form
}
function Save-EncryptedPassword {
    param(
        [Parameter(Mandatory)] [string]$Password,
        [Parameter(Mandatory)] [string]$DestinationFolder
    )
    
    # Validar que DestinationFolder no esté vacío
    if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
        Write-Warning "No se pudo guardar la contraseña: la carpeta de destino está vacía"
        return
    }
    
    if (-not (Test-Path $DestinationFolder)) { New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null }
    $destFile   = Join-Path $DestinationFolder 'temp.pass'
    $secure    = ConvertTo-SecureString $Password -AsPlainText -Force
    $encrypted = ConvertFrom-SecureString $secure -Key (Get-PassKey)
    Set-Content -Path $destFile -Value $encrypted -Encoding ASCII
}

# ========================================
# UTILIDADES DE VERSIÓN Y DESCARGA (igual)
# ========================================
function Get-NRCExe {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return $null }

    Get-ChildItem $Path -Filter $EXE_PATTERN |
        Sort-Object { [Version](Get-VersionString $_) } -Descending |
        Select-Object -First 1
}

function Test-ServerAccessible { param([string]$ServerPath) try { return Test-Path $ServerPath -ErrorAction Stop } catch { return $false } }
function FastCopy-FromServer {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        # En instalación limpia se copian TODOS los archivos (incluidos SQLite de datos).
        # En actualización se preservan los SQLite locales (datos de usuario).
        [switch]$FirstInstall
    )
    if (-not (Test-Path $DestinationPath)) {
        New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
    }
    # Robocopy flags:
    #   /MIR       = Mirror (copia nuevos/modificados, elimina obsoletos del destino)
    #   /Z         = Modo reiniciable (retoma archivos parcialmente copiados si se corta)
    #   /FFT       = Tolerancia 2s en timestamps (recomendado para shares de red)
    #   /DCOPY:DAT = Copia timestamps de directorios
    #   /R:3 /W:10 = 3 reintentos, espera 10s entre ellos (archivos transientemente en uso)
    #   /XD logs   = Nunca sobrescribir/borrar logs locales con contenido del servidor
    #   /XF        = Excluir ficheros locales (credenciales, logs)
    #   /NJH /NJS /NDL = Suprimir cabeceras de robocopy en la salida
    #
    # NOTA: database/ y libs/ ya NO se excluyen con /XD para que siempre se sincronicen
    # los JSON de configuración y las DLLs desde el servidor.
    # En actualizaciones se excluyen *.sqlite* para proteger datos de usuario locales.
    if ($FirstInstall) {
        # Instalación limpia: copiar absolutamente todo (no hay datos locales que proteger)
        $robArgs = @(
            $SourcePath, $DestinationPath,
            '/MIR', '/Z', '/FFT', '/DCOPY:DAT',
            '/R:3', '/W:10',
            '/XD', 'logs',
            '/XF', 'temp.pass', '*.log',
            '/NJH', '/NJS', '/NDL'
        )
    } else {
        # Actualización: sincronizar todo EXCEPTO datos de usuario en SQLite
        $robArgs = @(
            $SourcePath, $DestinationPath,
            '/MIR', '/Z', '/FFT', '/DCOPY:DAT',
            '/R:3', '/W:10',
            '/XD', 'logs',
            '/XF', 'temp.pass', '*.log', '*.sqlite*',
            '/NJH', '/NJS', '/NDL'
        )
    }
    & robocopy.exe @robArgs | Out-Null
    # Robocopy: 0-7 = exito (distintos niveles), 8+ = error real
    if ($LASTEXITCODE -ge 8) {
        throw "Robocopy error (codigo $LASTEXITCODE) - algunos archivos no pudieron copiarse."
    }
    return $true
}

function Get-VersionString {
    param([System.IO.FileInfo]$File)

    if ($File.Name -match '^launcher_v(\d+(?:\.\d+){0,2})\.exe$') {
        $ver = $matches[1]
        # Si solo tiene dos componentes (por ejemplo 1.0), añade .0
        if ($ver -match '^\d+\.\d+$') {
            $ver += '.0'
        }
        return $ver
    }
    return '0.0.0'
}

function Is-Newer {
    param(
        [string]$vLocal,
        [string]$vRemote
    )
    try   { return ([Version]$vRemote) -gt ([Version]$vLocal) }
    catch { return $vRemote -gt $vLocal }
}

function Remove-LocalInstallation {
    param([string]$LocalPath)

    try {
        if (Test-Path $LocalPath) {
            # Cerrar procesos relacionados antes de eliminar
            Write-Output "Cerrando procesos relacionados con la aplicación..."
            try {
                # Buscar y matar procesos que tengan módulos cargados desde la carpeta local
                $processes = Get-Process | Where-Object {
                    $_.Modules | Where-Object { $_.FileName -like "$LocalPath\*" }
                }
                foreach ($proc in $processes) {
                    try {
                        Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                        Write-Output "Cerrado proceso $($proc.Name) (PID: $($proc.Id))"
                    } catch {
                        Write-Warning "No se pudo cerrar proceso $($proc.Name) (PID: $($proc.Id)): $_"
                    }
                }
                Write-Output "Procesos relacionados cerrados."
            } catch {
                Write-Warning "Error al cerrar procesos: $_"
            }
            
            # Esperar un poco para que los procesos se cierren completamente
            Start-Sleep -Seconds 2
            
            # Intentar eliminar el contenido de la carpeta (no la carpeta raíz para evitar bloqueos)
            Remove-Item -Path "$LocalPath\*" -Recurse -Force -ErrorAction Stop
            Write-Output "Local installation content cleaned."
        }
    }
    catch {
        # Si falla, intentar de nuevo después de otro retraso
        Write-Warning "Primera eliminación falló: $($_.Exception.Message). Reintentando..."
        Start-Sleep -Seconds 3
        try {
            Remove-Item -Path "$LocalPath\*" -Recurse -Force -ErrorAction Stop
            Write-Output "Local installation content cleaned en segundo intento."
        } catch {
            throw "Failed to remove local installation content after retry: $($_.Exception.Message)"
        }
    }
}

function Ensure-FullControl {
    param(
        [string] $Path,
        [string] $User
    )
    
    # Validar que Path no esté vacío
    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-Warning "No se pudo ajustar permisos: la ruta está vacía"
        return
    }
    
    try {
        $acl = Get-Acl -Path $Path
        $ruleExists = $acl.Access | Where-Object {
            $_.IdentityReference -eq $User -and
            ($_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::FullControl)
        }
        if (-not $ruleExists) {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $User,
                [System.Security.AccessControl.FileSystemRights]::FullControl,
                [System.Security.AccessControl.InheritanceFlags]::None,
                [System.Security.AccessControl.PropagationFlags]::None,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
            $acl.AddAccessRule($rule)
            Set-Acl -Path $Path -AclObject $acl
            Write-Output "→ Permisos otorgados: $User ahora tiene Control total sobre $Path"
        }
    } catch {
        Write-Warning "No se pudo ajustar permisos en $Path $_"
    }
}
# ========================================
# LANZAR launcher_v*.exe
# ========================================
function Launch-App {
    param(
        [System.IO.FileInfo]$ScriptFile,
        [System.Management.Automation.PSCredential]$Credential
    )
    
    # Validar que ScriptFile no sea nulo
    if ($null -eq $ScriptFile) {
        Write-Error "No se puede lanzar la aplicación: el archivo ejecutable es nulo"
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($ScriptFile.FullName)) {
        Write-Error "No se puede lanzar la aplicación: la ruta del ejecutable está vacía"
        return
    }
    
    Write-Output "Lanzando $($ScriptFile.FullName)..."
    
    try {
        # Añadir WorkingDirectory y usar Wait para ver qué pasa
        $proc = Start-Process -FilePath $ScriptFile.FullName `
                             -Credential $Credential `
                             -WorkingDirectory $ScriptFile.DirectoryName `
                             -PassThru
        
        Write-Output "Proceso lanzado con PID: $($proc.Id)"
        
        # Esperar un momento para ver si el proceso sigue vivo
        Start-Sleep -Seconds 2
        if ($proc.HasExited) {
            Write-Output "El proceso terminó con código: $($proc.ExitCode)"
        }
    } catch {
        Write-Output "Error al lanzar: $_"
    }
}

# ========================================
# LÓGICA PRINCIPAL
# ========================================
function Main {
    Write-Output '=== NRC_APP Launcher ==='
    $creds = Show-CredentialsDialog
    $securePass = ConvertTo-SecureString $creds.Password -AsPlainText -Force
    $psCred     = New-Object System.Management.Automation.PSCredential($creds.User,$securePass)

    # Verificar instalación local
    $localExe = Get-NRCExe -Path $LOCAL_PATH
    $hasLocal = $null -ne $localExe

    if (-not $hasLocal) {
        Write-Output "No local installation found."

        if (-not (Test-ServerAccessible -ServerPath $SERVER_PATH)) {
            Write-Error "NRC_APP no está instalado y el servidor no responde."
            exit 1
        }

        Write-Output "Server accessible. Copying application..."
        FastCopy-FromServer -SourcePath $SERVER_PATH -DestinationPath $LOCAL_PATH -FirstInstall

        $localExe = Get-NRCExe -Path $LOCAL_PATH
        if (-not $localExe) {
            Write-Error "Failed to find executable after initial copy."
            exit 1
        }

        Save-EncryptedPassword -Password $creds.Password -DestinationFolder $LOCAL_PATH
        Ensure-FullControl -Path $localExe.FullName -User $creds.User
        Launch-App -ScriptFile $localExe -Credential $psCred
        exit 0
    }

    # Si hay instalación local
    Write-Output "Local installation found."
    if (-not (Test-ServerAccessible -ServerPath $SERVER_PATH)) {
        Write-Warning "Servidor no disponible; se ejecuta la versión local."
        Save-EncryptedPassword -Password $creds.Password -DestinationFolder $LOCAL_PATH
        Launch-App -ScriptFile $localExe -Credential $psCred
        exit 0
    }

    # Comparar versiones
    Write-Output "Servidor disponible. Comparando versiones..."
    $serverExe = Get-NRCExe -Path $SERVER_PATH

    if (-not $serverExe) {
        Write-Warning "No se encontró ejecutable en servidor. Usando versión local."
        Save-EncryptedPassword -Password $creds.Password -DestinationFolder $LOCAL_PATH
        Launch-App -ScriptFile $localExe -Credential $psCred
        exit 0
    }

    $vLocal  = Get-VersionString $localExe
    $vRemote = Get-VersionString $serverExe

    if (-not (Is-Newer $vLocal $vRemote)) {
        Write-Output "La versión local ($vLocal) es la más reciente."
        Save-EncryptedPassword -Password $creds.Password -DestinationFolder $LOCAL_PATH
        Launch-App -ScriptFile $localExe -Credential $psCred
        exit 0
    }

    Write-Output "Actualizando de $vLocal a $vRemote..."
    $updateForm = Show-UpdateDialog -Version $vRemote
    try {
        # Robocopy /MIR sincroniza directamente sin necesidad de borrar antes.
        # Los archivos en uso son manejados con /Z (modo reiniciable) y /R:3 /W:10.
        FastCopy-FromServer -SourcePath $SERVER_PATH -DestinationPath $LOCAL_PATH
    } finally {
        $updateForm.Close()
        $updateForm.Dispose()
    }

    $updatedExe = Get-NRCExe -Path $LOCAL_PATH
    if (-not $updatedExe) {
        Write-Error "No se encontró ejecutable tras la actualización."
        exit 1
    }

    Save-EncryptedPassword -Password $creds.Password -DestinationFolder $LOCAL_PATH
    Write-Output "Actualización completada. Lanzando aplicación..."
    Launch-App -ScriptFile $updatedExe -Credential $psCred
    exit 0
}


Main

