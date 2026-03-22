#==================================================================
# Script: LastLogin.ps1
# Descripción: Muestra el último inicio de sesión de usuarios en un equipo remoto
# Uso: .\LastLogin.ps1 -ComputerName NOMBRE_EQUIPO
#==================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$ComputerName
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#==================================================================
# Funcion: obtener el ultimo inicio de sesion detectado a partir de los perfiles locales.
#==================================================================
function Get-UltimoLogonExitoso {
    param(
        [string]$Usuario,
        [string]$Equipo
    )

    try {
        # Event IDs para logon exitoso: 4624 - An account was successfully logged on
        # Filtrar por logon types: 2 (Interactive), 10 (RemoteInteractive), 11 (CachedInteractive)
        # Ampliar búsqueda a más eventos (últimos 2 meses aproximadamente)
        
        $filterXml = @"
<QueryList>
  <Query Id="0" Path="Security">
    <Select Path="Security">
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and (EventID=4624)]]
      and
      *[EventData[Data[@Name='TargetUserName'] and (Data='$Usuario')]]
      and
      *[EventData[Data[@Name='LogonType'] and (Data='2' or Data='10' or Data='11')]]
    </Select>
  </Query>
</QueryList>
"@

        # Obtener el evento más reciente - Aumentar MaxEvents para buscar más atrás en el tiempo
        $eventos = Get-WinEvent -ComputerName $Equipo -FilterXml $filterXml -MaxEvents 5000 -ErrorAction Stop
        
        if ($eventos) {
            return $eventos[0]
        } else {
            return $null
        }
        
    } catch {
        return $null
    }
}

#==================================================================
# Función: Obtener lista de perfiles (SIMPLE - solo nombres)
#==================================================================
function Get-PerfilesLocales {
    param ([string]$Equipo)

    # Lista de exclusión extendida
    $excluir = @(
        'Default', 'Default User', 'Public', 'All Users',
        'systemprofile', 'LocalService', 'NetworkService',
        'defaultuser0', 'WDAGUtilityAccount', 'Administrador',
        'Guest', 'Invitado', 'DefaultAccount'
    )
    
    # Patrones a excluir (SQL Server, servicios, etc.)
    $patronesExcluir = @(
        '^MSSQL.*',
        '^SQL.*',
        '.*SQLEXPRESS.*',
        '.*SQLTELEMETRY.*',
        '^svc_.*',
        '^AppPool.*',
        '.*\.NET.*',
        '^ASPNET.*',
        '^IIS.*',
        '^IUSR.*',
        '^IWAM.*'
    )

    $Perfiles = @()

    try {
        $RutaPerfiles = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        $RegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Equipo)
        $SubKey = $RegKey.OpenSubKey($RutaPerfiles)

        foreach ($SID in $SubKey.GetSubKeyNames()) {
            # Filtrar SIDs de sistema (que terminan en .bak o son muy cortos)
            if ($SID -notmatch '\.bak$' -and $SID.Length -gt 20) {
                $SubClave = $SubKey.OpenSubKey($SID)
                $rutaPerfil = $SubClave.GetValue("ProfileImagePath")
                
                if ([string]::IsNullOrWhiteSpace($rutaPerfil)) { continue }
                
                $Usuario = [System.IO.Path]::GetFileName($rutaPerfil)
                
                # Filtrar usuarios del sistema
                if ($excluir -contains $Usuario) { continue }
                
                # Filtrar por patrones (SQL, servicios, etc.)
                $esExcluido = $false
                foreach ($patron in $patronesExcluir) {
                    if ($Usuario -match $patron) {
                        $esExcluido = $true
                        break
                    }
                }
                if ($esExcluido) { continue }
                
                $Perfiles += [PSCustomObject]@{
                    Usuario = $Usuario
                    SID = $SID
                }
            }
        }
        $RegKey.Close()

        return $Perfiles
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al obtener los perfiles: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $null
    }
}

#==================================================================
# MAIN - Ejecución Principal
#==================================================================

# Paso 1: Obtener lista de perfiles (RÁPIDO - sin consultar eventos)
$Perfiles = Get-PerfilesLocales -Equipo $ComputerName

if (-not $Perfiles -or $Perfiles.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No se encontraron perfiles en el equipo $ComputerName.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Paso 2: Mostrar formulario con lista de usuarios
$form = New-Object System.Windows.Forms.Form
$form.Text = "Seleccionar Usuario - $ComputerName"
$form.Width = 450
$form.Height = 400
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = "Selecciona un usuario para consultar su último inicio de sesión:"
$label.Dock = "Top"
$label.Height = 30
$label.TextAlign = "MiddleCenter"

# ListBox simple para mostrar usuarios
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Dock = "Fill"
$listBox.Height = 250

# Añadir usuarios al ListBox
foreach ($perfil in $Perfiles) {
    [void]$listBox.Items.Add($perfil.Usuario)
}

$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = "Bottom"
$buttonPanel.Height = 50

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Consultar Último Login"
$okButton.Dock = "Right"
$okButton.Width = 150
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$buttonPanel.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cerrar"
$cancelButton.Dock = "Right"
$cancelButton.Width = 100
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$buttonPanel.Controls.Add($cancelButton)

$form.Controls.Add($listBox)
$form.Controls.Add($label)
$form.Controls.Add($buttonPanel)
$form.AcceptButton = $okButton

# Paso 3: Cuando se selecciona un usuario, ENTONCES consultar eventos
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $listBox.SelectedItem) {
    $usuarioSeleccionado = $listBox.SelectedItem.ToString()
    
    # Mostrar mensaje de espera mientras se consultan eventos
    $formEspera = New-Object System.Windows.Forms.Form
    $formEspera.Text = "Consultando..."
    $formEspera.Width = 300
    $formEspera.Height = 100
    $formEspera.StartPosition = "CenterScreen"
    $formEspera.FormBorderStyle = "FixedDialog"
    $formEspera.ControlBox = $false

    $labelEspera = New-Object System.Windows.Forms.Label
    $labelEspera.Text = "Consultando visor de eventos...`nPor favor, espere."
    $labelEspera.Dock = "Fill"
    $labelEspera.TextAlign = "MiddleCenter"
    $formEspera.Controls.Add($labelEspera)

    $formEspera.Show()
    [System.Windows.Forms.Application]::DoEvents()

    # Obtener información detallada del último login
    $evento = Get-UltimoLogonExitoso -Usuario $usuarioSeleccionado -Equipo $ComputerName

    $formEspera.Close()
    $formEspera.Dispose()

    if ($evento) {
        # Determinar el tipo de logon
        $logonType = $evento.Properties[8].Value
        $tipoLogonTexto = switch ($logonType) {
            2 { "Interactivo (consola física)" }
            10 { "Remoto Interactivo (RDP)" }
            11 { "Interactivo en caché" }
            default { "Tipo $logonType" }
        }
        
        # Calcular hace cuánto tiempo fue el último login
        $tiempoTranscurrido = (Get-Date) - $evento.TimeCreated
        $diasTranscurridos = [math]::Floor($tiempoTranscurrido.TotalDays)
        
        $tiempoTexto = if ($diasTranscurridos -eq 0) {
            "Hoy"
        } elseif ($diasTranscurridos -eq 1) {
            "Ayer"
        } elseif ($diasTranscurridos -lt 7) {
            "Hace $diasTranscurridos días"
        } elseif ($diasTranscurridos -lt 30) {
            $semanas = [math]::Floor($diasTranscurridos / 7)
            "Hace $semanas semana(s)"
        } elseif ($diasTranscurridos -lt 365) {
            $meses = [math]::Floor($diasTranscurridos / 30)
            "Hace $meses mes(es)"
        } else {
            $anos = [math]::Floor($diasTranscurridos / 365)
            "Hace $anos año(s)"
        }

        $mensaje = @"
╔═══════════════════════════════════════════════════════════╗
║          INFORMACIÓN DE ÚLTIMO INICIO DE SESIÓN           ║
╚═══════════════════════════════════════════════════════════╝

Usuario: $usuarioSeleccionado
Equipo: $ComputerName

📅 Fecha y Hora: $($evento.TimeCreated.ToString("dddd, dd/MM/yyyy HH:mm:ss"))
⏰ Última conexión: $tiempoTexto ($diasTranscurridos días)
🔑 Tipo de Inicio: $tipoLogonTexto
🆔 ID de Evento: 4624

─────────────────────────────────────────────────────────────
Detalles adicionales:
─────────────────────────────────────────────────────────────
• Dominio: $($evento.Properties[6].Value)
• Proceso de inicio: $($evento.Properties[9].Value)
• Estación de trabajo: $($evento.Properties[11].Value)
• ID de inicio de sesión: $($evento.Properties[7].Value)
"@
        # Crear formulario de resultado en lugar de MessageBox
        $formResultado = New-Object System.Windows.Forms.Form
        $formResultado.Text = "Último Login - $usuarioSeleccionado"
        $formResultado.Width = 650
        $formResultado.Height = 450
        $formResultado.StartPosition = "CenterScreen"
        $formResultado.FormBorderStyle = "FixedDialog"
        $formResultado.MaximizeBox = $false
        
        $textBoxResultado = New-Object System.Windows.Forms.TextBox
        $textBoxResultado.Multiline = $true
        $textBoxResultado.ReadOnly = $true
        $textBoxResultado.Dock = "Fill"
        $textBoxResultado.Font = New-Object System.Drawing.Font("Consolas", 10)
        $textBoxResultado.Text = $mensaje
        $textBoxResultado.ScrollBars = "Vertical"
        
        $btnCerrar = New-Object System.Windows.Forms.Button
        $btnCerrar.Text = "Cerrar"
        $btnCerrar.Dock = "Bottom"
        $btnCerrar.Height = 40
        $btnCerrar.DialogResult = [System.Windows.Forms.DialogResult]::OK
        
        $formResultado.Controls.Add($textBoxResultado)
        $formResultado.Controls.Add($btnCerrar)
        $formResultado.AcceptButton = $btnCerrar
        
        [void]$formResultado.ShowDialog()
        $formResultado.Dispose()
    } else {
        $mensaje = @"
╔═══════════════════════════════════════════════════════════╗
║              SIN INFORMACIÓN DISPONIBLE                   ║
╚═══════════════════════════════════════════════════════════╝

Usuario: $usuarioSeleccionado
Equipo: $ComputerName

⚠️ No se encontraron eventos de inicio de sesión interactivo 
   recientes en el visor de eventos de seguridad.

Esto puede deberse a:
• El usuario no ha iniciado sesión en los últimos meses
• Los eventos han sido rotados del log de seguridad
• Permisos insuficientes para acceder al log de seguridad
• El perfil existe pero nunca se ha utilizado

Nota: Se buscaron hasta 5000 eventos en el log de seguridad.
"@
        # Crear formulario de resultado en lugar de MessageBox
        $formResultado = New-Object System.Windows.Forms.Form
        $formResultado.Text = "Sin Información - $usuarioSeleccionado"
        $formResultado.Width = 650
        $formResultado.Height = 400
        $formResultado.StartPosition = "CenterScreen"
        $formResultado.FormBorderStyle = "FixedDialog"
        $formResultado.MaximizeBox = $false
        
        $textBoxResultado = New-Object System.Windows.Forms.TextBox
        $textBoxResultado.Multiline = $true
        $textBoxResultado.ReadOnly = $true
        $textBoxResultado.Dock = "Fill"
        $textBoxResultado.Font = New-Object System.Drawing.Font("Consolas", 10)
        $textBoxResultado.Text = $mensaje
        $textBoxResultado.ScrollBars = "Vertical"
        
        $btnCerrar = New-Object System.Windows.Forms.Button
        $btnCerrar.Text = "Cerrar"
        $btnCerrar.Dock = "Bottom"
        $btnCerrar.Height = 40
        $btnCerrar.DialogResult = [System.Windows.Forms.DialogResult]::OK
        
        $formResultado.Controls.Add($textBoxResultado)
        $formResultado.Controls.Add($btnCerrar)
        $formResultado.AcceptButton = $btnCerrar
        
        [void]$formResultado.ShowDialog()
        $formResultado.Dispose()
    }
}

$form.Dispose()
