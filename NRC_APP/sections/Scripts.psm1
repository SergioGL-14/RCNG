# Sección encargada del menú Scripts.
# Aquí se registran los built-in, se leen los scripts custom y se decide
# cómo lanzar cada entrada según su método de ejecución.
param()

#==================================================================
# BLOQUE: Helpers de UI y rutas
#==================================================================

# Recupera el equipo actual desde el cuadro principal.
function Get-RemoteComputer {
    $cn = $global:textbox_computername.Text.Trim()
    return $cn
}

# Cuadro de confirmación estándar para acciones destructivas o sensibles.
function Confirm-Action {
    param([string]$Message, [string]$Title = "Confirmación")
    $r = [System.Windows.Forms.MessageBox]::Show(
        $Message, $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

# Cuadro de error reutilizable.
function Show-Error {
    param([string]$Message, [string]$Title = "Error")
    [System.Windows.Forms.MessageBox]::Show(
        $Message, $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

# Aviso específico cuando se intenta lanzar algo sin equipo indicado.
function Show-NoComputer {
    [System.Windows.Forms.MessageBox]::Show(
        "Por favor, introduzca un nombre de equipo válido.",
        "Equipo no especificado",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
}

# Crea un encabezado visual de categoría dentro del menú desplegable.
function New-MenuHeader {
    param([string]$Text)
    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text    = $Text
    $item.Enabled = $false
    $item.Font    = New-Object System.Drawing.Font(
        "Segoe UI", 8.5,
        [System.Drawing.FontStyle]::Bold
    )
    return $item
}

# Crea un ToolStripMenuItem normal con icono opcional y handler de clic
function New-MenuItem {
    param(
        [string]$Text,
        [string]$IconName,
        [scriptblock]$OnClick
    )
    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text = $Text

    if (-not [string]::IsNullOrWhiteSpace($IconName)) {
        $iconPath = Join-Path $Global:ScriptRoot "icos\$IconName"
        if (Test-Path $iconPath) {
            try { $item.Image = [System.Drawing.Image]::FromFile($iconPath) } catch {}
        }
    }

    if ($null -ne $OnClick) {
        $item.Add_Click($OnClick)
    }
    return $item
}

#==================================================================
# BLOQUE: Métodos de ejecución disponibles
#==================================================================

function Get-ExecutionMethods {
    return @(
        [PSCustomObject]@{
            Id               = 'standard'
            Name             = 'Estándar (PS1 con -ComputerName)'
            Description      = 'Ejecuta un script PowerShell localmente, pasándole el nombre del equipo remoto como parámetro -ComputerName. El script debe gestionar internamente la conexión remota (CIM, WMI, Invoke-Command, etc.).'
            FileTypes        = @('.ps1')
            RequiresComputer = $true
        }
        [PSCustomObject]@{
            Id               = 'psexec-system'
            Name             = 'PsExec Remoto (SYSTEM)'
            Description      = 'Copia el script .ps1 al equipo remoto (C:\Windows\Temp) y lo ejecuta como SYSTEM vía PsExec. Ideal para scripts que necesitan privilegios de administrador local y acceso directo a recursos del equipo remoto. El script NO debe tener parámetro -ComputerName; usará $env:COMPUTERNAME.'
            FileTypes        = @('.ps1')
            RequiresComputer = $true
        }
        [PSCustomObject]@{
            Id               = 'batch-remote'
            Name             = 'Batch/CMD Remoto'
            Description      = 'Ejecuta un script .bat/.cmd pasando el nombre del equipo como primer argumento (%1). La ventana de comandos permanece abierta para ver el resultado.'
            FileTypes        = @('.bat', '.cmd')
            RequiresComputer = $true
        }
        [PSCustomObject]@{
            Id               = 'local'
            Name             = 'Ejecución Local'
            Description      = 'Ejecuta el script o aplicación en el equipo del técnico, sin pasar nombre de equipo. Útil para herramientas y utilidades locales.'
            FileTypes        = @('.ps1', '.bat', '.cmd', '.exe', '.vbs')
            RequiresComputer = $false
        }
    )
}

function Show-ExecutionMethodInfoDialog {
    $methods = Get-ExecutionMethods

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = "Métodos de Ejecución - Información"
    $form.Size            = New-Object System.Drawing.Size(620, 560)
    $form.StartPosition   = "CenterScreen"
    $form.TopMost         = $true
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false

    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.Location  = New-Object System.Drawing.Point(12, 12)
    $rtb.Size      = New-Object System.Drawing.Size(580, 460)
    $rtb.ReadOnly  = $true
    $rtb.BackColor = [System.Drawing.Color]::White
    $rtb.Font      = New-Object System.Drawing.Font("Segoe UI", 9.5)

    $text  = "MÉTODOS DE EJECUCIÓN DISPONIBLES`r`n"
    $text += "=" * 50 + "`r`n`r`n"

    foreach ($m in $methods) {
        $text += "▶ $($m.Name)`r`n"
        $text += "  ID interno: $($m.Id)`r`n"
        $text += "  Tipos de archivo: $($m.FileTypes -join ', ')`r`n"
        $text += "  Requiere equipo remoto: $(if ($m.RequiresComputer) {'Sí'} else {'No'})`r`n"
        $text += "`r`n  $($m.Description)`r`n"
        $text += "`r`n" + "-" * 50 + "`r`n`r`n"
    }

    $text += "CÓMO DISEÑAR LOS SCRIPTS PARA CADA MÉTODO`r`n"
    $text += "=" * 50 + "`r`n`r`n"

    $text += "Estándar (standard):`r`n"
    $text += "  - El script recibe -ComputerName como parámetro.`r`n"
    $text += "  - Internamente usa CIM, WMI, Invoke-Command o`r`n"
    $text += "    registro remoto para operar sobre el equipo.`r`n"
    $text += "  - Se ejecuta en la máquina del técnico.`r`n"
    $text += "  - Ejemplo: param([string]`$ComputerName)`r`n"
    $text += "    Get-CimInstance Win32_OS -ComputerName `$ComputerName`r`n`r`n"

    $text += "PsExec Remoto (psexec-system):`r`n"
    $text += "  - El script se copia a \\equipo\C`$\Windows\Temp\`r`n"
    $text += "  - Se ejecuta en el equipo remoto como SYSTEM.`r`n"
    $text += "  - NO debe tener -ComputerName. Usa `$env:COMPUTERNAME.`r`n"
    $text += "  - Puede incluir #Requires -RunAsAdministrator.`r`n"
    $text += "  - Ideal para reparaciones locales, AppX, registro HKU.`r`n`r`n"

    $text += "Batch Remoto (batch-remote):`r`n"
    $text += "  - El script recibe el nombre de equipo como %1.`r`n"
    $text += "  - Se ejecuta en cmd.exe /k (ventana abierta).`r`n"
    $text += "  - Ejemplo: net stop spooler /y && ...`r`n`r`n"

    $text += "Ejecución Local (local):`r`n"
    $text += "  - Se ejecuta localmente sin argumentos remotos.`r`n"
    $text += "  - Para herramientas del técnico de soporte.`r`n"

    $rtb.Text = $text
    $form.Controls.Add($rtb)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text         = "Cerrar"
    $btnClose.Location     = New-Object System.Drawing.Point(510, 485)
    $btnClose.Size         = New-Object System.Drawing.Size(80, 28)
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnClose)

    $form.ShowDialog() | Out-Null
}

#==================================================================
# BLOQUE: Base de datos de scripts (JSON) - Esquema unificado
#==================================================================

function Get-ScriptsDatabase {
    # Leer desde el servidor si está disponible (vía SharedDataManager), fallback local
    if (Get-Command 'Get-ScriptsDbContent' -ErrorAction SilentlyContinue) {
        $raw = Get-ScriptsDbContent
    } else {
        $dbPath = Join-Path $Global:ScriptRoot 'database\scripts_db.json'
        $raw = if (Test-Path $dbPath) { Get-Content $dbPath -Raw -Encoding UTF8 } else { '[]' }
    }

    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    try {
        $data = $raw | ConvertFrom-Json
        if ($data -isnot [array]) { $data = @($data) }
        # Normalizar: asegurar que todas las entradas tienen los campos requeridos
        foreach ($entry in $data) {
            if (-not $entry.PSObject.Properties['ExecutionMethod']) {
                $entry | Add-Member -NotePropertyName 'ExecutionMethod' -NotePropertyValue 'standard' -Force
            }
            if (-not $entry.PSObject.Properties['BuiltIn']) {
                $entry | Add-Member -NotePropertyName 'BuiltIn' -NotePropertyValue $false -Force
            }
            if (-not $entry.PSObject.Properties['BuiltInKey']) {
                $entry | Add-Member -NotePropertyName 'BuiltInKey' -NotePropertyValue '' -Force
            }
            if (-not $entry.PSObject.Properties['CustomHandler']) {
                $entry | Add-Member -NotePropertyName 'CustomHandler' -NotePropertyValue '' -Force
            }
            if (-not $entry.PSObject.Properties['SortOrder']) {
                $entry | Add-Member -NotePropertyName 'SortOrder' -NotePropertyValue 100 -Force
            }
            if (-not $entry.PSObject.Properties['Category']) {
                $entry | Add-Member -NotePropertyName 'Category' -NotePropertyValue 'Custom' -Force
            }
        }
        return $data
    } catch { return @() }
}

function Save-ScriptsDatabase {
    param([Parameter(Mandatory=$true)] $Database)
    $json = if ($Database.Count -eq 0) { '[]' } else { $Database | ConvertTo-Json -Depth 3 }

    # Guardar SOLO en local. Para replicar al servidor usar Sync-ScriptsJsonToServer.
    if (Get-Command 'Save-ScriptsDbLocal' -ErrorAction SilentlyContinue) {
        Save-ScriptsDbLocal -JsonContent $json
    } else {
        $dbPath = Join-Path $Global:ScriptRoot 'database\scripts_db.json'
        $json | Out-File -FilePath $dbPath -Encoding UTF8 -Force
    }
}

function Get-BuiltInScriptDefaults {
    return @(
        @{ BuiltInKey='proxy';      Name='Configurar Proxy';        FileName='';                             IconFile='proxy_config.ico';           Category='Configuracion'; ExecutionMethod='custom';       CustomHandler='Invoke-Scripts_ConfigurarProxy';   SortOrder=1  }
        @{ BuiltInKey='edge';       Name='Configurar Edge';         FileName='Configurar Edge.ps1';          IconFile='Edge.ico';                   Category='Configuracion'; ExecutionMethod='standard';     CustomHandler='';                                 SortOrder=2  }
        @{ BuiltInKey='lastlogin';  Name='Last Login';              FileName='LastLogin.ps1';                IconFile='LastLogin.ico';              Category='Utilidades';    ExecutionMethod='standard';     CustomHandler='';                                 SortOrder=10 }
        @{ BuiltInKey='kb';         Name='Desinstalar KB';          FileName='Desinstalar KB.ps1';           IconFile='DesinstalarKB.ico';          Category='Utilidades';    ExecutionMethod='custom';       CustomHandler='Invoke-Scripts_DesinstalarKB';     SortOrder=12 }
        @{ BuiltInKey='spooler';    Name='Clean Spooler';           FileName='Clean Spooler.bat';            IconFile='Clean_Spooler.ico';          Category='Workaround';    ExecutionMethod='batch-remote'; CustomHandler='';                                 SortOrder=20 }
        @{ BuiltInKey='temporales'; Name='Clean Temp';              FileName='Clean Temp.bat';               IconFile='Clean_Temporales.ico';       Category='Workaround';    ExecutionMethod='batch-remote'; CustomHandler='';                                 SortOrder=21 }
        @{ BuiltInKey='tarjeta';    Name='Reset Scardvr';           FileName='Reset Scardvr.bat';            IconFile='Reset_Card.ico';             Category='Workaround';    ExecutionMethod='batch-remote'; CustomHandler='';                                 SortOrder=22 }
        @{ BuiltInKey='lector';     Name='Reconectar Lector';       FileName='Reconectar Lector.ps1';        IconFile='Reconectar_Lector.ico';      Category='Workaround';    ExecutionMethod='standard';     CustomHandler='';                                 SortOrder=23 }
        @{ BuiltInKey='taskbar';    Name='Repair Taskbar';          FileName='Repair Taskbar.ps1';           IconFile='Taskbar.ico';                Category='Workaround';    ExecutionMethod='psexec-system'; CustomHandler='';                                SortOrder=26 }
        @{ BuiltInKey='perfil';     Name='Renombrar Perfil';        FileName='';                             IconFile='RenombrarPerfil.ico';        Category='Workaround';    ExecutionMethod='custom';       CustomHandler='Invoke-Scripts_RenombrarPerfil';   SortOrder=27 }
    )
}

function Register-BuiltInScripts {
    $db       = @(Get-ScriptsDatabase)
    $defaults = Get-BuiltInScriptDefaults
    $changed  = $false
    $defaultKeys = @($defaults | ForEach-Object { $_.BuiltInKey })

    foreach ($legacyEntry in @($db | Where-Object {
        $_.BuiltIn -and $_.BuiltInKey -and ($defaultKeys -notcontains $_.BuiltInKey)
    })) {
        $db = @($db | Where-Object { $_.BuiltInKey -ne $legacyEntry.BuiltInKey })
        $changed = $true
    }

    foreach ($def in $defaults) {
        $existing = $db | Where-Object { $_.BuiltInKey -eq $def.BuiltInKey }
        if (-not $existing) {
            $db += [PSCustomObject]@{
                Name            = $def.Name
                FileName        = $def.FileName
                IconFile        = $def.IconFile
                Category        = $def.Category
                ExecutionMethod = $def.ExecutionMethod
                CustomHandler   = $def.CustomHandler
                BuiltIn         = $true
                BuiltInKey      = $def.BuiltInKey
                SortOrder       = $def.SortOrder
                AddedOn         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            }
            $changed = $true
        }
    }

    if ($changed) { Save-ScriptsDatabase -Database $db }
    return @(Get-ScriptsDatabase)
}

#------------------------------------------------------------------
# FUNCIÓN COMPARTIDA: diálogo de configuración de script
#   - Nombre, Icono, Método de ejecución
#   - Botón ℹ con información de métodos
# Devuelve @{Name; IconFile; ExecutionMethod} o $null si se canceló.
#------------------------------------------------------------------
function Show-ScriptConfigDialog {
    param(
        [string]$InitialName   = '',
        [string]$InitialIcon   = '',
        [string]$FileName      = '',
        [string]$InitialMethod = 'standard',
        [switch]$IsCustomHandler
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = "Configurar Script en el Menú"
    $form.Size            = New-Object System.Drawing.Size(460, 310)
    $form.StartPosition   = "CenterScreen"
    $form.TopMost         = $true
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false

    $y = 16

    # -- Nombre --
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text     = "Nombre a mostrar en el menú:"
    $lblName.Location = New-Object System.Drawing.Point(12, $y)
    $lblName.Size     = New-Object System.Drawing.Size(420, 18)
    $form.Controls.Add($lblName)
    $y += 22

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Text     = $InitialName
    $txtName.Location = New-Object System.Drawing.Point(12, $y)
    $txtName.Size     = New-Object System.Drawing.Size(420, 24)
    $form.Controls.Add($txtName)
    $y += 36

    # -- Método de ejecución --
    $lblMethod = New-Object System.Windows.Forms.Label
    $lblMethod.Text     = "Método de ejecución:"
    $lblMethod.Location = New-Object System.Drawing.Point(12, $y)
    $lblMethod.Size     = New-Object System.Drawing.Size(300, 18)
    $form.Controls.Add($lblMethod)
    $y += 22

    $cmbMethod = New-Object System.Windows.Forms.ComboBox
    $cmbMethod.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbMethod.Location      = New-Object System.Drawing.Point(12, $y)
    $cmbMethod.Size           = New-Object System.Drawing.Size(370, 24)

    $methods = Get-ExecutionMethods
    foreach ($m in $methods) { [void]$cmbMethod.Items.Add($m.Name) }

    # Selección inicial
    if ($IsCustomHandler) {
        [void]$cmbMethod.Items.Add('Personalizado (handler interno)')
        $cmbMethod.SelectedIndex = $cmbMethod.Items.Count - 1
        $cmbMethod.Enabled       = $false
    } else {
        $selIdx = 0
        for ($i = 0; $i -lt $methods.Count; $i++) {
            if ($methods[$i].Id -eq $InitialMethod) { $selIdx = $i; break }
        }
        $cmbMethod.SelectedIndex = $selIdx
    }
    $form.Controls.Add($cmbMethod)

    # Botón ℹ
    $btnInfo = New-Object System.Windows.Forms.Button
    $btnInfo.Text      = [char]0x2139   # ℹ
    $btnInfo.Location  = New-Object System.Drawing.Point(390, $y)
    $btnInfo.Size      = New-Object System.Drawing.Size(42, $cmbMethod.Height)
    $btnInfo.FlatStyle = [System.Windows.Forms.FlatStyle]::System
    $btnInfo.Add_Click({ Show-ExecutionMethodInfoDialog })
    $form.Controls.Add($btnInfo)
    $y += 36

    # -- Icono --
    $lblIcon = New-Object System.Windows.Forms.Label
    $lblIcon.Text     = "Icono (opcional - se copiará a la carpeta icos):"
    $lblIcon.Location = New-Object System.Drawing.Point(12, $y)
    $lblIcon.Size     = New-Object System.Drawing.Size(420, 18)
    $form.Controls.Add($lblIcon)
    $y += 22

    $currentIconPath = ''
    if (-not [string]::IsNullOrWhiteSpace($InitialIcon)) {
        $p = Join-Path $Global:ScriptRoot "icos\$InitialIcon"
        if (Test-Path $p) { $currentIconPath = $p }
    }

    $txtIcon = New-Object System.Windows.Forms.TextBox
    $txtIcon.Text      = $currentIconPath
    $txtIcon.Location  = New-Object System.Drawing.Point(12, $y)
    $txtIcon.Size      = New-Object System.Drawing.Size(338, 24)
    $txtIcon.ReadOnly  = $true
    $txtIcon.BackColor = [System.Drawing.SystemColors]::Window
    $form.Controls.Add($txtIcon)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text     = "..."
    $btnBrowse.Location = New-Object System.Drawing.Point(358, $y)
    $btnBrowse.Size     = New-Object System.Drawing.Size(74, 24)
    $btnBrowse.Add_Click({
        $ifd = New-Object System.Windows.Forms.OpenFileDialog
        $ifd.Title  = "Seleccionar Icono"
        $ifd.Filter = "Imágenes de icono (*.ico;*.png;*.jpg)|*.ico;*.png;*.jpg"
        if ($ifd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtIcon.Text = $ifd.FileName
        }
    })
    $form.Controls.Add($btnBrowse)
    $y += 46

    # -- Botones OK / Cancelar --
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text         = "Aceptar"
    $btnOk.Location     = New-Object System.Drawing.Point(260, $y)
    $btnOk.Size         = New-Object System.Drawing.Size(80, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)
    $form.AcceptButton  = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Cancelar"
    $btnCancel.Location     = New-Object System.Drawing.Point(355, $y)
    $btnCancel.Size         = New-Object System.Drawing.Size(80, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    # --- Procesar resultado ---
    $name = $txtName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = if (-not [string]::IsNullOrWhiteSpace($FileName)) {
            [System.IO.Path]::GetFileNameWithoutExtension($FileName)
        } else { $InitialName }
    }

    # Método seleccionado
    $selectedMethod = $InitialMethod
    if (-not $IsCustomHandler) {
        $selectedMethod = $methods[$cmbMethod.SelectedIndex].Id
    }

    # Procesar icono: copiar si se seleccionó un archivo nuevo fuera de icos\
    $iconFileName = $InitialIcon
    if (-not [string]::IsNullOrWhiteSpace($txtIcon.Text) -and (Test-Path $txtIcon.Text)) {
        $iconSrc = $txtIcon.Text
        $icosDir = Join-Path $Global:ScriptRoot 'icos'
        if (-not $iconSrc.StartsWith($icosDir)) {
            $iconBase = if (-not [string]::IsNullOrWhiteSpace($FileName)) {
                [System.IO.Path]::GetFileNameWithoutExtension($FileName)
            } else {
                [System.IO.Path]::GetFileNameWithoutExtension($iconSrc)
            }
            $iconExt  = [System.IO.Path]::GetExtension($iconSrc)
            $iconDest = Join-Path $icosDir "$iconBase$iconExt"
            try {
                Copy-Item -Path $iconSrc -Destination $iconDest -Force -ErrorAction Stop
                $iconFileName = "$iconBase$iconExt"
                # Publicar también al servidor
                if (Get-Command 'Publish-IconToServer' -ErrorAction SilentlyContinue) {
                    Publish-IconToServer -LocalIconPath $iconDest | Out-Null
                }
            } catch { }
        }
    } elseif ([string]::IsNullOrWhiteSpace($txtIcon.Text)) {
        $iconFileName = ''
    }

    return @{ Name = $name; IconFile = $iconFileName; ExecutionMethod = $selectedMethod }
}

#==================================================================
# BLOQUE: Ejecución unificada de scripts
#==================================================================

function Invoke-ScriptByMethod {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$ScriptEntry
    )

    $method = $ScriptEntry.ExecutionMethod
    $cn     = Get-RemoteComputer

    # Custom handler: invocar la función específica
    if ($method -eq 'custom') {
        $handler = $ScriptEntry.CustomHandler
        if (-not [string]::IsNullOrWhiteSpace($handler)) {
            & $handler
        }
        return
    }

    # Todos los métodos no-custom requieren un FileName
    $fileName   = $ScriptEntry.FileName
    $scriptPath = Join-Path $Global:ScriptRoot "scripts\$fileName"

    if ([string]::IsNullOrWhiteSpace($fileName) -or -not (Test-Path $scriptPath)) {
        Show-Error "No se encuentra el script:`n$scriptPath" "Script no encontrado"
        return
    }

    # Verificar si requiere equipo remoto
    $methods   = Get-ExecutionMethods
    $methodDef = $methods | Where-Object { $_.Id -eq $method }
    if ($methodDef -and $methodDef.RequiresComputer) {
        if ([string]::IsNullOrWhiteSpace($cn)) {
            Show-NoComputer
            return
        }
    }

    $displayName = $ScriptEntry.Name

    switch ($method) {
        'standard' {
            Add-Logs -text "$cn - $displayName lanzado"
            Start-Process -FilePath "powershell.exe" `
                -ArgumentList "-NoExit -ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`" -ComputerName `"$cn`""
        }
        'psexec-system' {
            Add-Logs -text "$cn - $displayName lanzado (PsExec)"
            Invoke-NRCPsExecScript -ScriptPath $scriptPath -ComputerName $cn -DisplayName $displayName
        }
        'batch-remote' {
            Add-Logs -text "$cn - $displayName lanzado"
            $cmdArgs = "/k call `"$scriptPath`" `"$cn`""
            Start-Process -FilePath "cmd.exe" -ArgumentList $cmdArgs
        }
        'local' {
            Add-Logs -text "Localhost - $displayName lanzado"
            Start-Process $scriptPath
        }
        default {
            # Fallback genérico via ScriptRunner
            Invoke-NRCScript -ScriptPath $scriptPath -ComputerName $cn
        }
    }
}

#==================================================================
# BLOQUE: Construcción unificada de ítems de menú
#==================================================================

function New-ScriptMenuItem {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$ScriptEntry,

        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripMenuItem]$ParentMenu
    )

    $captEntry   = $ScriptEntry
    $displayName = $ScriptEntry.Name
    $iconName    = if ($ScriptEntry.IconFile) { $ScriptEntry.IconFile } else { '' }
    $isBuiltIn   = [bool]$ScriptEntry.BuiltIn

    # Click izquierdo: ejecutar el script según su método
    $item = New-MenuItem -Text $displayName -IconName $iconName -OnClick {
        Invoke-ScriptByMethod -ScriptEntry $captEntry
    }.GetNewClosure()

    # --- Menú contextual (clic derecho) ---
    $ctx = New-Object System.Windows.Forms.ContextMenuStrip

    # Modificar (disponible para todos)
    $miModify      = New-Object System.Windows.Forms.ToolStripMenuItem
    $miModify.Text = "Modificar"
    [void]$ctx.Items.Add($miModify)

    # Eliminar (solo scripts no built-in)
    $miDelete = $null
    if (-not $isBuiltIn) {
        $miDelete      = New-Object System.Windows.Forms.ToolStripMenuItem
        $miDelete.Text = "Eliminar"
        [void]$ctx.Items.Add($miDelete)
    }

    # Restaurar valores por defecto (solo built-in)
    $miRestore = $null
    if ($isBuiltIn) {
        $miRestore      = New-Object System.Windows.Forms.ToolStripMenuItem
        $miRestore.Text = "Restaurar valores por defecto"
        [void]$ctx.Items.Add($miRestore)
    }

    $captItem   = $item
    $captParent = $ParentMenu
    $captIcon   = $iconName

    # -- Handler: Modificar --
    $miModify.Add_Click({
        $isCustom = (-not [string]::IsNullOrWhiteSpace($captEntry.CustomHandler))
        $res = Show-ScriptConfigDialog `
            -InitialName   $captEntry.Name `
            -InitialIcon   $captIcon `
            -FileName      $captEntry.FileName `
            -InitialMethod $captEntry.ExecutionMethod `
            -IsCustomHandler:$isCustom
        if ($null -eq $res) { return }

        # Actualizar en BD
        $db = @(Get-ScriptsDatabase)
        foreach ($e in $db) {
            $isMatch = $false
            if ($captEntry.BuiltIn -and $captEntry.BuiltInKey) {
                $isMatch = ($e.BuiltInKey -eq $captEntry.BuiltInKey)
            } else {
                $isMatch = ($e.FileName -eq $captEntry.FileName -and -not $e.BuiltIn)
            }
            if ($isMatch) {
                $e.Name            = $res.Name
                $e.IconFile        = $res.IconFile
                $e.ExecutionMethod = $res.ExecutionMethod
                break
            }
        }
        Save-ScriptsDatabase -Database $db

        # Preguntar si replicar cambios en el servidor
        if ((Get-Command 'Test-ServerAvailable' -ErrorAction SilentlyContinue) -and (Test-ServerAvailable)) {
            $rRep = [System.Windows.Forms.MessageBox]::Show(
                "¿Desea replicar los cambios en el servidor?`n`nScript: $($res.Name)",
                "Replicar en Servidor",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($rRep -eq [System.Windows.Forms.DialogResult]::Yes) {
                Sync-ScriptsJsonToServer | Out-Null
                # Preguntar si también actualizar el archivo del script
                if (-not $captEntry.BuiltIn -and -not [string]::IsNullOrWhiteSpace($captEntry.FileName)) {
                    $rFile = [System.Windows.Forms.MessageBox]::Show(
                        "¿Actualizar también el archivo del script ($($captEntry.FileName)) en el servidor?",
                        "Actualizar Archivo en Servidor",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question
                    )
                    if ($rFile -eq [System.Windows.Forms.DialogResult]::Yes) {
                        $localFile = Join-Path $Global:ScriptRoot "scripts\$($captEntry.FileName)"
                        Publish-ScriptToServer -LocalScriptPath $localFile | Out-Null
                    }
                }
            }
        }

        # Actualizar UI
        $captItem.Text  = $res.Name
        $captItem.Image = $null
        $captIcon = $res.IconFile
        if (-not [string]::IsNullOrWhiteSpace($res.IconFile)) {
            $ip = Join-Path $Global:ScriptRoot "icos\$($res.IconFile)"
            if (Test-Path $ip) { try { $captItem.Image = [System.Drawing.Image]::FromFile($ip) } catch {} }
        }

        # Actualizar referencia capturada para futuras ejecuciones
        $captEntry.Name            = $res.Name
        $captEntry.IconFile        = $res.IconFile
        $captEntry.ExecutionMethod = $res.ExecutionMethod
    }.GetNewClosure())

    # -- Handler: Eliminar (solo custom scripts) --
    if ($miDelete) {
        $miDelete.Add_Click({
            $r = [System.Windows.Forms.MessageBox]::Show(
                "¿Eliminar el script '$($captItem.Text)' del menú?",
                "Confirmar Eliminación",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }

            $db = @(Get-ScriptsDatabase | Where-Object {
                -not ($_.FileName -eq $captEntry.FileName -and -not $_.BuiltIn)
            })
            Save-ScriptsDatabase -Database $db

            # Preguntar si eliminar también del servidor
            if ((Get-Command 'Test-ServerAvailable' -ErrorAction SilentlyContinue) -and (Test-ServerAvailable)) {
                $rSrv = [System.Windows.Forms.MessageBox]::Show(
                    "¿Desea eliminar también el script del servidor?`n`nArchivo: $($captEntry.FileName)",
                    "Eliminar del Servidor",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
                if ($rSrv -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Sync-ScriptsJsonToServer | Out-Null
                    if (-not [string]::IsNullOrWhiteSpace($captEntry.FileName)) {
                        Remove-ScriptFromServer -FileName $captEntry.FileName | Out-Null
                    }
                }
            }

            [void]$captParent.DropDownItems.Remove($captItem)
        }.GetNewClosure())
    }

    # -- Handler: Restaurar valores por defecto (solo built-in) --
    if ($miRestore) {
        $miRestore.Add_Click({
            $defaults = Get-BuiltInScriptDefaults
            $def = $defaults | Where-Object { $_.BuiltInKey -eq $captEntry.BuiltInKey }
            if (-not $def) { return }

            $db = @(Get-ScriptsDatabase)
            foreach ($e in $db) {
                if ($e.BuiltInKey -eq $captEntry.BuiltInKey) {
                    $e.Name            = $def.Name
                    $e.IconFile        = $def.IconFile
                    $e.ExecutionMethod = $def.ExecutionMethod
                    break
                }
            }
            Save-ScriptsDatabase -Database $db

            # Actualizar UI
            $captItem.Text  = $def.Name
            $captItem.Image = $null
            if (-not [string]::IsNullOrWhiteSpace($def.IconFile)) {
                $ip = Join-Path $Global:ScriptRoot "icos\$($def.IconFile)"
                if (Test-Path $ip) { try { $captItem.Image = [System.Drawing.Image]::FromFile($ip) } catch {} }
            }

            $captEntry.Name            = $def.Name
            $captEntry.IconFile        = $def.IconFile
            $captEntry.ExecutionMethod = $def.ExecutionMethod
            $captIcon = $def.IconFile

            [System.Windows.Forms.MessageBox]::Show(
                "Valores restaurados a los predeterminados.",
                "Restaurado",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }.GetNewClosure())
    }

    # Mantener dropdown abierto mientras el menú contextual está visible
    $ctx.Add_Closed({
        $captParent.DropDown.AutoClose = $true
    }.GetNewClosure())

    # Clic derecho -> evitar cierre del dropdown y mostrar menú contextual
    $item.Add_MouseDown({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            $captParent.DropDown.AutoClose = $false
            $ctx.Show([System.Windows.Forms.Cursor]::Position)
        }
    }.GetNewClosure())

    return $item
}

#==================================================================
# BLOQUE: Custom Handlers (funciones con UI propia)
#==================================================================

function Invoke-Scripts_ConfigurarProxy {
    $cn = Get-RemoteComputer
    if ([string]::IsNullOrWhiteSpace($cn)) { Show-NoComputer; return }
    $proxyPacUrl = 'http://proxy.example.local/proxy.pac'
    if (Get-Command 'Get-AppSettingValue' -ErrorAction SilentlyContinue) {
        $proxyPacUrl = Get-AppSettingValue -Key 'ProxyPacUrl' -DefaultValue $proxyPacUrl
    }

    $ok = Confirm-Action `
        "Desea configurar el proxy en el equipo '$cn' para el usuario activo?`n`nProxy PAC: $proxyPacUrl" `
        "Configurar Proxy"
    if (-not $ok) { return }

    if (-not (Test-Connection -ComputerName $cn -Count 1 -Quiet)) {
        Show-Error "El equipo '$cn' no responde al ping.`nCompruebe que está encendido y accesible." "Sin Conectividad"
        return
    }

    try {
        $profiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $cn -ErrorAction Stop
        $loaded   = $profiles | Where-Object { $_.Loaded -eq $true }
        if (-not $loaded) {
            Show-Error "No se encontro ningun perfil activo ('Loaded') en '$cn'." "Usuario no encontrado"
            return
        }
        $sid = $loaded[0].SID

        $remoteReg  = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey(
            [Microsoft.Win32.RegistryHive]::Users, $cn
        )
        $subKeyPath = "$sid\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        $subKey     = $remoteReg.OpenSubKey($subKeyPath, $true)
        if (-not $subKey) { $subKey = $remoteReg.CreateSubKey($subKeyPath) }

        $subKey.SetValue("AutoConfigURL", $proxyPacUrl,
            [Microsoft.Win32.RegistryValueKind]::String)
        $subKey.SetValue("AutoDetect", 1, [Microsoft.Win32.RegistryValueKind]::DWord)
        $subKey.Close()
        $remoteReg.Close()

        Add-Logs -text "$cn - Proxy configurado correctamente"
        [System.Windows.Forms.MessageBox]::Show(
            "Proxy configurado correctamente en '$cn'.",
            "Proxy Configurado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

    } catch {
        Show-Error "Error al configurar proxy en '$cn':`n$($_.Exception.Message)" "Error"
    }
}

function Invoke-Scripts_DesinstalarKB {
    $cn = Get-RemoteComputer
    if ([string]::IsNullOrWhiteSpace($cn)) { Show-NoComputer; return }

    $ps1Path = Join-Path $Global:ScriptRoot 'scripts\desinstalar_kb.ps1'
    if (-not (Test-Path $ps1Path)) {
        Show-Error "No se encuentra el script:`n$ps1Path" "Script no encontrado"
        return
    }

    # Formulario para introducir el número de KB
    $kbForm = New-Object System.Windows.Forms.Form
    $kbForm.Text            = "Desinstalar actualización KB"
    $kbForm.Size            = New-Object System.Drawing.Size(360, 160)
    $kbForm.StartPosition   = "CenterScreen"
    $kbForm.TopMost         = $true
    $kbForm.FormBorderStyle = "FixedDialog"
    $kbForm.MaximizeBox     = $false

    $lblKB = New-Object System.Windows.Forms.Label
    $lblKB.Text     = "Número de KB a desinstalar (sin 'KB'):"
    $lblKB.Location = New-Object System.Drawing.Point(12, 16)
    $lblKB.Size     = New-Object System.Drawing.Size(320, 18)
    $kbForm.Controls.Add($lblKB)

    $txtKB = New-Object System.Windows.Forms.TextBox
    $txtKB.Location = New-Object System.Drawing.Point(12, 40)
    $txtKB.Size     = New-Object System.Drawing.Size(320, 24)
    $kbForm.Controls.Add($txtKB)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text         = "Aceptar"
    $btnOk.Location     = New-Object System.Drawing.Point(150, 80)
    $btnOk.Size         = New-Object System.Drawing.Size(80, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $kbForm.Controls.Add($btnOk)
    $kbForm.AcceptButton = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Cancelar"
    $btnCancel.Location     = New-Object System.Drawing.Point(245, 80)
    $btnCancel.Size         = New-Object System.Drawing.Size(80, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $kbForm.Controls.Add($btnCancel)

    if ($kbForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $kbNumber = $txtKB.Text.Trim() -replace '^KB', ''
    if ([string]::IsNullOrWhiteSpace($kbNumber)) {
        Show-Error "Debe introducir un número de KB válido." "KB no especificado"
        return
    }

    $ok = Confirm-Action `
        "Desea desinstalar KB$kbNumber del equipo '$cn'?`n`nEl proceso se ejecutará en la sesión del usuario activo." `
        "Desinstalar KB$kbNumber"
    if (-not $ok) { return }

    Add-Logs -text "$cn - Desinstalar KB$kbNumber lanzado"
    Start-Process -FilePath "powershell.exe" `
        -ArgumentList "-NoExit -ExecutionPolicy Bypass -NoProfile -File `"$ps1Path`" -ComputerName `"$cn`" -KB `"$kbNumber`""
}

function Invoke-Scripts_RenombrarPerfil {
    $cn = Get-RemoteComputer
    if ([string]::IsNullOrWhiteSpace($cn)) { Show-NoComputer; return }

    function local:Get-PerfilesRemotos {
        param([string]$Equipo)
        try {
            $rutaPerfiles = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
            $regKey   = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Equipo)
            $subKey   = $regKey.OpenSubKey($rutaPerfiles)
            $perfiles = @()
            foreach ($sid in $subKey.GetSubKeyNames()) {
                $subClave = $subKey.OpenSubKey($sid)
                $usuario  = $subClave.GetValue("ProfileImagePath") -replace '^.*\\', ''
                $perfiles += [PSCustomObject]@{ Usuario = $usuario; SID = $sid }
            }
            $regKey.Close()
            return $perfiles
        } catch {
            Show-Error "Error obteniendo perfiles de '$Equipo':`n$_" "Error"
            return $null
        }
    }

    $perfiles = local:Get-PerfilesRemotos -Equipo $cn
    if (-not $perfiles) {
        [System.Windows.Forms.MessageBox]::Show(
            "No se encontraron perfiles en '$cn'.",
            "Sin perfiles",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = "Renombrar Perfil - $cn"
    $form.Width           = 420
    $form.Height          = 340
    $form.StartPosition   = "CenterScreen"
    $form.TopMost         = $true
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false

    $lblTitulo = New-Object System.Windows.Forms.Label
    $lblTitulo.Text      = "Selecciona o escribe el usuario a renombrar:"
    $lblTitulo.Location  = New-Object System.Drawing.Point(12, 12)
    $lblTitulo.Size      = New-Object System.Drawing.Size(380, 20)
    $form.Controls.Add($lblTitulo)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(12, 38)
    $listBox.Size     = New-Object System.Drawing.Size(380, 140)
    $perfiles | ForEach-Object { $listBox.Items.Add($_.Usuario) }
    $form.Controls.Add($listBox)

    $lblManual = New-Object System.Windows.Forms.Label
    $lblManual.Text     = "O introduce manualmente:"
    $lblManual.Location = New-Object System.Drawing.Point(12, 188)
    $lblManual.Size     = New-Object System.Drawing.Size(200, 18)
    $form.Controls.Add($lblManual)

    $txtManual = New-Object System.Windows.Forms.TextBox
    $txtManual.Location = New-Object System.Drawing.Point(12, 210)
    $txtManual.Size     = New-Object System.Drawing.Size(380, 24)
    $form.Controls.Add($txtManual)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text         = "Aceptar"
    $btnOk.Location     = New-Object System.Drawing.Point(210, 250)
    $btnOk.Size         = New-Object System.Drawing.Size(85, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)
    $form.AcceptButton  = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Cancelar"
    $btnCancel.Location     = New-Object System.Drawing.Point(305, 250)
    $btnCancel.Size         = New-Object System.Drawing.Size(85, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $usuarioSel = if (-not [string]::IsNullOrWhiteSpace($txtManual.Text)) {
        $txtManual.Text.Trim()
    } elseif ($listBox.SelectedItem) {
        $listBox.SelectedItem
    } else {
        Show-Error "No se seleccionó ni introdujo ningún usuario." "Error"; return
    }

    $sidUsuario = ($perfiles | Where-Object { $_.Usuario -eq $usuarioSel } | Select-Object -First 1).SID
    if (-not $sidUsuario) {
        Show-Error "No se encontró un perfil para '$usuarioSel' en '$cn'." "Usuario no encontrado"
        return
    }

    $ok = Confirm-Action `
        "Se va a eliminar del registro el perfil '$usuarioSel' (SID: $sidUsuario)`ny renombrar su carpeta a ${usuarioSel}_old en '$cn'.`n`nContinuar?" `
        "Confirmar Renombrar Perfil"
    if (-not $ok) { return }

    try {
        $clave  = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sidUsuario"
        $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $cn)
        if ($regKey.OpenSubKey($clave)) {
            $regKey.DeleteSubKeyTree($clave)
            Add-Logs -text "$cn - Clave de registro eliminada: $clave"
        }
        $regKey.Close()
    } catch {
        Show-Error "Error eliminando clave de registro:`n$_" "Error"
    }

    $perfilPath = "\\$cn\C`$\Users\$usuarioSel"
    if (Test-Path $perfilPath) {
        $nuevoPath = "\\$cn\C`$\Users\${usuarioSel}_old"
        try {
            Rename-Item -Path $perfilPath -NewName $nuevoPath -ErrorAction Stop
            Add-Logs -text "$cn - Carpeta renombrada: $nuevoPath"
            [System.Windows.Forms.MessageBox]::Show(
                "Perfil '$usuarioSel' procesado correctamente en '$cn':`n- Clave de registro eliminada`n- Carpeta renombrada a ${usuarioSel}_old",
                "Operación completada",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        } catch {
            Show-Error "Error renombrando carpeta del perfil:`n$_" "Error"
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "La carpeta del usuario no existe en la ruta:`n$perfilPath`n`n(La clave de registro ya fue eliminada si correspondia.)",
            "Carpeta no encontrada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
}

#==================================================================
# BLOQUE: Añadir script en caliente
#==================================================================

function Invoke-Scripts_AddScript {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripMenuItem]$ParentMenu
    )

    # 1) Seleccionar archivo de script
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title            = "Seleccionar Script a Añadir"
    $ofd.Filter           = "Scripts (*.ps1;*.bat;*.cmd)|*.ps1;*.bat;*.cmd|PowerShell (*.ps1)|*.ps1|Batch (*.bat;*.cmd)|*.bat;*.cmd|Todos|*.*"
    $ofd.InitialDirectory = [System.Environment]::GetFolderPath("Desktop")

    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $srcPath  = $ofd.FileName
    $fileName = [System.IO.Path]::GetFileName($srcPath)
    $dstDir   = Join-Path $Global:ScriptRoot 'scripts'
    $dstPath  = Join-Path $dstDir $fileName

    # 2) Comprobar si ya existe
    if (Test-Path $dstPath) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            "Ya existe '$fileName' en la carpeta scripts.`nDesea sobreescribirlo?",
            "Archivo existente",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    # 3) Copiar script local
    try {
        Copy-Item -Path $srcPath -Destination $dstPath -Force -ErrorAction Stop
    } catch {
        Show-Error "Error al copiar '$fileName':`n$_" "Error al copiar"
        return
    }

    # 4) Método por defecto según extensión
    $ext = [System.IO.Path]::GetExtension($fileName).ToLower()
    $defaultMethod = switch ($ext) {
        '.ps1' { 'standard' }
        '.bat' { 'batch-remote' }
        '.cmd' { 'batch-remote' }
        default { 'local' }
    }

    # 5) Diálogo compartido: nombre + icono + método de ejecución
    $result = Show-ScriptConfigDialog `
        -InitialName ([System.IO.Path]::GetFileNameWithoutExtension($fileName)) `
        -FileName $fileName `
        -InitialMethod $defaultMethod
    if ($null -eq $result) { return }

    # 6) Comprobar duplicado en BD
    $alreadyInDb = @(Get-ScriptsDatabase) | Where-Object { $_.FileName -eq $fileName -and -not $_.BuiltIn }

    # 7) Guardar en BD
    if (-not $alreadyInDb) {
        $db = @(Get-ScriptsDatabase)
        $db += [PSCustomObject]@{
            Name            = $result.Name
            FileName        = $fileName
            IconFile        = $result.IconFile
            Category        = 'Custom'
            ExecutionMethod = $result.ExecutionMethod
            CustomHandler   = ''
            BuiltIn         = $false
            BuiltInKey      = ''
            SortOrder       = 100
            AddedOn         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
        Save-ScriptsDatabase -Database $db

        # Preguntar si replicar en el servidor
        if ((Get-Command 'Test-ServerAvailable' -ErrorAction SilentlyContinue) -and (Test-ServerAvailable)) {
            $rRep = [System.Windows.Forms.MessageBox]::Show(
                "¿Desea replicar el script en el servidor?`n`nScript: $($result.Name)`nArchivo: $fileName",
                "Replicar en Servidor",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($rRep -eq [System.Windows.Forms.DialogResult]::Yes) {
                Sync-ScriptsJsonToServer | Out-Null
                Publish-ScriptToServer -LocalScriptPath $dstPath | Out-Null
                if (-not [string]::IsNullOrWhiteSpace($result.IconFile)) {
                    $icoPath = Join-Path $Global:ScriptRoot "icos\$($result.IconFile)"
                    if (Test-Path $icoPath) { Publish-IconToServer -LocalIconPath $icoPath | Out-Null }
                }
            }
        }

        # 8) Añadir al menú
        $newEntry = $db[-1]
        $newItem  = New-ScriptMenuItem -ScriptEntry $newEntry -ParentMenu $ParentMenu
        $count    = $ParentMenu.DropDownItems.Count
        $insertAt = if ($count -ge 2) { $count - 2 } else { $count }
        $ParentMenu.DropDownItems.Insert($insertAt, $newItem)
    }

    [System.Windows.Forms.MessageBox]::Show(
        "Script '$($result.Name)' añadido al menú correctamente.`nArchivo copiado a: scripts\$fileName",
        "Script Añadido",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

#==================================================================
# BLOQUE: Inicialización del menú Scripts
#         Llamar desde el script principal DESPUÉS de que todos los
#         controles UI estén creados y configurados.
#==================================================================

function Initialize-ScriptsMenu {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripMenuItem]$Menu
    )

    $Menu.DropDownItems.Clear()

    # Registrar scripts built-in (crea entradas si no existen en BD)
    $db = Register-BuiltInScripts

    # Orden de categorías predefinido
    $categoryOrder = @('Configuracion', 'Utilidades', 'Workaround')

    # Agrupar scripts por categoría
    $grouped = @{}
    foreach ($entry in $db) {
        $cat = if ($entry.Category) { $entry.Category } else { 'Custom' }
        if (-not $grouped.ContainsKey($cat)) { $grouped[$cat] = @() }
        $grouped[$cat] += $entry
    }

    # Ordenar dentro de cada grupo por SortOrder
    foreach ($cat in @($grouped.Keys)) {
        $grouped[$cat] = @($grouped[$cat] | Sort-Object { [int]$_.SortOrder })
    }

    # Construir menú en orden de categorías predefinidas
    $isFirst = $true
    foreach ($cat in $categoryOrder) {
        if (-not $grouped.ContainsKey($cat)) { continue }

        if (-not $isFirst) {
            [void]$Menu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        }
        [void]$Menu.DropDownItems.Add((New-MenuHeader -Text "-- $cat --"))

        foreach ($entry in $grouped[$cat]) {
            $menuItem = New-ScriptMenuItem -ScriptEntry $entry -ParentMenu $Menu
            [void]$Menu.DropDownItems.Add($menuItem)
        }

        $isFirst = $false
    }

    # Categorías no predefinidas (Custom y las que añada el usuario)
    foreach ($cat in ($grouped.Keys | Where-Object { $_ -notin $categoryOrder })) {
        [void]$Menu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        [void]$Menu.DropDownItems.Add((New-MenuHeader -Text "-- $cat --"))

        foreach ($entry in $grouped[$cat]) {
            $menuItem = New-ScriptMenuItem -ScriptEntry $entry -ParentMenu $Menu
            [void]$Menu.DropDownItems.Add($menuItem)
        }
    }

    # -- SEPARADOR + Añadir script... (SIEMPRE AL FINAL) --
    [void]$Menu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $addItem           = New-Object System.Windows.Forms.ToolStripMenuItem
    $addItem.Text      = "Añadir script..."
    $addItem.ForeColor = [System.Drawing.Color]::Gray
    $capturedMenu      = $Menu
    $addItem.Add_Click({ Invoke-Scripts_AddScript -ParentMenu $capturedMenu }.GetNewClosure())
    [void]$Menu.DropDownItems.Add($addItem)
}

Export-ModuleMember -Function Initialize-ScriptsMenu, `
    Invoke-Scripts_ConfigurarProxy, Invoke-Scripts_DesinstalarKB, `
    Invoke-Scripts_RenombrarPerfil, Invoke-Scripts_AddScript, `
    Get-ScriptsDatabase, Save-ScriptsDatabase, Register-BuiltInScripts, `
    Get-BuiltInScriptDefaults, Get-ExecutionMethods, `
    Show-ScriptConfigDialog, Show-ExecutionMethodInfoDialog, `
    New-ScriptMenuItem, Invoke-ScriptByMethod
