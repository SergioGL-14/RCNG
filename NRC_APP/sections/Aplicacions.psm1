# SecciÃ³n del menÃº AplicaciÃ³ns.
# Combina subaplicaciones fijas del proyecto con accesos externos aÃ±adidos
# por el usuario y guardados en apps_db.json.
param()

#==================================================================
# BLOQUE: Helpers privados de UI
#==================================================================

# Crea un encabezado visual dentro del menú de aplicaciones.
function New-AppsHeader {
    param([string]$Text)
    $item           = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text      = $Text
    $item.Enabled   = $false
    $item.Font      = New-Object System.Drawing.Font(
        "Segoe UI", 8.5,
        [System.Drawing.FontStyle]::Bold
    )
    return $item
}

# Crea una entrada de menú con icono opcional y manejador de clic.
function New-AppsMenuItem {
    param(
        [string]$Text,
        [string]$IconName,
        [scriptblock]$OnClick
    )
    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text = $Text

    if ($IconName) {
        $iconPath = Join-Path $Global:ScriptRoot "icos\$IconName"
        if (Test-Path $iconPath) {
            try { $item.Image = [System.Drawing.Image]::FromFile($iconPath) } catch {}
        }
    }

    if ($OnClick) {
        $capturedAction = $OnClick
        $item.Add_Click($capturedAction.GetNewClosure())
    }

    return $item
}

#==================================================================
# BLOQUE: Acciones individuales de cada aplicación
#==================================================================

# Copia el lado remoto del chat al %TEMP% del usuario y abre la consola local
# de soporte apuntando al archivo compartido de la conversación.
function Invoke-Apps_ChatRemoto {
    try {
        $remoteComputerName = $global:textbox_computername.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($remoteComputerName)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Por favor, introduzca un nombre de equipo válido.",
                "Chat Remoto",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        Add-Logs -text "$remoteComputerName - Verificando conectividad para Chat Remoto..."
        if (-not (Test-Connection -ComputerName $remoteComputerName -Count 1 -Quiet)) {
            [System.Windows.Forms.MessageBox]::Show(
                "El equipo '$remoteComputerName' no está accesible.",
                "Chat Remoto - Error de Conexión",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            return
        }

        Add-Logs -text "$remoteComputerName - Comprobando sesión de usuario activa..."
        try {
            $sessionInfo = (Get-WmiObject -Query "SELECT * FROM Win32_ComputerSystem" `
                -ComputerName $remoteComputerName -ErrorAction Stop | Select-Object -ExpandProperty UserName)

            if ([string]::IsNullOrWhiteSpace($sessionInfo)) {
                Add-Logs -text "$remoteComputerName - No se detectó ningún usuario logueado."
                [System.Windows.Forms.MessageBox]::Show(
                    "No hay ninguna sesión de usuario activa en el equipo '$remoteComputerName'.",
                    "Chat Remoto - Sin Sesión Activa",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            Add-Logs -text "$remoteComputerName - Usuario activo: $sessionInfo"
            $remoteUser = $sessionInfo.Split('\')[1]

        } catch {
            Add-Logs -text "$remoteComputerName - Error al verificar sesión: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show(
                "No se pudo verificar la sesión en el equipo '$remoteComputerName'.",
                "Chat Remoto - Error de Verificación",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            return
        }

        $chatScriptLocal   = Join-Path $Global:ScriptRoot 'app\Chat\chat.ps1'
        $chatScriptUsuario = Join-Path $Global:ScriptRoot 'app\Chat\chat_usu.ps1'

        foreach ($f in @($chatScriptLocal, $chatScriptUsuario)) {
            if (-not (Test-Path $f)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "No se encuentra el archivo:`n$f",
                    "Chat Remoto - Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
                return
            }
        }

        $remoteTempPath = "\\$remoteComputerName\C`$\Users\$remoteUser\AppData\Local\Temp"
        if (-not (Test-Path $remoteTempPath)) {
            New-Item -Path $remoteTempPath -ItemType Directory -Force | Out-Null
        }

        Copy-Item -Path $chatScriptUsuario -Destination "$remoteTempPath\chat_usu.ps1" -Force
        "" | Out-File -FilePath "$remoteTempPath\chat.txt" -Encoding utf8 -Force

        # Limpiar archivos de control de sesiones anteriores
        foreach ($sigFile in @("chat_usu.pid", "chat_close.signal")) {
            $sigPath = "$remoteTempPath\$sigFile"
            if (Test-Path $sigPath) { Remove-Item $sigPath -Force -ErrorAction SilentlyContinue }
        }
        Add-Logs -text "$remoteComputerName - Archivos de chat copiados a %temp% remoto."

        $psexecPath = Join-Path $Global:ScriptRoot 'tools\PsExec.exe'
        if (-not (Test-Path $psexecPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "No se encuentra PsExec.exe en: tools\PsExec.exe",
                "Chat Remoto - Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            return
        }

        $psexecArgs = "\\$remoteComputerName -s -i -d -accepteula powershell.exe " +
                      "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass " +
                      "-File `"C:\Users\$remoteUser\AppData\Local\Temp\chat_usu.ps1`""
        Start-Process -FilePath $psexecPath -ArgumentList $psexecArgs -WindowStyle Hidden -PassThru | Out-Null
        Add-Logs -text "$remoteComputerName - chat_usu.ps1 lanzado via PsExec."

        Start-Sleep -Seconds 1

        $chatRemotePath = "$remoteTempPath\chat.txt"
        $psCommand = "`$RutaArchivo = '$chatRemotePath'; & '$chatScriptLocal'"
        Start-Process powershell.exe -ArgumentList "-WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -Command `"$psCommand`"" -WindowStyle Hidden -PassThru | Out-Null

        Add-Logs -text "$remoteComputerName - Chat iniciado correctamente."
        [System.Windows.Forms.MessageBox]::Show(
            "Chat iniciado con el equipo '$remoteComputerName'.`nAmbas ventanas de chat están ahora abiertas.",
            "Chat Remoto",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

    } catch {
        $errorMsg = $_.Exception.Message
        Add-Logs -text "Chat Remoto - Error: $errorMsg"
        [System.Windows.Forms.MessageBox]::Show(
            "No se pudo iniciar el chat. Error: $errorMsg",
            "Chat Remoto - Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

# ExtFinder: busca extensiones instaladas en el equipo remoto.
function Invoke-Apps_ExtFinder {
    $scriptPath = Join-Path $Global:ScriptRoot 'app\ExtFinder\ExtFinder.ps1'
    if (Test-Path $scriptPath) {
        Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "No se encontró el script:`n$scriptPath",
            "ExtFinder", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
}

#==================================================================
# BLOQUE: Base de datos de aplicaciones externas (JSON)
#==================================================================

function Get-AppsDatabase {
    # Leer desde el servidor si está disponible (vía SharedDataManager), fallback local
    if (Get-Command 'Get-AppsDbContent' -ErrorAction SilentlyContinue) {
        $raw = Get-AppsDbContent
    } else {
        $dbPath = Join-Path $Global:ScriptRoot 'database\apps_db.json'
        $raw = if (Test-Path $dbPath) { Get-Content $dbPath -Raw -Encoding UTF8 } else { '[]' }
    }
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    try {
        $data = $raw | ConvertFrom-Json
        if ($data -isnot [array]) { $data = @($data) }
        return $data
    } catch { return @() }
}

function Add-AppToDatabase {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        [string]$AppPath
        ,
        [string]$IconName = ''
    )
    $db = @(Get-AppsDatabase)
    $db += [PSCustomObject]@{
        Name     = $Name
        AppPath  = $AppPath
        IconName = $IconName
        AddedOn  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
    $json = $db | ConvertTo-Json -Depth 3
    # Guardar SOLO en local. Para replicar al servidor usar Sync-AppsJsonToServer.
    if (Get-Command 'Save-AppsDbLocal' -ErrorAction SilentlyContinue) {
        Save-AppsDbLocal -JsonContent $json
    } else {
        $dbPath = Join-Path $Global:ScriptRoot 'database\apps_db.json'
        $json | Out-File -FilePath $dbPath -Encoding UTF8 -Force
    }
}

function Remove-AppFromDatabase {
    param(
        [Parameter(Mandatory=$true)]
        [string]$AppPath
    )
    $db = @(Get-AppsDatabase | Where-Object { $_.AppPath -ne $AppPath })
    $json = if ($db.Count -eq 0) { '[]' } else { $db | ConvertTo-Json -Depth 3 }
    # Guardar SOLO en local. Para replicar al servidor usar Sync-AppsJsonToServer.
    if (Get-Command 'Save-AppsDbLocal' -ErrorAction SilentlyContinue) {
        Save-AppsDbLocal -JsonContent $json
    } else {
        $dbPath = Join-Path $Global:ScriptRoot 'database\apps_db.json'
        $json | Out-File -FilePath $dbPath -Encoding UTF8 -Force
    }
}

function Show-AddAppDialog {
    param(
        [string]$InitialName = '',
        [string]$InitialPath = '',
        [string]$Title       = "Añadir aplicación externa",
        [string]$InitialIcon = ''
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = $Title
    $form.Size            = New-Object System.Drawing.Size(440, 240)
    $form.StartPosition   = "CenterScreen"
    $form.TopMost         = $true
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text     = "Nombre a mostrar:"
    $lblName.Location = New-Object System.Drawing.Point(12, 20)
    $lblName.Size     = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($lblName)

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(140, 17)
    $txtName.Size     = New-Object System.Drawing.Size(270, 23)
    $form.Controls.Add($txtName)

    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text     = "Ruta del ejecutable:"
    $lblPath.Location = New-Object System.Drawing.Point(12, 55)
    $lblPath.Size     = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($lblPath)

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point(140, 52)
    $txtPath.Size     = New-Object System.Drawing.Size(230, 23)
    $form.Controls.Add($txtPath)

    $lblIcon = New-Object System.Windows.Forms.Label
    $lblIcon.Text     = "Icono (archivo .ico):"
    $lblIcon.Location = New-Object System.Drawing.Point(12, 90)
    $lblIcon.Size     = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($lblIcon)

    $txtIcon = New-Object System.Windows.Forms.TextBox
    $txtIcon.Location = New-Object System.Drawing.Point(140, 88)
    $txtIcon.Size     = New-Object System.Drawing.Size(230, 23)
    $form.Controls.Add($txtIcon)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text     = "..."
    $btnBrowse.Location = New-Object System.Drawing.Point(375, 51)
    $btnBrowse.Size     = New-Object System.Drawing.Size(35, 25)
    $btnBrowse.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "Ejecutables (*.exe;*.ps1;*.bat)|*.exe;*.ps1;*.bat|Todos (*.*)|*.*"
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtPath.Text = $ofd.FileName
            if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
                $txtName.Text = [System.IO.Path]::GetFileNameWithoutExtension($ofd.FileName)
            }
        }
    })
    $form.Controls.Add($btnBrowse)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text           = "Aceptar"
    $btnOK.DialogResult   = [System.Windows.Forms.DialogResult]::OK
    $btnOK.Location       = New-Object System.Drawing.Point(240, 160)
    $btnOK.Size           = New-Object System.Drawing.Size(80, 28)
    $form.Controls.Add($btnOK)
    $form.AcceptButton    = $btnOK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Cancelar"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.Location     = New-Object System.Drawing.Point(330, 160)
    $btnCancel.Size         = New-Object System.Drawing.Size(80, 28)
    $form.Controls.Add($btnCancel)
    $form.CancelButton      = $btnCancel

    # Rellenar valores iniciales si se proporcionan (modo edición)
    $txtName.Text = $InitialName
    $txtPath.Text = $InitialPath
    $txtIcon.Text = $InitialIcon

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    $name = $txtName.Text.Trim()
    $path = $txtPath.Text.Trim()
    $icon = $txtIcon.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($path)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Debe especificar un nombre y una ruta.",
            "Datos incompletos",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $null
    }

    return @{ Name = $name; AppPath = $path; IconName = $icon }
}

#------------------------------------------------------------------
# Crea un ítem dinámico para aplicación externa con menú contextual
# de clic derecho para eliminarla.
#------------------------------------------------------------------
function New-ExternalAppMenuItem {
    param(
        [Parameter(Mandatory=$true)]
        [string]$AppName,
        [Parameter(Mandatory=$true)]
        [string]$AppPath,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripMenuItem]$ParentMenu,
        [string]$IconName
    )

    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text = $AppName

    if ($IconName) {
        $iconPath = Join-Path $Global:ScriptRoot "icos\$IconName"
        if (Test-Path $iconPath) {
            try { $item.Image = [System.Drawing.Image]::FromFile($iconPath) } catch {}
        }
    }

    $captName   = $AppName
    $captPath   = $AppPath
    $captItem   = $item
    $captParent = $ParentMenu
    $captIcon   = $IconName

    $item.Add_Click({
        $p = $captPath
        if (-not (Test-Path $p)) {
            [System.Windows.Forms.MessageBox]::Show(
                "No se encontró la aplicación:`n$p",
                "Archivo no encontrado",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            return
        }
        Add-Logs -text "Aplicación externa: $captName"
        Start-Process $p
    }.GetNewClosure())

    $ctx      = New-Object System.Windows.Forms.ContextMenuStrip
    $miModify = New-Object System.Windows.Forms.ToolStripMenuItem
    $miModify.Text = "Modificar"
    $miDelete = New-Object System.Windows.Forms.ToolStripMenuItem
    $miDelete.Text = "Eliminar"
    [void]$ctx.Items.Add($miModify)
    [void]$ctx.Items.Add($miDelete)

    $miModify.Add_Click({
        $res = Show-AddAppDialog -InitialName $captName -InitialPath $captPath -InitialIcon $captIcon -Title "Modificar aplicación"
        if ($null -eq $res) { return }
        Remove-AppFromDatabase -AppPath $captPath
        Add-AppToDatabase -Name $res.Name -AppPath $res.AppPath -IconName $res.IconName

        # Preguntar si replicar en el servidor
        if ((Get-Command 'Test-ServerAvailable' -ErrorAction SilentlyContinue) -and (Test-ServerAvailable)) {
            $rRep = [System.Windows.Forms.MessageBox]::Show(
                "¿Desea replicar los cambios en el servidor?`n`nAplicación: $($res.Name)",
                "Replicar en Servidor",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($rRep -eq [System.Windows.Forms.DialogResult]::Yes) {
                Sync-AppsJsonToServer | Out-Null
                if (-not [string]::IsNullOrWhiteSpace($res.IconName)) {
                    $icoPath = Join-Path $Global:ScriptRoot "icos\$($res.IconName)"
                    if (Test-Path $icoPath) { Publish-IconToServer -LocalIconPath $icoPath | Out-Null }
                }
            }
        }

        Initialize-AplicacionsMenu -Menu $captParent
    }.GetNewClosure())

    $miDelete.Add_Click({
        $r = [System.Windows.Forms.MessageBox]::Show(
            "¿Eliminar '$captName' del menú?",
            "Confirmar Eliminación",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        Remove-AppFromDatabase -AppPath $captPath

        # Preguntar si eliminar también del servidor
        if ((Get-Command 'Test-ServerAvailable' -ErrorAction SilentlyContinue) -and (Test-ServerAvailable)) {
            $rSrv = [System.Windows.Forms.MessageBox]::Show(
                "¿Desea eliminar también la aplicación del servidor?`n`nNombre: $captName",
                "Eliminar del Servidor",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($rSrv -eq [System.Windows.Forms.DialogResult]::Yes) {
                Sync-AppsJsonToServer | Out-Null
            }
        }

        [void]$captParent.DropDownItems.Remove($captItem)
    }.GetNewClosure())

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
# BLOQUE: Inicialización del menú Aplicacións
#         Llamar desde el script principal DESPUÉS de que todos los
#         controles UI estén creados y configurados.
#==================================================================

function Initialize-AplicacionsMenu {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripMenuItem]$Menu
    )

    $Menu.DropDownItems.Clear()

    # -- APLICACIÓNS LOCALES --
    [void]$Menu.DropDownItems.Add((New-AppsHeader  -Text "-- Aplicacións Locales --"))
    [void]$Menu.DropDownItems.Add((New-AppsMenuItem -Text "Chat Remoto"       -IconName "Chat.ico"       -OnClick { Invoke-Apps_ChatRemoto       }))
    [void]$Menu.DropDownItems.Add((New-AppsMenuItem -Text "ExtFinder"         -IconName "ExtFinder.ico"  -OnClick { Invoke-Apps_ExtFinder        }))

    # -- APLICACIONES EXTERNAS (desde BD) --
    $appDb = @(Get-AppsDatabase)
    if ($appDb.Count -gt 0) {
        [void]$Menu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        [void]$Menu.DropDownItems.Add((New-AppsHeader -Text "-- Aplicaciones Externas --"))
        foreach ($entry in $appDb) {
            $extItem = New-ExternalAppMenuItem -AppName $entry.Name -AppPath $entry.AppPath -ParentMenu $Menu -IconName ($entry.IconName)
            [void]$Menu.DropDownItems.Add($extItem)
        }
    }

    # -- Añadir nuevas aplicaciones (SIEMPRE AL FINAL) --
    [void]$Menu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $addItem           = New-Object System.Windows.Forms.ToolStripMenuItem
    $addItem.Text      = "Añadir nuevas aplicaciones..."
    $addItem.ForeColor = [System.Drawing.Color]::Gray
    $capturedMenu      = $Menu
    $addItem.Add_Click({
        $res = Show-AddAppDialog
        if ($null -eq $res) { return }
        # Evitar duplicados por AppPath
        if (@(Get-AppsDatabase) | Where-Object { $_.AppPath -eq $res.AppPath }) {
            [System.Windows.Forms.MessageBox]::Show(
                "Ya existe una aplicación con esa ruta en el menú.",
                "Entrada duplicada",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }
        Add-AppToDatabase -Name $res.Name -AppPath $res.AppPath -IconName $res.IconName

        # Preguntar si replicar en el servidor
        if ((Get-Command 'Test-ServerAvailable' -ErrorAction SilentlyContinue) -and (Test-ServerAvailable)) {
            $rRep = [System.Windows.Forms.MessageBox]::Show(
                "¿Desea replicar la nueva aplicación en el servidor?`n`nAplicación: $($res.Name)",
                "Replicar en Servidor",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($rRep -eq [System.Windows.Forms.DialogResult]::Yes) {
                Sync-AppsJsonToServer | Out-Null
                if (-not [string]::IsNullOrWhiteSpace($res.IconName)) {
                    $icoPath = Join-Path $Global:ScriptRoot "icos\$($res.IconName)"
                    if (Test-Path $icoPath) { Publish-IconToServer -LocalIconPath $icoPath | Out-Null }
                }
            }
        }

        # Diferir el rebuild fuera del evento de clic actual para evitar que
        # DropDownItems.Clear() se llame mientras WinForms gestiona este evento
        $capMenu = $capturedMenu
        $tmr = New-Object System.Windows.Forms.Timer
        $tmr.Interval = 50
        $tmrRef = $tmr
        $tmr.Add_Tick({
            $tmrRef.Stop()
            $tmrRef.Dispose()
            Initialize-AplicacionsMenu -Menu $capMenu
        }.GetNewClosure())
        $tmr.Start()
    }.GetNewClosure())
    [void]$Menu.DropDownItems.Add($addItem)
}

Export-ModuleMember -Function Initialize-AplicacionsMenu, `
    Invoke-Apps_ChatRemoto, Invoke-Apps_ExtFinder, `
    Get-AppsDatabase, Add-AppToDatabase, Remove-AppFromDatabase, `
    Show-AddAppDialog, New-ExternalAppMenuItem


