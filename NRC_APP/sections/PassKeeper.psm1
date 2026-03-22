# Pass Keeper gestiona el panel lateral de valores rápidos.
# Cada entrada se cifra para el mismo equipo y se guarda en una ruta por
# usuario dentro de LocalAppData, no en la carpeta database del proyecto.
param()

#------------------------------------------------------------------
# Variables de módulo
#------------------------------------------------------------------
$Script:PKDataFile = $null
$Script:PKPanel    = $null

# Resuelve la ruta del `passkeeper.json` propia del usuario actual.
function Get-PKDataFilePath {
    param()
    try {
        $localApp = $env:LOCALAPPDATA
        if ([string]::IsNullOrWhiteSpace($localApp)) { $localApp = $env:USERPROFILE }
        $baseDir = Join-Path $localApp 'LazyWinAdmin'

        # Obtener el SID del usuario da una carpeta estable aunque cambie el nombre visible.
        $sid = (New-Object System.Security.Principal.NTAccount($env:USERNAME)).Translate([System.Security.Principal.SecurityIdentifier]).Value

        # Se usa un hash corto del SID para no dejar la ruta completamente en claro.
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $sidBytes = [System.Text.Encoding]::UTF8.GetBytes($sid)
        $hashBytes = $sha.ComputeHash($sidBytes)
        $sha.Dispose()
        $hex = ([System.BitConverter]::ToString($hashBytes)).Replace('-','').ToLower()
        $sub = $hex.Substring(0,16)

        $dir = Join-Path $baseDir ("pk_$sub")
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        return Join-Path $dir 'passkeeper.json'
    } catch {
        # Si algo falla, usar una ruta simple dentro del perfil local.
        $dir = Join-Path $env:LOCALAPPDATA 'LazyWinAdmin'
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        return Join-Path $dir 'passkeeper.json'
    }
}

#------------------------------------------------------------------
# Cifrado basado en el MachineGuid del equipo
#------------------------------------------------------------------

function Get-PKEncryptionKey {
    $guid = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Cryptography' -Name MachineGuid).MachineGuid
    $salt = [System.Text.Encoding]::UTF8.GetBytes('PK_NRC_v6_Salt_2026')
    $derive = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($guid, $salt, 10000)
    $key = $derive.GetBytes(32)
    $derive.Dispose()
    return $key
}

function Protect-PKValue {
    param([string]$Plain)
    $key = Get-PKEncryptionKey
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Key = $key
    $aes.GenerateIV()
    $encryptor  = $aes.CreateEncryptor()
    $plainBytes = [System.Text.Encoding]::UTF8.GetBytes($Plain)
    $encrypted  = $encryptor.TransformFinalBlock($plainBytes, 0, $plainBytes.Length)
    $result = New-Object byte[] ($aes.IV.Length + $encrypted.Length)
    [Array]::Copy($aes.IV, 0, $result, 0, $aes.IV.Length)
    [Array]::Copy($encrypted, 0, $result, $aes.IV.Length, $encrypted.Length)
    $aes.Dispose()
    return [Convert]::ToBase64String($result)
}

function Unprotect-PKValue {
    param([string]$Base64)
    $key  = Get-PKEncryptionKey
    $data = [Convert]::FromBase64String($Base64)
    $iv   = New-Object byte[] 16
    [Array]::Copy($data, 0, $iv, 0, 16)
    $cipherLen  = $data.Length - 16
    $ciphertext = New-Object byte[] $cipherLen
    [Array]::Copy($data, 16, $ciphertext, 0, $cipherLen)
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Key = $key
    $aes.IV  = $iv
    $decryptor  = $aes.CreateDecryptor()
    $plainBytes = $decryptor.TransformFinalBlock($ciphertext, 0, $ciphertext.Length)
    $aes.Dispose()
    return [System.Text.Encoding]::UTF8.GetString($plainBytes)
}

#------------------------------------------------------------------
# Persistencia JSON (siempre array, siempre -InputObject)
#------------------------------------------------------------------

function Get-PKEntries {
    if ([string]::IsNullOrWhiteSpace($Script:PKDataFile)) { $Script:PKDataFile = Get-PKDataFilePath }
    if (-not (Test-Path $Script:PKDataFile)) { return @() }
    try {
        $raw = [System.IO.File]::ReadAllText($Script:PKDataFile, [System.Text.Encoding]::UTF8)
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $parsed = ConvertFrom-Json -InputObject $raw
        if ($null -eq $parsed) { return @() }
        return @($parsed)
    } catch { return @() }
}

function Save-PKEntries {
    param([object[]]$Entries)
    if ([string]::IsNullOrWhiteSpace($Script:PKDataFile)) { $Script:PKDataFile = Get-PKDataFilePath }
    $dir = Split-Path $Script:PKDataFile
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $arr = @($Entries)
    $json = ConvertTo-Json -InputObject $arr -Depth 3
    [System.IO.File]::WriteAllText($Script:PKDataFile, $json, [System.Text.Encoding]::UTF8)
}

#------------------------------------------------------------------
# UI: Reconstruir botones (layout 2 columnas)
#------------------------------------------------------------------

function Refresh-PKButtons {
    if ($null -eq $Script:PKPanel) { return }

    $Script:PKPanel.SuspendLayout()
    $Script:PKPanel.Controls.Clear()

    $dataFile = $Script:PKDataFile
    $panel    = $Script:PKPanel
    $entries  = @(Get-PKEntries)

    # Capturar referencias a funciones del modulo para closures
    $unprotectFn = Get-Command Unprotect-PKValue
    $getPKFn     = Get-Command Get-PKEntries
    $savePKFn    = Get-Command Save-PKEntries
    $refreshFn   = Get-Command Refresh-PKButtons

    # Layout 2 columnas
    $margin    = 6
    $gap       = 6
    $totalW    = $panel.ClientSize.Width
    if ($totalW -lt 60) { $totalW = 340 }
    $colWidth  = [Math]::Floor(($totalW - $margin * 2 - $gap) / 2)
    $rowHeight = 28
    $rowGap    = 6
    $delWidth  = 22

    $col = 0
    $y   = $margin

    foreach ($entry in $entries) {
        $entryId    = $entry.Id
        $labelLocal = $entry.Label
        $encLocal   = $entry.EncryptedValue

        if ($col -eq 0) { $x = $margin } else { $x = $margin + $colWidth + $gap }

        # Contenedor de celda
        $cell          = New-Object System.Windows.Forms.Panel
        $cell.Location = New-Object System.Drawing.Point($x, $y)
        $cell.Size     = New-Object System.Drawing.Size($colWidth, $rowHeight)

        # Boton principal: copia valor al portapapeles
        $btn           = New-Object System.Windows.Forms.Button
        $btn.Dock      = 'Fill'
        $btn.Text      = $labelLocal
        $btn.FlatStyle = 'Flat'
        $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(180, 200, 220)
        $btn.BackColor = [System.Drawing.Color]::White
        $btn.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 60)
        $btn.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
        $btn.TextAlign = 'MiddleLeft'
        $btn.Padding   = New-Object System.Windows.Forms.Padding(4, 0, 0, 0)
        $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand

        $btn.Add_Click({
            try {
                $plain = & $unprotectFn -Base64 $encLocal
                [System.Windows.Forms.Clipboard]::SetText($plain)
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "No es posible descifrar esta entrada.`n`n" +
                    "Los datos pueden haber sido guardados en otro equipo " +
                    "o estar corruptos.`n`n" +
                    "Ve a: Configuracion -> Limpiar Pass Keeper",
                    "Pass Keeper",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        }.GetNewClosure())

        # Boton eliminar
        $btnDel           = New-Object System.Windows.Forms.Button
        $btnDel.Dock      = 'Right'
        $btnDel.Width     = $delWidth
        $btnDel.Text      = [char]0x2715
        $btnDel.FlatStyle = 'Flat'
        $btnDel.FlatAppearance.BorderSize = 0
        $btnDel.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 220, 220)
        $btnDel.BackColor = [System.Drawing.Color]::Transparent
        $btnDel.ForeColor = [System.Drawing.Color]::FromArgb(160, 60, 60)
        $btnDel.Font      = New-Object System.Drawing.Font('Segoe UI', 7)
        $btnDel.Cursor    = [System.Windows.Forms.Cursors]::Hand

        $btnDel.Add_Click({
            $res = [System.Windows.Forms.MessageBox]::Show(
                "Eliminar '$labelLocal'?",
                "Pass Keeper",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
                $allCurrent = @(& $getPKFn)
                $remaining  = @($allCurrent | Where-Object { $_.Id -ne $entryId })
                if ($remaining.Count -eq 0) {
                    [System.IO.File]::Delete($dataFile)
                } else {
                    & $savePKFn -Entries $remaining
                }
                & $refreshFn
            }
        }.GetNewClosure())

        $cell.Controls.Add($btn)     # Fill (index 0)
        $cell.Controls.Add($btnDel)  # Right (index 1)
        $panel.Controls.Add($cell)

        # Avanzar columna/fila
        $col++
        if ($col -ge 2) {
            $col = 0
            $y  += $rowHeight + $rowGap
        }
    }

    # Si terminamos en columna 1, avanzar fila
    if ($col -ne 0) { $y += $rowHeight + $rowGap }

    $minH = $y + $margin
    $Script:PKPanel.AutoScrollMinSize = New-Object System.Drawing.Size(0, $minH)
    $Script:PKPanel.ResumeLayout($true)
}

#------------------------------------------------------------------
# FUNCION PUBLICA: Initialize-PassKeeper
#------------------------------------------------------------------
function Initialize-PassKeeper {
    param(
        [System.Windows.Forms.Panel]$ButtonsPanel,
        [string]$DataFile
    )

    $Script:PKPanel = $ButtonsPanel

    # Usar ruta por-usuario ofuscada para el fichero de datos
    $userPath = Get-PKDataFilePath
    $Script:PKDataFile = $userPath

    $dir = Split-Path $Script:PKDataFile
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    # Migracion simple: si no existe el fichero por-usuario y se proporcionó uno antiguo, copiarlo
    if (-not (Test-Path $Script:PKDataFile) -and $null -ne $DataFile -and (Test-Path $DataFile)) {
        try {
            Copy-Item -Path $DataFile -Destination $Script:PKDataFile -Force
        } catch {
            # ignorar errores de migracion
        }
    }

    Refresh-PKButtons
}

#------------------------------------------------------------------
# FUNCION PUBLICA: Show-AddPKDialog
#------------------------------------------------------------------
function Show-AddPKDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Pass Keeper - Nueva entrada"
    $dlg.Size            = New-Object System.Drawing.Size(320, 200)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.BackColor       = [System.Drawing.Color]::WhiteSmoke

    $lblName          = New-Object System.Windows.Forms.Label
    $lblName.Text     = "Etiqueta:"
    $lblName.Location = New-Object System.Drawing.Point(14, 18)
    $lblName.Size     = New-Object System.Drawing.Size(72, 20)
    $lblName.Font     = New-Object System.Drawing.Font('Segoe UI', 9)

    $txtName          = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(90, 15)
    $txtName.Size     = New-Object System.Drawing.Size(200, 22)
    $txtName.Font     = New-Object System.Drawing.Font('Segoe UI', 9)

    $lblVal           = New-Object System.Windows.Forms.Label
    $lblVal.Text      = "Valor:"
    $lblVal.Location  = New-Object System.Drawing.Point(14, 52)
    $lblVal.Size      = New-Object System.Drawing.Size(72, 20)
    $lblVal.Font      = New-Object System.Drawing.Font('Segoe UI', 9)

    $txtVal                       = New-Object System.Windows.Forms.TextBox
    $txtVal.Location              = New-Object System.Drawing.Point(90, 49)
    $txtVal.Size                  = New-Object System.Drawing.Size(200, 22)
    $txtVal.Font                  = New-Object System.Drawing.Font('Segoe UI', 9)
    $txtVal.UseSystemPasswordChar = $true

    $chkShow          = New-Object System.Windows.Forms.CheckBox
    $chkShow.Text     = "Mostrar valor"
    $chkShow.Location = New-Object System.Drawing.Point(90, 76)
    $chkShow.Size     = New-Object System.Drawing.Size(110, 20)
    $chkShow.Font     = New-Object System.Drawing.Font('Segoe UI', 8)
    $chkShow.Add_CheckedChanged({ $txtVal.UseSystemPasswordChar = -not $chkShow.Checked })

    $btnOK              = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "Guardar"
    $btnOK.Location     = New-Object System.Drawing.Point(90, 108)
    $btnOK.Size         = New-Object System.Drawing.Size(96, 28)
    $btnOK.FlatStyle    = 'Flat'
    $btnOK.BackColor    = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnOK.ForeColor    = 'White'
    $btnOK.Font         = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnOK.DialogResult = 'OK'

    $btnCancel              = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Cancelar"
    $btnCancel.Location     = New-Object System.Drawing.Point(196, 108)
    $btnCancel.Size         = New-Object System.Drawing.Size(96, 28)
    $btnCancel.FlatStyle    = 'Flat'
    $btnCancel.Font         = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnCancel.DialogResult = 'Cancel'

    $dlg.AcceptButton = $btnOK
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.AddRange(@($lblName, $txtName, $lblVal, $txtVal, $chkShow, $btnOK, $btnCancel))

    $owner = [System.Windows.Forms.Application]::OpenForms[0]
    if ($dlg.ShowDialog($owner) -eq 'OK') {
        $label = $txtName.Text.Trim()
        $val   = $txtVal.Text
        if ($label -ne '' -and $val -ne '') {
            try {
                $enc     = Protect-PKValue -Plain $val
                $current = @(Get-PKEntries)
                $newEntry = [PSCustomObject]@{
                    Id             = [guid]::NewGuid().ToString()
                    Label          = $label
                    EncryptedValue = $enc
                }
                $updated = [System.Collections.ArrayList]::new()
                foreach ($e in $current) { [void]$updated.Add($e) }
                [void]$updated.Add($newEntry)
                Save-PKEntries -Entries $updated.ToArray()
                Refresh-PKButtons
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Error al guardar: $_",
                    "Pass Keeper",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    }
    $dlg.Dispose()
}

#------------------------------------------------------------------
# FUNCION PUBLICA: Clear-PassKeeperData
#------------------------------------------------------------------
function Clear-PassKeeperData {
    $res = [System.Windows.Forms.MessageBox]::Show(
        "Eliminar TODAS las entradas de Pass Keeper?`n`nEsta accion no se puede deshacer.",
        "Pass Keeper - Limpiar",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
        if ($null -ne $Script:PKDataFile -and (Test-Path $Script:PKDataFile)) {
            [System.IO.File]::Delete($Script:PKDataFile)
        }
        Refresh-PKButtons
    }
}

#==================================================================
# EXPORTACION
#==================================================================
Export-ModuleMember -Function Initialize-PassKeeper, Show-AddPKDialog, Clear-PassKeeperData

