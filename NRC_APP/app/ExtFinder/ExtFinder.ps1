# Directorio de extensiones.
# Trabaja sobre la tabla extensions de ComputerNames.sqlite y sincroniza la
# copia local con el recurso compartido cuando detecta cambios mÃ¡s recientes.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Rutas de trabajo y acceso a SQLite.
$ScriptRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllPath     = [System.IO.Path]::GetFullPath((Join-Path $ScriptRoot "..\..\libs\System.Data.SQLite.dll"))
$localDbPath = [System.IO.Path]::GetFullPath((Join-Path $ScriptRoot "..\..\database\ComputerNames.sqlite"))
$settingsPath = [System.IO.Path]::GetFullPath((Join-Path $ScriptRoot "..\..\database\appsettings.json"))
$serverBase  = '\\server\share\NRC_APP'
if (Test-Path $settingsPath) {
    try {
        $settings = Get-Content $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
        if ($settings.SharedServerBase) {
            $serverBase = [string]$settings.SharedServerBase
        }
    } catch {}
}
$serverDb    = Join-Path $serverBase 'ComputerNames.sqlite'
# La app usa siempre la copia local y solo sincroniza cuando hace falta.
$dbPath = $localDbPath

try {
    $loaded = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'System.Data.SQLite' }
    if (-not $loaded) { Add-Type -Path $dllPath -ErrorAction Stop }
} catch {
    [System.Windows.Forms.MessageBox]::Show("No se pudo cargar SQLite DLL:`n$($_.Exception.Message)","Error","OK","Error")
    return
}

#=========================
# Funciones de acceso a BD. AquÃ­ se usa DataReader porque la consulta es simple.
#=========================
$script:Conn = $null
$script:ServerAvail = $null

function Test-ExtFinderServer {
    if ($null -ne $script:ServerAvail) { return $script:ServerAvail }
    $script:ServerAvail = [System.IO.Directory]::Exists($serverBase)
    return $script:ServerAvail
}

function Sync-DbFromServer {
    <# Si el servidor tiene una versión más nueva de la BD, copia a local #>
    if (-not (Test-ExtFinderServer)) { return }
    if (-not (Test-Path $serverDb)) { return }
    try {
        if (-not (Test-Path $localDbPath)) {
            Copy-Item -Path $serverDb -Destination $localDbPath -Force
            return
        }
        $serverTime = (Get-Item $serverDb).LastWriteTime
        $localTime  = (Get-Item $localDbPath).LastWriteTime
        if ($serverTime -gt $localTime) {
            # Cerrar conexión abierta antes de copiar
            if ($script:Conn) {
                try { $script:Conn.Close(); $script:Conn.Dispose() } catch {}
                $script:Conn = $null
            }
            Copy-Item -Path $serverDb -Destination $localDbPath -Force
        }
    } catch {}
}

function Sync-DbToServer {
    <# Copia la BD local al servidor (tras confirmación del usuario) #>
    if (-not (Test-ExtFinderServer)) { return $false }
    try {
        Copy-Item -Path $localDbPath -Destination $serverDb -Force
        return $true
    } catch { return $false }
}

function Open-DB {
    if ($script:Conn -and $script:Conn.State -eq 'Open') { return $script:Conn }
    $connString = "Data Source=$dbPath;Version=3;Journal Mode=DELETE;BusyTimeout=5000;"
    $c = New-Object System.Data.SQLite.SQLiteConnection -ArgumentList $connString
    $c.ParseViaFramework = $true
    $c.Open()
    $script:Conn = $c
    return $c
}

function Search-Extensions {
    param([string]$Filter = "")
    $c = Open-DB
    $cmd = $c.CreateCommand()
    if ([string]::IsNullOrWhiteSpace($Filter)) {
        $cmd.CommandText = "SELECT equipo, extension FROM extensions ORDER BY equipo COLLATE NOCASE"
    } else {
        $cmd.CommandText = "SELECT equipo, extension FROM extensions WHERE equipo LIKE @f OR extension LIKE @f ORDER BY equipo COLLATE NOCASE"
        [void]$cmd.Parameters.AddWithValue('@f', "%$($Filter.Trim())%")
    }
    $results = [System.Collections.ArrayList]::new()
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        [void]$results.Add(@($reader.GetString(0), $reader.GetString(1)))
    }
    $reader.Close()
    return ,$results
}

function Save-Extension {
    param([string]$Equipo, [string]$Extension)
    $c = Open-DB
    $cmd = $c.CreateCommand()
    $cmd.CommandText = "INSERT OR REPLACE INTO extensions (equipo, extension) VALUES (@e, @x)"
    [void]$cmd.Parameters.AddWithValue('@e', $Equipo.Trim().ToUpper())
    [void]$cmd.Parameters.AddWithValue('@x', $Extension.Trim())
    [void]$cmd.ExecuteNonQuery()
}

function Remove-Extension {
    param([string]$Equipo)
    $c = Open-DB
    $cmd = $c.CreateCommand()
    $cmd.CommandText = "DELETE FROM extensions WHERE equipo = @e"
    [void]$cmd.Parameters.AddWithValue('@e', $Equipo.Trim().ToUpper())
    [void]$cmd.ExecuteNonQuery()
}

function Get-ExtCount {
    $c = Open-DB
    $cmd = $c.CreateCommand()
    $cmd.CommandText = "SELECT COUNT(*) FROM extensions"
    return [int]$cmd.ExecuteScalar()
}

#=========================
# UI - Formulario compacto
#=========================
$form               = New-Object System.Windows.Forms.Form
$form.Text           = "ExtFinder"
$form.Size           = New-Object System.Drawing.Size(480, 400)
$form.MinimumSize    = New-Object System.Drawing.Size(420, 320)
$form.StartPosition  = "CenterScreen"
$form.Font           = New-Object System.Drawing.Font("Segoe UI", 9)

# -- Barra busqueda --
$txtSearch           = New-Object System.Windows.Forms.TextBox
$txtSearch.Location  = New-Object System.Drawing.Point(8, 8)
$txtSearch.Size      = New-Object System.Drawing.Size(340, 24)
$txtSearch.Font      = New-Object System.Drawing.Font("Consolas", 10)
$txtSearch.Anchor    = [System.Windows.Forms.AnchorStyles]"Top,Left,Right"
$form.Controls.Add($txtSearch)

$btnBuscar           = New-Object System.Windows.Forms.Button
$btnBuscar.Text      = "Buscar"
$btnBuscar.Location  = New-Object System.Drawing.Point(354, 7)
$btnBuscar.Size      = New-Object System.Drawing.Size(54, 26)
$btnBuscar.Anchor    = [System.Windows.Forms.AnchorStyles]"Top,Right"
$form.Controls.Add($btnBuscar)

$btnTodos            = New-Object System.Windows.Forms.Button
$btnTodos.Text       = "Todos"
$btnTodos.Location   = New-Object System.Drawing.Point(412, 7)
$btnTodos.Size       = New-Object System.Drawing.Size(50, 26)
$btnTodos.Anchor     = [System.Windows.Forms.AnchorStyles]"Top,Right"
$form.Controls.Add($btnTodos)

# -- ListView --
$lv                  = New-Object System.Windows.Forms.ListView
$lv.View             = [System.Windows.Forms.View]::Details
$lv.FullRowSelect    = $true
$lv.GridLines        = $true
$lv.MultiSelect      = $false
$lv.Location         = New-Object System.Drawing.Point(8, 38)
$lv.Size             = New-Object System.Drawing.Size(454, 282)
$lv.Anchor           = [System.Windows.Forms.AnchorStyles]"Top,Bottom,Left,Right"
$lv.Font             = New-Object System.Drawing.Font("Consolas", 9.5)
[void]$lv.Columns.Add("Equipo", 230)
[void]$lv.Columns.Add("Extension", 200)
$form.Controls.Add($lv)

# -- Botones inferiores --
$btnAdd              = New-Object System.Windows.Forms.Button
$btnAdd.Text         = "Anadir"
$btnAdd.Location     = New-Object System.Drawing.Point(8, 326)
$btnAdd.Size         = New-Object System.Drawing.Size(70, 28)
$btnAdd.Anchor       = [System.Windows.Forms.AnchorStyles]"Bottom,Left"
$form.Controls.Add($btnAdd)

$btnMod              = New-Object System.Windows.Forms.Button
$btnMod.Text         = "Modificar"
$btnMod.Location     = New-Object System.Drawing.Point(84, 326)
$btnMod.Size         = New-Object System.Drawing.Size(76, 28)
$btnMod.Anchor       = [System.Windows.Forms.AnchorStyles]"Bottom,Left"
$form.Controls.Add($btnMod)

$btnDel              = New-Object System.Windows.Forms.Button
$btnDel.Text         = "Eliminar"
$btnDel.Location     = New-Object System.Drawing.Point(166, 326)
$btnDel.Size         = New-Object System.Drawing.Size(70, 28)
$btnDel.Anchor       = [System.Windows.Forms.AnchorStyles]"Bottom,Left"
$form.Controls.Add($btnDel)

$lblInfo             = New-Object System.Windows.Forms.Label
$lblInfo.Location    = New-Object System.Drawing.Point(248, 330)
$lblInfo.Size        = New-Object System.Drawing.Size(210, 20)
$lblInfo.TextAlign   = "MiddleRight"
$lblInfo.ForeColor   = [System.Drawing.Color]::Gray
$lblInfo.Font        = New-Object System.Drawing.Font("Segoe UI", 8)
$lblInfo.Anchor      = [System.Windows.Forms.AnchorStyles]"Bottom,Right"
$form.Controls.Add($lblInfo)

#=========================
# Funciones UI
#=========================
function Fill-List {
    param([string]$Filter = "")
    $lv.BeginUpdate()
    $lv.Items.Clear()
    $rows = Search-Extensions -Filter $Filter
    foreach ($pair in $rows) {
        $item = New-Object System.Windows.Forms.ListViewItem ([string]$pair[0])
        [void]$item.SubItems.Add([string]$pair[1])
        [void]$lv.Items.Add($item)
    }
    $lv.EndUpdate()
    $n = $rows.Count
    $lblInfo.Text = if ($Filter) { "$n resultado(s)" } else { "$n registros" }
}

function Show-EditDlg {
    param([string]$Title = "Anadir", [string]$Eq = "", [string]$Ex = "", [bool]$Lock = $false)
    $d = New-Object System.Windows.Forms.Form
    $d.Text = $Title; $d.Size = New-Object System.Drawing.Size(340, 160)
    $d.StartPosition = "CenterParent"; $d.FormBorderStyle = "FixedDialog"
    $d.MaximizeBox = $false; $d.MinimizeBox = $false; $d.TopMost = $true
    $d.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $l1 = New-Object System.Windows.Forms.Label; $l1.Text = "Equipo:"
    $l1.Location = New-Object System.Drawing.Point(10, 14); $l1.Size = New-Object System.Drawing.Size(70, 20)
    $d.Controls.Add($l1)
    $t1 = New-Object System.Windows.Forms.TextBox; $t1.Text = $Eq; $t1.CharacterCasing = "Upper"
    $t1.Location = New-Object System.Drawing.Point(85, 12); $t1.Size = New-Object System.Drawing.Size(230, 22)
    if ($Lock) { $t1.ReadOnly = $true; $t1.BackColor = [System.Drawing.Color]::FromArgb(230,230,230) }
    $d.Controls.Add($t1)

    $l2 = New-Object System.Windows.Forms.Label; $l2.Text = "Extension:"
    $l2.Location = New-Object System.Drawing.Point(10, 46); $l2.Size = New-Object System.Drawing.Size(70, 20)
    $d.Controls.Add($l2)
    $t2 = New-Object System.Windows.Forms.TextBox; $t2.Text = $Ex
    $t2.Location = New-Object System.Drawing.Point(85, 44); $t2.Size = New-Object System.Drawing.Size(230, 22)
    $d.Controls.Add($t2)

    $ok = New-Object System.Windows.Forms.Button; $ok.Text = "Aceptar"
    $ok.DialogResult = "OK"; $ok.Location = New-Object System.Drawing.Point(140, 82); $ok.Size = New-Object System.Drawing.Size(80, 28)
    $d.Controls.Add($ok); $d.AcceptButton = $ok

    $ca = New-Object System.Windows.Forms.Button; $ca.Text = "Cancelar"
    $ca.DialogResult = "Cancel"; $ca.Location = New-Object System.Drawing.Point(228, 82); $ca.Size = New-Object System.Drawing.Size(80, 28)
    $d.Controls.Add($ca); $d.CancelButton = $ca

    if ($d.ShowDialog($form) -ne "OK") { return $null }
    $e = $t1.Text.Trim(); $x = $t2.Text.Trim()
    if (-not $e -or -not $x) {
        [System.Windows.Forms.MessageBox]::Show("Ambos campos son obligatorios.","Aviso","OK","Warning")
        return $null
    }
    return @{ Equipo = $e; Extension = $x }
}

#=========================
# Eventos
#=========================
$btnBuscar.Add_Click({ Fill-List -Filter $txtSearch.Text.Trim() })
$txtSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $btnBuscar.PerformClick(); $_.SuppressKeyPress = $true } })
$btnTodos.Add_Click({ $txtSearch.Text = ""; Fill-List })

$btnAdd.Add_Click({
    $r = Show-EditDlg -Title "Anadir extension"
    if (-not $r) { return }
    Save-Extension -Equipo $r.Equipo -Extension $r.Extension
    Fill-List -Filter $txtSearch.Text.Trim()
    # Preguntar si replicar al servidor
    if ((Test-ExtFinderServer)) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "¿Actualizar la base de datos en el servidor?",
            "Replicar al servidor", "YesNo", "Question")
        if ($ans -eq 'Yes') { Sync-DbToServer | Out-Null }
    }
})

$btnMod.Add_Click({
    if ($lv.SelectedItems.Count -eq 0) { return }
    $sel = $lv.SelectedItems[0]
    $r = Show-EditDlg -Title "Modificar" -Eq $sel.Text -Ex $sel.SubItems[1].Text -Lock $true
    if (-not $r) { return }
    Save-Extension -Equipo $r.Equipo -Extension $r.Extension
    Fill-List -Filter $txtSearch.Text.Trim()
    # Preguntar si replicar al servidor
    if ((Test-ExtFinderServer)) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "¿Actualizar la base de datos en el servidor?",
            "Replicar al servidor", "YesNo", "Question")
        if ($ans -eq 'Yes') { Sync-DbToServer | Out-Null }
    }
})

$btnDel.Add_Click({
    if ($lv.SelectedItems.Count -eq 0) { return }
    $sel = $lv.SelectedItems[0]
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Eliminar '$($sel.SubItems[1].Text)' del equipo '$($sel.Text)'?",
        "Confirmar", "YesNo", "Warning")
    if ($ans -ne "Yes") { return }
    Remove-Extension -Equipo $sel.Text
    Fill-List -Filter $txtSearch.Text.Trim()
    # Preguntar si replicar al servidor
    if ((Test-ExtFinderServer)) {
        $ans2 = [System.Windows.Forms.MessageBox]::Show(
            "¿Actualizar la base de datos en el servidor?",
            "Replicar al servidor", "YesNo", "Question")
        if ($ans2 -eq 'Yes') { Sync-DbToServer | Out-Null }
    }
})

$lv.Add_DoubleClick({ $btnMod.PerformClick() })
$lv.Add_KeyDown({
    if ($_.KeyCode -eq 'Delete') { $btnDel.PerformClick() }
    if ($_.KeyCode -eq 'F2')     { $btnMod.PerformClick() }
})

$form.Add_FormClosed({
    try { if ($script:Conn) { $script:Conn.Close(); $script:Conn.Dispose() } } catch {}
})

#=========================
# Arranque
#=========================
# Sincronizar desde servidor si hay versión más nueva
Sync-DbFromServer

try {
    Open-DB | Out-Null
    $lblInfo.Text = "$(Get-ExtCount) registros en BD"
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error abriendo BD:`n$($_.Exception.Message)","Error","OK","Error")
    return
}

$form.TopMost = $true
$form.Add_Shown({ $form.Activate(); $txtSearch.Focus() })
[void]$form.ShowDialog()
