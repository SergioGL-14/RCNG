# Seccion de Configuracion.
# Reune el mantenimiento local de equipos y la configuracion del entorno
# para desacoplar rutas UNC, proxy y portales del codigo principal.
param()

#==================================================================
# BLOQUE: Helpers privados de conexion
#==================================================================

function script:Open-CompDBEdit {
    $dll = Join-Path $Global:ScriptRoot 'libs\System.Data.SQLite.dll'
    $db  = Join-Path $Global:ScriptRoot 'database\ComputerNames.sqlite'
    Add-Type -Path $dll -ErrorAction SilentlyContinue
    $c = New-Object System.Data.SQLite.SQLiteConnection "Data Source=$db;Version=3;"
    $c.Open()
    return $c
}

#==================================================================
# BLOQUE: Helpers de interfaz
#==================================================================

function script:New-CfgButton {
    param(
        [string]$Text,
        [int]$X, [int]$Y,
        [int]$W = 90, [int]$H = 26,
        [string]$Color = 'default'
    )

    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $Text
    $b.Location  = New-Object System.Drawing.Point($X, $Y)
    $b.Size      = New-Object System.Drawing.Size($W, $H)
    $b.FlatStyle = 'Flat'
    $b.Font      = New-Object System.Drawing.Font('Segoe UI', 9)

    switch ($Color) {
        'blue'  { $b.BackColor = [System.Drawing.Color]::FromArgb(0,120,215);   $b.ForeColor = 'White' }
        'red'   { $b.BackColor = [System.Drawing.Color]::FromArgb(196,43,28);   $b.ForeColor = 'White' }
        'green' { $b.BackColor = [System.Drawing.Color]::FromArgb(16,124,16);   $b.ForeColor = 'White' }
        'gray'  { $b.BackColor = [System.Drawing.Color]::FromArgb(100,100,100); $b.ForeColor = 'White' }
    }

    return $b
}

function script:Show-CfgMessage {
    param(
        [string]$Message,
        [string]$Title = 'Configuracion',
        [string]$Icon = 'Information'
    )

    [System.Windows.Forms.MessageBox]::Show($Message, $Title, 'OK', $Icon) | Out-Null
}

Set-StrictMode -Off

#==================================================================
# BLOQUE: Dialogo Gestion de Equipos
#==================================================================

function Show-EquiposDialog {
    <#
    .SYNOPSIS
        CRUD + paginacion + importacion CSV sobre la tabla computers.
    #>

    $pageSize = 500
    $state    = @{ Page = 0; Filter = '' }

    $form = New-Object System.Windows.Forms.Form
    $form.Text          = 'Gestion de Equipos'
    $form.Size          = New-Object System.Drawing.Size(880, 630)
    $form.StartPosition = 'CenterParent'
    $form.MinimumSize   = New-Object System.Drawing.Size(720, 480)
    $form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

    try {
        $conn = script:Open-CompDBEdit
    } catch {
        script:Show-CfgMessage -Message "No se pudo abrir la base de datos de equipos.`n$($_.Exception.Message)" -Title 'Error' -Icon 'Error'
        return
    }

    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add('id',     [int])
    [void]$dt.Columns.Add('ou',     [string])
    [void]$dt.Columns.Add('equipo', [string])
    $dt.Columns['id'].ReadOnly = $true

    $panelTop        = New-Object System.Windows.Forms.Panel
    $panelTop.Dock   = 'Top'
    $panelTop.Height = 52

    $lblSearch          = New-Object System.Windows.Forms.Label
    $lblSearch.Text     = 'Buscar:'
    $lblSearch.AutoSize = $true
    $lblSearch.Location = '8,15'

    $txtSearch          = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = '60,12'
    $txtSearch.Width    = 210

    $btnSearch    = script:New-CfgButton 'Buscar'        278 11 70
    $btnClear     = script:New-CfgButton 'Todo'          354 11 55
    $btnAdd       = script:New-CfgButton '+ Anadir'      418 11 85  26 'blue'
    $btnDelete    = script:New-CfgButton 'Eliminar'      510 11 80  26 'red'
    $btnImportCSV = script:New-CfgButton 'Importar CSV'  598 11 110 26 'gray'
    $panelTop.Controls.AddRange(@($lblSearch,$txtSearch,$btnSearch,$btnClear,$btnAdd,$btnDelete,$btnImportCSV))

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock                  = 'Fill'
    $grid.AllowUserToAddRows    = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode   = 'Fill'
    $grid.SelectionMode         = 'FullRowSelect'
    $grid.MultiSelect           = $false
    $grid.RowHeadersVisible     = $false
    $grid.BackgroundColor       = 'White'
    $grid.BorderStyle           = 'None'
    $grid.GridColor             = [System.Drawing.Color]::FromArgb(220,220,220)
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248,248,248)
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersHeightSizeMode = 'AutoSize'
    $grid.DataSource = $dt

    $panelPage           = New-Object System.Windows.Forms.Panel
    $panelPage.Dock      = 'Bottom'
    $panelPage.Height    = 34
    $panelPage.BackColor = [System.Drawing.Color]::FromArgb(235,235,235)
    $btnPrev = script:New-CfgButton '< Anterior'   0 0 100 34 'gray'
    $btnNext = script:New-CfgButton 'Siguiente >'  0 0 100 34 'gray'
    $btnPrev.Dock = 'Left'
    $btnNext.Dock = 'Right'
    $btnPrev.Enabled = $false
    $btnNext.Enabled = $false

    $lblPage           = New-Object System.Windows.Forms.Label
    $lblPage.Dock      = 'Fill'
    $lblPage.TextAlign = 'MiddleCenter'
    $lblPage.Font      = New-Object System.Drawing.Font('Segoe UI', 9)

    $panelPage.Controls.Add($lblPage)
    $panelPage.Controls.Add($btnPrev)
    $panelPage.Controls.Add($btnNext)

    $panelBottom           = New-Object System.Windows.Forms.Panel
    $panelBottom.Dock      = 'Bottom'
    $panelBottom.Height    = 45
    $panelBottom.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)

    $btnSave = script:New-CfgButton 'Guardar cambios' 10 10 130 26 'green'
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.AutoSize  = $true
    $lblStatus.Location  = '150,15'
    $lblStatus.ForeColor = 'Green'

    $pClose       = New-Object System.Windows.Forms.Panel
    $pClose.Dock  = 'Right'
    $pClose.Width = 90
    $btnClose     = script:New-CfgButton 'Cerrar' 5 10 80 26 'gray'
    $pClose.Controls.Add($btnClose)
    $panelBottom.Controls.AddRange(@($btnSave, $lblStatus, $pClose))

    $form.Controls.AddRange(@($grid, $panelPage, $panelBottom, $panelTop))

    function GetTotalEquipos {
        $sql = 'SELECT COUNT(*) FROM computers'
        if ($state.Filter) {
            $sql += ' WHERE UPPER(equipo) LIKE UPPER(@f) OR UPPER(ou) LIKE UPPER(@f)'
        }
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $sql
        if ($state.Filter) {
            [void]$cmd.Parameters.AddWithValue('@f', "%$($state.Filter)%")
        }
        return [int]$cmd.ExecuteScalar()
    }

    function LoadEquipos {
        try {
            $dt.Rows.Clear()
            $offset = $state.Page * $pageSize
            $sql = 'SELECT id, ou, equipo FROM computers'
            if ($state.Filter) {
                $sql += ' WHERE UPPER(equipo) LIKE UPPER(@f) OR UPPER(ou) LIKE UPPER(@f)'
            }
            $sql += " ORDER BY ou, equipo LIMIT $pageSize OFFSET $offset"

            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $sql
            if ($state.Filter) {
                [void]$cmd.Parameters.AddWithValue('@f', "%$($state.Filter)%")
            }

            $reader = $cmd.ExecuteReader()
            while ($reader.Read()) {
                $row = $dt.NewRow()
                $row['id']     = [int]$reader['id']
                $row['ou']     = [string]$reader['ou']
                $row['equipo'] = [string]$reader['equipo']
                [void]$dt.Rows.Add($row)
            }
            $reader.Close()
            $dt.AcceptChanges()

            if ($grid.Columns.Contains('id'))     { $grid.Columns['id'].Visible = $false }
            if ($grid.Columns.Contains('ou'))     { $grid.Columns['ou'].HeaderText = 'OU / Grupo';    $grid.Columns['ou'].FillWeight = 40 }
            if ($grid.Columns.Contains('equipo')) { $grid.Columns['equipo'].HeaderText = 'Nombre Equipo'; $grid.Columns['equipo'].FillWeight = 60 }

            $total      = GetTotalEquipos
            $totalPages = [Math]::Max(1, [Math]::Ceiling($total / $pageSize))
            $lblPage.Text    = "Pagina $($state.Page + 1) de $totalPages   ($total equipos)"
            $btnPrev.Enabled = $state.Page -gt 0
            $btnNext.Enabled = $state.Page -lt ($totalPages - 1)
            $lblStatus.Text  = ''
        } catch {
            $lblStatus.Text      = "Error cargando: $($_.Exception.Message)"
            $lblStatus.ForeColor = 'Red'
        }
    }

    function SaveEquipos {
        $changes = $dt.GetChanges()
        if ($null -eq $changes) {
            $lblStatus.Text      = 'Sin cambios pendientes.'
            $lblStatus.ForeColor = 'Gray'
            return
        }

        $tx = $conn.BeginTransaction()
        $ins = 0
        $upd = 0
        $del = 0

        try {
            foreach ($row in $changes.Rows) {
                switch ($row.RowState) {
                    'Added' {
                        $c = $conn.CreateCommand()
                        $c.Transaction = $tx
                        $c.CommandText = "INSERT INTO computers (ou, equipo, orig_line) VALUES (@ou,@eq,@ol)"
                        [void]$c.Parameters.AddWithValue('@ou', [string]$row['ou'])
                        [void]$c.Parameters.AddWithValue('@eq', [string]$row['equipo'])
                        [void]$c.Parameters.AddWithValue('@ol', "$($row['ou'])\$($row['equipo'])")
                        $c.ExecuteNonQuery() | Out-Null
                        $ins++
                    }
                    'Modified' {
                        $c = $conn.CreateCommand()
                        $c.Transaction = $tx
                        $c.CommandText = "UPDATE computers SET ou=@ou, equipo=@eq, orig_line=@ol WHERE id=@id"
                        [void]$c.Parameters.AddWithValue('@ou', [string]$row['ou'])
                        [void]$c.Parameters.AddWithValue('@eq', [string]$row['equipo'])
                        [void]$c.Parameters.AddWithValue('@ol', "$($row['ou'])\$($row['equipo'])")
                        [void]$c.Parameters.AddWithValue('@id', $row.Item('id', [System.Data.DataRowVersion]::Original))
                        $c.ExecuteNonQuery() | Out-Null
                        $upd++
                    }
                    'Deleted' {
                        $origId = [int]$row.Item('id', [System.Data.DataRowVersion]::Original)
                        if ($origId -gt 0) {
                            $c = $conn.CreateCommand()
                            $c.Transaction = $tx
                            $c.CommandText = "DELETE FROM computers WHERE id=@id"
                            [void]$c.Parameters.AddWithValue('@id', $origId)
                            $c.ExecuteNonQuery() | Out-Null
                            $del++
                        }
                    }
                }
            }

            $tx.Commit()
            $dt.AcceptChanges()
            $lblStatus.Text      = "Guardado: +$ins  editados:$upd  eliminados:$del"
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(16,124,16)

            if (Get-Command 'Sync-ComputerDbToServer' -ErrorAction SilentlyContinue) {
                $ans = [System.Windows.Forms.MessageBox]::Show(
                    "Desea replicar los cambios de equipos al recurso compartido?",
                    'Replicar cambios',
                    'YesNo',
                    'Question'
                )
                if ($ans -eq 'Yes') {
                    Sync-ComputerDbToServer | Out-Null
                }
            }
        } catch {
            try { $tx.Rollback() } catch {}
            $lblStatus.Text      = "Error: $($_.Exception.Message)"
            $lblStatus.ForeColor = 'Red'
        }
    }

    function ShowAddEquipo {
        $af = New-Object System.Windows.Forms.Form
        $af.Text            = 'Anadir Equipo'
        $af.Size            = New-Object System.Drawing.Size(380, 175)
        $af.StartPosition   = 'CenterParent'
        $af.FormBorderStyle = 'FixedDialog'
        $af.MaximizeBox     = $false
        $af.MinimizeBox     = $false
        $af.Font            = New-Object System.Drawing.Font('Segoe UI', 9)

        $l1 = New-Object System.Windows.Forms.Label
        $l2 = New-Object System.Windows.Forms.Label
        $t1 = New-Object System.Windows.Forms.TextBox
        $t2 = New-Object System.Windows.Forms.TextBox

        $l1.Text = 'OU / Grupo:'
        $l1.AutoSize = $true
        $l1.Location = '10,22'
        $l2.Text = 'Nombre equipo:'
        $l2.AutoSize = $true
        $l2.Location = '10,57'
        $t1.Location = '115,19'
        $t1.Width    = 240
        $t2.Location = '115,54'
        $t2.Width    = 240

        $bOK     = script:New-CfgButton 'Anadir'   115 90 90 26 'blue'
        $bCancel = script:New-CfgButton 'Cancelar' 212 90 90 26
        $bOK.DialogResult = 'OK'
        $bCancel.DialogResult = 'Cancel'

        $af.AcceptButton = $bOK
        $af.CancelButton = $bCancel
        $af.Controls.AddRange(@($l1,$l2,$t1,$t2,$bOK,$bCancel))

        if ($af.ShowDialog($form) -eq 'OK') {
            $ou = $t1.Text.Trim()
            $eq = $t2.Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($eq)) {
                $row = $dt.NewRow()
                $row['id'] = 0
                $row['ou'] = $ou
                $row['equipo'] = $eq
                [void]$dt.Rows.Add($row)
                $grid.FirstDisplayedScrollingRowIndex = $grid.RowCount - 1
                $lblStatus.Text      = "Fila anadida. Pulsa 'Guardar cambios' para confirmar."
                $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0,100,200)
            }
        }

        $af.Dispose()
    }

    function ImportCSVEquipos {
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title  = 'Importar Equipos desde CSV'
        $ofd.Filter = 'CSV files (*.csv)|*.csv|Todos los archivos (*.*)|*.*'
        if ($ofd.ShowDialog($form) -ne 'OK') {
            $ofd.Dispose()
            return
        }

        $csvPath = $ofd.FileName
        $ofd.Dispose()

        $firstLine = Get-Content $csvPath -First 1
        $delim = if ($firstLine -match ';') { ';' } else { ',' }

        try {
            $csvData = Import-Csv -Path $csvPath -Delimiter $delim -Encoding UTF8
        } catch {
            script:Show-CfgMessage -Message "Error leyendo CSV.`n$($_.Exception.Message)" -Title 'Error' -Icon 'Error'
            return
        }

        if (-not $csvData -or $csvData.Count -eq 0) {
            script:Show-CfgMessage -Message 'El CSV esta vacio.' -Title 'Aviso'
            return
        }

        $cols  = $csvData[0].PSObject.Properties.Name
        $colEq = $cols | Where-Object { $_ -imatch '^equipo$' } | Select-Object -First 1
        $colOu = $cols | Where-Object { $_ -imatch '^ou$' } | Select-Object -First 1
        if (-not $colEq -and $cols.Count -ge 1) { $colEq = $cols[0] }

        if (-not $colEq) {
            script:Show-CfgMessage -Message "No se encontro una columna 'equipo'.`nColumnas detectadas: $($cols -join ', ')" -Title 'Formato incorrecto' -Icon 'Warning'
            return
        }

        $ouInfo = if ($colOu) { "'$colOu'" } else { '(no encontrada, quedara vacia)' }
        $choice = [System.Windows.Forms.MessageBox]::Show(
            "CSV: $($csvData.Count) filas.`nColumna equipo: '$colEq'`nColumna ou: $ouInfo`n`nSi = Reemplazar todo   No = Anadir   Cancelar = Abortar",
            'Importar CSV - Equipos',
            'YesNoCancel',
            'Question'
        )
        if ($choice -eq 'Cancel') { return }

        $tx = $conn.BeginTransaction()
        try {
            if ($choice -eq 'Yes') {
                $c = $conn.CreateCommand()
                $c.Transaction = $tx
                $c.CommandText = 'DELETE FROM computers'
                $c.ExecuteNonQuery() | Out-Null
            }

            $count = 0
            foreach ($fila in $csvData) {
                $eq = ($fila.$colEq -replace '^\s+|\s+$')
                $ou = if ($colOu) { ($fila.$colOu -replace '^\s+|\s+$') } else { '' }
                if ([string]::IsNullOrWhiteSpace($eq)) { continue }

                $ic = $conn.CreateCommand()
                $ic.Transaction = $tx
                $ic.CommandText = 'INSERT INTO computers (ou, equipo, orig_line) VALUES (@ou,@eq,@ol)'
                [void]$ic.Parameters.AddWithValue('@ou', $ou)
                [void]$ic.Parameters.AddWithValue('@eq', $eq)
                [void]$ic.Parameters.AddWithValue('@ol', $(if ($ou) { "$ou\$eq" } else { $eq }))
                $ic.ExecuteNonQuery() | Out-Null
                $count++
            }

            $tx.Commit()
            $state.Page = 0
            $state.Filter = ''
            $txtSearch.Text = ''
            LoadEquipos
            $lblStatus.Text      = "CSV importado: $count equipos."
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(16,124,16)
        } catch {
            try { $tx.Rollback() } catch {}
            $lblStatus.Text      = "Error importando: $($_.Exception.Message)"
            $lblStatus.ForeColor = 'Red'
        }
    }

    $btnSearch.Add_Click({ $state.Page = 0; $state.Filter = $txtSearch.Text.Trim(); LoadEquipos })
    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq 'Return') {
            $state.Page = 0
            $state.Filter = $txtSearch.Text.Trim()
            LoadEquipos
        }
    })
    $btnClear.Add_Click({ $txtSearch.Text = ''; $state.Page = 0; $state.Filter = ''; LoadEquipos })
    $btnPrev.Add_Click({ if ($state.Page -gt 0) { $state.Page--; LoadEquipos } })
    $btnNext.Add_Click({ $state.Page++; LoadEquipos })
    $btnAdd.Add_Click({ ShowAddEquipo })
    $btnDelete.Add_Click({
        if ($null -ne $grid.CurrentRow -and -not $grid.CurrentRow.IsNewRow) {
            $drv = $grid.CurrentRow.DataBoundItem -as [System.Data.DataRowView]
            if ($null -ne $drv) {
                $name = $drv.Row['equipo']
                if ([System.Windows.Forms.MessageBox]::Show(
                    "Eliminar el equipo '$name'?",
                    'Confirmar eliminacion',
                    'YesNo',
                    'Warning'
                ) -eq 'Yes') {
                    $drv.Row.Delete()
                    $lblStatus.Text      = "Equipo marcado para eliminar. Pulsa 'Guardar cambios'."
                    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(180,60,0)
                }
            }
        }
    })
    $btnImportCSV.Add_Click({ ImportCSVEquipos })
    $btnSave.Add_Click({ SaveEquipos })
    $btnClose.Add_Click({
        if ($null -ne $dt.GetChanges()) {
            if ([System.Windows.Forms.MessageBox]::Show(
                'Hay cambios sin guardar. Cerrar de todos modos?',
                'Cambios pendientes',
                'YesNo',
                'Warning'
            ) -ne 'Yes') {
                return
            }
        }
        $form.Close()
    })

    $form.Add_FormClosed({
        try { $conn.Close(); $conn.Dispose() } catch {}
        try { Initialize-ComputerDB | Out-Null } catch {}
    })
    $form.Add_Shown({ LoadEquipos })

    [void]$form.ShowDialog([System.Windows.Forms.Application]::OpenForms[0])
    $form.Dispose()
}

#==================================================================
# BLOQUE: Dialogo de Entorno
#==================================================================

function Show-EnvironmentDialog {
    if (-not (Get-Command 'Get-AppSettings' -ErrorAction SilentlyContinue) -or
        -not (Get-Command 'Save-AppSettings' -ErrorAction SilentlyContinue)) {
        script:Show-CfgMessage -Message 'La configuracion de entorno no esta disponible porque SharedDataManager no esta cargado.' -Title 'Error' -Icon 'Error'
        return
    }

    $settings = Get-AppSettings

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Configuracion de Entorno'
    $form.Size            = New-Object System.Drawing.Size(860, 760)
    $form.StartPosition   = 'CenterParent'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.Font            = New-Object System.Drawing.Font('Segoe UI', 9)

    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = 'Top'
    $headerPanel.Height = 72
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = 'Entorno global'
    $titleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $titleLabel.AutoSize = $true
    $titleLabel.Location = New-Object System.Drawing.Point(18, 12)

    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = 'Configura aqui los valores de ejemplo que dependen de cada implantacion: repositorio compartido, proxy, portales, WOL, DHCP y datos visibles de soporte.'
    $subtitleLabel.MaximumSize = New-Object System.Drawing.Size(790, 0)
    $subtitleLabel.AutoSize = $true
    $subtitleLabel.ForeColor = [System.Drawing.Color]::DimGray
    $subtitleLabel.Location = New-Object System.Drawing.Point(18, 36)

    $headerPanel.Controls.AddRange(@($titleLabel, $subtitleLabel))

    $footerPanel = New-Object System.Windows.Forms.Panel
    $footerPanel.Dock = 'Bottom'
    $footerPanel.Height = 66
    $footerPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)

    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Dock = 'Fill'
    $contentPanel.AutoScroll = $true
    $contentPanel.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 10)

    $labels = @(
        @{ Key='SharedServerBase';      Text='Ruta UNC del repositorio compartido';       Example='Ejemplo: \\server\share\NRC_APP' },
        @{ Key='ProxyPacUrl';           Text='URL del archivo PAC';                        Example='Ejemplo: http://proxy.example.local/proxy.pac' },
        @{ Key='PortalUrl';             Text='Portal web principal';                       Example='Ejemplo: https://portal.example.local' },
        @{ Key='MailPortalUrl';         Text='Portal web secundario';                      Example='Ejemplo: https://mail.example.local' },
        @{ Key='WolCsvShare';           Text='Share UNC para inventario WOL';              Example='Ejemplo: \\server\share\network' },
        @{ Key='WolCsvFileName';        Text='Nombre del CSV usado por WOL';               Example='Ejemplo: network_inventory_sample.csv' },
        @{ Key='DhcpServer';            Text='Servidor DHCP por defecto';                  Example='Ejemplo: DHCP-SERVER' },
        @{ Key='SupportDisplayName';    Text='Nombre visible de soporte';                  Example='Ejemplo: NRC_APP Support' },
        @{ Key='SupportEmail';          Text='Correo visible de soporte';                  Example='Ejemplo: support@example.local' },
        @{ Key='PrimaryGroupSearchBase';   Text='Base LDAP principal para grupos';         Example='Ejemplo: OU=GrupoPrincipal,DC=example,DC=local' },
        @{ Key='SecondaryGroupSearchBase'; Text='Base LDAP secundaria para grupos';        Example='Ejemplo: OU=GrupoSecundario,DC=example,DC=local' }
    )

    $textBoxes = @{}
    $y = 12
    foreach ($item in $labels) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $item.Text
        $lbl.AutoSize = $true
        $lbl.Location = New-Object System.Drawing.Point(18, $y)
        $contentPanel.Controls.Add($lbl)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point(18, ($y + 20))
        $txt.Size     = New-Object System.Drawing.Size(792, 24)
        $txt.Text     = [string]$settings[$item.Key]
        $txt.Anchor   = 'Top, Left, Right'
        $contentPanel.Controls.Add($txt)
        $textBoxes[$item.Key] = $txt

        $lblExample = New-Object System.Windows.Forms.Label
        $lblExample.Text = $item.Example
        $lblExample.AutoSize = $true
        $lblExample.ForeColor = [System.Drawing.Color]::DimGray
        $lblExample.Location = New-Object System.Drawing.Point(18, ($y + 49))
        $contentPanel.Controls.Add($lblExample)

        $y += 74
    }

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = 'Sugerencia: usa valores de ejemplo durante la implantacion inicial y sustituyelos por los datos reales antes de distribuir el launcher.'
    $lblInfo.AutoSize = $true
    $lblInfo.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $lblInfo.MaximumSize = New-Object System.Drawing.Size(792, 0)
    $lblInfo.Location = New-Object System.Drawing.Point(18, $y)
    $contentPanel.Controls.Add($lblInfo)

    $btnDefaults = script:New-CfgButton 'Restaurar ejemplos' 402 18 150 30 'gray'
    $btnSave     = script:New-CfgButton 'Guardar'            562 18 120 30 'green'
    $btnCancel   = script:New-CfgButton 'Cerrar'             690 18 120 30
    $btnDefaults.Anchor = 'Top, Right'
    $btnSave.Anchor     = 'Top, Right'
    $btnCancel.Anchor   = 'Top, Right'
    $footerPanel.Controls.AddRange(@($btnDefaults, $btnSave, $btnCancel))

    $form.Controls.Add($contentPanel)
    $form.Controls.Add($footerPanel)
    $form.Controls.Add($headerPanel)

    $btnDefaults.Add_Click({
        $defaults = Get-DefaultAppSettings
        foreach ($key in $defaults.Keys) {
            if ($textBoxes.ContainsKey($key)) {
                $textBoxes[$key].Text = [string]$defaults[$key]
            }
        }
    })

    $btnSave.Add_Click({
        $payload = @{
            SharedServerBase      = $textBoxes['SharedServerBase'].Text.Trim()
            ProxyPacUrl           = $textBoxes['ProxyPacUrl'].Text.Trim()
            PortalUrl             = $textBoxes['PortalUrl'].Text.Trim()
            MailPortalUrl         = $textBoxes['MailPortalUrl'].Text.Trim()
            WolCsvShare           = $textBoxes['WolCsvShare'].Text.Trim()
            WolCsvFileName        = $textBoxes['WolCsvFileName'].Text.Trim()
            DhcpServer            = $textBoxes['DhcpServer'].Text.Trim()
            SupportDisplayName    = $textBoxes['SupportDisplayName'].Text.Trim()
            SupportEmail          = $textBoxes['SupportEmail'].Text.Trim()
            PrimaryGroupSearchBase   = $textBoxes['PrimaryGroupSearchBase'].Text.Trim()
            SecondaryGroupSearchBase = $textBoxes['SecondaryGroupSearchBase'].Text.Trim()
        }

        Save-AppSettings -Settings $payload | Out-Null
        script:Show-CfgMessage -Message 'Configuracion guardada correctamente.'
    })

    $btnCancel.Add_Click({ $form.Close() })

    [void]$form.ShowDialog([System.Windows.Forms.Application]::OpenForms[0])
    $form.Dispose()
}

#==================================================================
# BLOQUE: Inicializacion del menu Configuracion
#==================================================================

function Initialize-ConfiguracionMenu {
    param([System.Windows.Forms.ToolStripMenuItem]$Menu)

    $Menu.DropDownItems.Clear()

    $itemEquipos = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemEquipos.Text = 'Equipos'
    $itemEquipos.ToolTipText = 'Gestionar la base de datos local de equipos'
    $itemEquipos.Add_Click({ Show-EquiposDialog })

    $itemEntorno = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemEntorno.Text = 'Entorno global'
    $itemEntorno.ToolTipText = 'Configurar rutas UNC, proxy, WOL, DHCP y parametros visibles del entorno'
    $itemEntorno.Add_Click({ Show-EnvironmentDialog })

    $sep = New-Object System.Windows.Forms.ToolStripSeparator

    $itemPK = New-Object System.Windows.Forms.ToolStripMenuItem
    $itemPK.Text = 'Limpiar Pass Keeper'
    $itemPK.ToolTipText = 'Eliminar todas las entradas guardadas en Pass Keeper'
    $itemPK.ForeColor = [System.Drawing.Color]::FromArgb(160, 40, 40)
    $itemPK.Add_Click({
        if (Get-Command 'Clear-PassKeeperData' -ErrorAction SilentlyContinue) {
            Clear-PassKeeperData
        }
    })

    [void]$Menu.DropDownItems.Add($itemEquipos)
    [void]$Menu.DropDownItems.Add($itemEntorno)
    [void]$Menu.DropDownItems.Add($sep)
    [void]$Menu.DropDownItems.Add($itemPK)
}

Export-ModuleMember -Function Initialize-ConfiguracionMenu, Show-EquiposDialog, Show-EnvironmentDialog
