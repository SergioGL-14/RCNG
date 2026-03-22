function Invoke-LocalHost_netstatsListening {
    # Mientras se recoge netstat dejamos la opción deshabilitada para evitar dobles clics.
    if ($null -ne $global:ToolStripMenuItem_netstatsListening) {
        $global:ToolStripMenuItem_netstatsListening.Enabled = $false
    }
    Add-Logs -text "$env:ComputerName - Netstat"

    $tempFile = [System.IO.Path]::GetTempFileName()

    try {
        netstat -ano | Out-File -FilePath $tempFile -Encoding UTF8
        $contenido = Get-Content $tempFile -Raw
        Add-RichTextBox -text $contenido
    }
    catch {
        Add-RichTextBox -text "Error al ejecutar netstat: $_"
    }
    finally {
        # Limpiar el temporal aunque la lectura haya fallado.
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
        # Se reabre el menú unos segundos después. El timer queda en scope
        # script para que la referencia siga viva al salir de la función.
        $script:NetstatTimer = New-Object System.Windows.Forms.Timer
        $script:NetstatTimer.Interval = 3000

        $script:NetstatTimer.Add_Tick({
            if ($null -ne $global:ToolStripMenuItem_netstatsListening) {
                $global:ToolStripMenuItem_netstatsListening.Enabled = $true
            }
            # Al terminar, el timer se limpia a sí mismo y libera la referencia.
            $script:NetstatTimer.Stop()
            $script:NetstatTimer.Dispose()
            Remove-Variable -Name NetstatTimer -Scope Script -ErrorAction SilentlyContinue
        })
        $script:NetstatTimer.Start()
    }
}

function Invoke-LocalHost_registeredSnappins {
    $global:ToolStripMenuItem_registeredSnappins.Enabled = $false
    Add-Logs -text "Snap-ins & Módulos Disponibles"

    try {
        $snapins = Get-PSSnapin -Registered | Out-String
        $snapinsTexto = $snapins.Trim()

        if (![string]::IsNullOrWhiteSpace($snapinsTexto)) {
            Add-RichTextBox -text "Snap-ins registrados encontrados:`r`n$snapinsTexto"
        }
        else {
            $modulos = Get-Module -ListAvailable | Select-Object Name, Version, Path | Out-String
            $modulosTexto = $modulos.Trim()

            if (![string]::IsNullOrWhiteSpace($modulosTexto)) {
                Add-RichTextBox -text "No hay Snap-ins registrados.`r`nMostrando módulos disponibles:`r`n$modulosTexto"
            }
            else {
                Add-RichTextBox -text "No se encontraron ni Snap-ins ni módulos disponibles."
            }
        }
    }
    catch {
        Add-RichTextBox -text "Error al obtener Snap-ins o módulos: $_"
    }
    finally {
        $global:ToolStripMenuItem_registeredSnappins.Enabled = $true
    }
}


function Invoke-LocalHost_resetCredenciaisVNC {
    # Recrea el diálogo simple para guardar la contraseña temporal de VNC.
    $ButtonGuardar = [Windows.Forms.Button]@{
        Text = 'Guardar'; Location = [Drawing.Point]::new(130, 120); Width = 100; Height = 50
    }
    $ButtonGuardar.Add_Click({
        $contrasenaSegura = ConvertTo-SecureString $TextBoxContrasena.text -AsPlainText -Force
        $env:ContrasenaGuardada = ConvertFrom-SecureString $contrasenaSegura
        $formularioContrasena.Close()
    })

    $TextBoxContrasena = [Windows.Forms.TextBox]@{
        Location = [Drawing.Point]::new(30, 80); Width = 320; PasswordChar = "*"
    }
    $LblintroduceCont = [Windows.Forms.Label]@{
        Text = 'Introduce una contraseña:'; Location = [Drawing.Point]::new(30, 30); AutoSize = $false; Width = 300; Height = 20
    }
    $formularioContrasena = New-Object Windows.Forms.Form
    $formularioContrasena.Text = "Conexion"
    $formularioContrasena.Width = 400
    $formularioContrasena.Height = 250
    $formularioContrasena.Controls.AddRange(@($LblintroduceCont, $TextBoxContrasena, $ButtonGuardar))
    $formularioContrasena.ShowDialog()
}

function Invoke-LocalHost_systemInformationMSinfo32exe {
    Start-Process msinfo32.exe
}

function Invoke-LocalHost_systemproperties {
    Start-Process "sysdm.cpl"
}

function Invoke-LocalHost_devicemanager {
    Start-Process "devmgmt.msc"
}



function Invoke-LocalHost_certificateManager {
    Start-Process certmgr.msc
}

function Invoke-LocalHost_sharedFolders {
    Start-Process "fsmgmt.msc"
}

function Invoke-LocalHost_performanceMonitor {
    Start-Process "Perfmon.msc"
}


function Invoke-LocalHost_groupPolicyEditor {
    Start-Process "Gpedit.msc"
}

function Invoke-LocalHost_localUsersAndGroups {
    Start-Process "lusrmgr.msc"
}

function Invoke-LocalHost_diskManagement {
    Start-Process "diskmgmt.msc"
}

function Invoke-LocalHost_localSecuritySettings {
    Start-Process "secpol.msc"
}


function Invoke-LocalHost_scheduledTasks {
    Start-Process "control" -ArgumentList "schedtasks"
}

function Invoke-LocalHost_PowershellISE {
    Start-Process powershell_ise.exe
}

Export-ModuleMember -Function *


