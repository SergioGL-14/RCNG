# Lado soporte del chat remoto.
# Esta ventana trabaja contra el chat.txt compartido por UNC y se encarga de
# cerrar también la sesión remota si el operador lo confirma.

# Forzar STA para que WinForms y el portapapeles funcionen sin sorpresas.
[System.Threading.Thread]::CurrentThread.ApartmentState = [System.Threading.ApartmentState]::STA

# Cargar las librerías que usa la ventana.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configuración básica de refresco.
$Intervalo = 250  # milisegundos

# Identidad y título de la ventana del lado soporte.
$NombreEquipo = "SOPORTE"
$TituloVentana = "Chat - $NombreEquipo"
# La ruta del chat llega desde la app principal y apunta al archivo compartido remoto.
# Ejemplo de ruta: \\EQUIPO\C$\Users\Usuario\AppData\Local\Temp\chat.txt
if (-not $RutaArchivo) {
    [System.Windows.Forms.MessageBox]::Show(
        "Error: No se especificó la ruta del archivo de chat remoto.`n`nUse: `$RutaArchivo = '\\EQUIPO\...\chat.txt'; & 'chat.ps1'",
        "Error de Configuración",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# Crear ventana
$Form = New-Object System.Windows.Forms.Form
$Form.Text = $TituloVentana
$Form.Size = New-Object System.Drawing.Size(420, 500)
$Form.StartPosition = "CenterScreen"
$Form.Topmost = $true
$Form.BackColor = [System.Drawing.Color]::White

# Área de mensajes (RichTextBox para colores)
$ChatBox = New-Object System.Windows.Forms.RichTextBox
$ChatBox.Multiline = $true
$ChatBox.Location = New-Object System.Drawing.Point(10, 10)
$ChatBox.Size = New-Object System.Drawing.Size(390, 330)
$ChatBox.ScrollBars = "Vertical"
$ChatBox.ReadOnly = $true
$ChatBox.BackColor = [System.Drawing.Color]::White
$ChatBox.Font = New-Object System.Drawing.Font("Arial", 10)

# Campo de entrada
$InputBox = New-Object System.Windows.Forms.TextBox
$InputBox.Location = New-Object System.Drawing.Point(10, 352)
$InputBox.Size = New-Object System.Drawing.Size(235, 25)
$InputBox.Font = New-Object System.Drawing.Font("Arial", 10)
$InputBox.MaxLength = 300

# Botón enviar
$BotonEnviar = New-Object System.Windows.Forms.Button
$BotonEnviar.Location = New-Object System.Drawing.Point(255, 352)
$BotonEnviar.Size = New-Object System.Drawing.Size(72, 25)
$BotonEnviar.Text = "Enviar"

# Botón cerrar
$BotonCerrar = New-Object System.Windows.Forms.Button
$BotonCerrar.Location = New-Object System.Drawing.Point(337, 352)
$BotonCerrar.Size = New-Object System.Drawing.Size(72, 25)
$BotonCerrar.Text = "Cerrar"

# Etiqueta de estado
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Point(10, 388)
$StatusLabel.Size = New-Object System.Drawing.Size(390, 18)
$StatusLabel.Text = "Conectado como: $NombreEquipo"
$StatusLabel.Font = New-Object System.Drawing.Font("Arial", 8)
$StatusLabel.ForeColor = [System.Drawing.Color]::DarkGray

# Agregar controles
$Form.Controls.Add($ChatBox)
$Form.Controls.Add($InputBox)
$Form.Controls.Add($BotonEnviar)
$Form.Controls.Add($BotonCerrar)
$Form.Controls.Add($StatusLabel)

# Variable de control
$script:ChatActivo = $true
$script:UltimoContenido = ""
$script:TieneNotificacion = $false

# Guardar ruta del script actual (necesario para auto-eliminación)
$script:RutaScriptActual = $MyInvocation.MyCommand.Path

# Función enviar mensaje
function Enviar-Mensaje {
    if (-not [string]::IsNullOrWhiteSpace($InputBox.Text)) {
        $fecha = Get-Date -Format "HH:mm:ss"
        $linea = "[$fecha] $NombreEquipo`: $($InputBox.Text)"
        
        # Intentar escribir con reintentos para evitar conflictos
        $intentos = 0
        $maxIntentos = 5
        $escrito = $false
        
        while ($intentos -lt $maxIntentos -and -not $escrito) {
            try {
                Add-Content -Path $RutaArchivo -Value $linea -Encoding utf8 -ErrorAction Stop
                $InputBox.Text = ""
                $escrito = $true
            } catch {
                $intentos++
                Start-Sleep -Milliseconds 50
            }
        }
    }
}

$BotonEnviar.Add_Click({ Enviar-Mensaje })

# Enviar con Enter
$InputBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Enviar-Mensaje
        $_.SuppressKeyPress = $true
    }
})

# Función cerrar chat (limpieza remota con PowerShell)
function Cerrar-ChatOrdenado {
    # Preguntar si se desea cerrar también el chat del usuario remoto
    $respuesta = [System.Windows.Forms.MessageBox]::Show(
        "¿Desea cerrar también el chat en el equipo remoto?",
        "Cerrar Chat Remoto",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    # Detener timer
    $script:ChatActivo = $false
    $timer.Stop()

    # Enviar mensaje de desconexión (con reintentos)
    $intentos = 0
    $maxIntentos = 3
    while ($intentos -lt $maxIntentos) {
        try {
            $fecha = Get-Date -Format "HH:mm:ss"
            $lineaCierre = "[$fecha] *** $NombreEquipo se ha desconectado ***"
            Add-Content -Path $RutaArchivo -Value $lineaCierre -Encoding utf8 -ErrorAction Stop
            break
        } catch {
            $intentos++
            Start-Sleep -Milliseconds 50
        }
    }

    # Si el usuario confirma, cerrar el chat remoto
    if ($respuesta -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Extraer nombre del equipo remoto de la ruta UNC
        if ($RutaArchivo -match '\\\\([^\\]+)\\') {
            $equipoRemoto = $matches[1]
        } else {
            $equipoRemoto = $null
        }

        if ($equipoRemoto) {
            $carpetaRemota = Split-Path $RutaArchivo -Parent

            # Método 1 (preferido): escribir señal de cierre para que el script remoto se cierre solo
            $archivoSenal = Join-Path $carpetaRemota "chat_close.signal"
            "" | Out-File -FilePath $archivoSenal -Encoding ASCII -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2

            # Método 2 (fallback): matar el proceso remoto por PID específico
            $psexecPath = Join-Path $PSScriptRoot "..\..\tools\psexec.exe"
            $archivoPid  = Join-Path $carpetaRemota "chat_usu.pid"
            if ((Test-Path $archivoPid) -and (Test-Path $psexecPath)) {
                $pidRemoto = (Get-Content $archivoPid -Raw -ErrorAction SilentlyContinue).Trim()
                if ($pidRemoto -match '^\d+$') {
                    Start-Process -FilePath $psexecPath `
                        -ArgumentList "\\$equipoRemoto -accepteula taskkill /F /PID $pidRemoto" `
                        -WindowStyle Hidden -Wait
                    Start-Sleep -Seconds 1
                }
            }

            # Borrar archivos remotos con reintentos
            $chatTxtRemoto  = $RutaArchivo
            $chatPs1Remoto  = Join-Path $carpetaRemota "chat_usu.ps1"

            $maxIntentos = 10
            $intento = 0

            while ($intento -lt $maxIntentos) {
                $intento++
                $borrado = $true

                try {
                    if (Test-Path $chatTxtRemoto) {
                        Remove-Item -Path $chatTxtRemoto -Force -ErrorAction Stop
                    }
                } catch { $borrado = $false }

                try {
                    if (Test-Path $chatPs1Remoto) {
                        Remove-Item -Path $chatPs1Remoto -Force -ErrorAction Stop
                    }
                } catch { $borrado = $false }

                if ($borrado) { break }

                Start-Sleep -Seconds 1
            }
        }
    }

    # Salir inmediatamente (PRIORITARIO)
    [System.Environment]::Exit(0)
}

$BotonCerrar.Add_Click({ Cerrar-ChatOrdenado })

# Evento al cerrar ventana con X (mismo código de limpieza)
$Form.Add_FormClosing({
    param($sender, $e)

    # Preguntar si se desea cerrar también el chat del usuario remoto
    $respuesta = [System.Windows.Forms.MessageBox]::Show(
        "¿Desea cerrar también el chat en el equipo remoto?",
        "Cerrar Chat Remoto",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    # Detener timer
    $script:ChatActivo = $false
    $timer.Stop()

    # Enviar mensaje de desconexión (con reintentos)
    $intentos = 0
    $maxIntentos = 3
    while ($intentos -lt $maxIntentos) {
        try {
            $fecha = Get-Date -Format "HH:mm:ss"
            $lineaCierre = "[$fecha] *** $NombreEquipo se ha desconectado ***"
            Add-Content -Path $RutaArchivo -Value $lineaCierre -Encoding utf8 -ErrorAction Stop
            break
        } catch {
            $intentos++
            Start-Sleep -Milliseconds 50
        }
    }

    # Si el usuario confirma, cerrar el chat remoto
    if ($respuesta -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Extraer nombre del equipo remoto de la ruta UNC
        if ($RutaArchivo -match '\\\\([^\\]+)\\') {
            $equipoRemoto = $matches[1]
        } else {
            $equipoRemoto = $null
        }

        if ($equipoRemoto) {
            $carpetaRemota = Split-Path $RutaArchivo -Parent

            # Método 1 (preferido): escribir señal de cierre para que el script remoto se cierre solo
            $archivoSenal = Join-Path $carpetaRemota "chat_close.signal"
            "" | Out-File -FilePath $archivoSenal -Encoding ASCII -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2

            # Método 2 (fallback): matar el proceso remoto por PID específico
            $psexecPath = Join-Path $PSScriptRoot "..\..\tools\psexec.exe"
            $archivoPid  = Join-Path $carpetaRemota "chat_usu.pid"
            if ((Test-Path $archivoPid) -and (Test-Path $psexecPath)) {
                $pidRemoto = (Get-Content $archivoPid -Raw -ErrorAction SilentlyContinue).Trim()
                if ($pidRemoto -match '^\d+$') {
                    Start-Process -FilePath $psexecPath `
                        -ArgumentList "\\$equipoRemoto -accepteula taskkill /F /PID $pidRemoto" `
                        -WindowStyle Hidden -Wait
                    Start-Sleep -Seconds 1
                }
            }

            # Borrar archivos remotos con reintentos
            $chatTxtRemoto  = $RutaArchivo
            $chatPs1Remoto  = Join-Path $carpetaRemota "chat_usu.ps1"

            $maxIntentos = 10
            $intento = 0

            while ($intento -lt $maxIntentos) {
                $intento++
                $borrado = $true

                try {
                    if (Test-Path $chatTxtRemoto) {
                        Remove-Item -Path $chatTxtRemoto -Force -ErrorAction Stop
                    }
                } catch { $borrado = $false }

                try {
                    if (Test-Path $chatPs1Remoto) {
                        Remove-Item -Path $chatPs1Remoto -Force -ErrorAction Stop
                    }
                } catch { $borrado = $false }

                if ($borrado) { break }

                Start-Sleep -Seconds 1
            }
        }
    }

    # Salir inmediatamente (PRIORITARIO)
    [System.Environment]::Exit(0)
})


# Función para actualizar ChatBox con colores
function Actualizar-ChatConColores {
    param([string]$contenido)
    
    $ChatBox.Clear()
    $lineas = $contenido -split "`n"
    
    foreach ($linea in $lineas) {
        if ([string]::IsNullOrWhiteSpace($linea)) { continue }
        
        # Parsear línea: [HH:mm:ss] NOMBRE: texto
        if ($linea -match '^\[(\d{2}:\d{2}:\d{2})\]\s+(\S+?):\s+(.*)$') {
            $hora = $matches[1]
            $nombre = $matches[2]
            $mensaje = $matches[3]
            
            # Añadir hora (gris claro)
            $ChatBox.SelectionStart = $ChatBox.TextLength
            $ChatBox.SelectionColor = [System.Drawing.Color]::Gray
            $ChatBox.SelectionFont = New-Object System.Drawing.Font("Arial", 8)
            $ChatBox.AppendText("[$hora] ")
            
            # Añadir nombre (azul para SOPORTE, negro para otros)
            $ChatBox.SelectionStart = $ChatBox.TextLength
            if ($nombre -eq "SOPORTE") {
                $ChatBox.SelectionColor = [System.Drawing.Color]::Blue
            } else {
                $ChatBox.SelectionColor = [System.Drawing.Color]::Black
            }
            $ChatBox.SelectionFont = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
            $ChatBox.AppendText("$nombre`: ")
            
            # Añadir mensaje (negro, tamaño normal)
            $ChatBox.SelectionStart = $ChatBox.TextLength
            $ChatBox.SelectionColor = [System.Drawing.Color]::Black
            $ChatBox.SelectionFont = New-Object System.Drawing.Font("Arial", 10)
            $ChatBox.AppendText("$mensaje`n")
            
        } elseif ($linea -match '^\[(\d{2}:\d{2}:\d{2})\]\s+\*\*\*\s+(.+?)\s+se ha (conectado|desconectado)\s+\*\*\*') {
            # Mensajes del sistema (gris)
            $ChatBox.SelectionStart = $ChatBox.TextLength
            $ChatBox.SelectionColor = [System.Drawing.Color]::DarkGray
            $ChatBox.SelectionFont = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Italic)
            $ChatBox.AppendText("$linea`n")
        } else {
            # Línea sin formato reconocido
            $ChatBox.SelectionStart = $ChatBox.TextLength
            $ChatBox.SelectionColor = [System.Drawing.Color]::Black
            $ChatBox.SelectionFont = New-Object System.Drawing.Font("Arial", 10)
            $ChatBox.AppendText("$linea`n")
        }
    }
    
    $ChatBox.SelectionStart = $ChatBox.TextLength
    $ChatBox.ScrollToCaret()
}

# Timer para sincronización
$script:UltimoContenido = ""
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $Intervalo

# Restaurar título al recibir foco
$Form.Add_Activated({
    if ($script:TieneNotificacion) {
        $Form.Text = $TituloVentana
        $script:TieneNotificacion = $false
    }
})

$timer.Add_Tick({
    if (-not $script:ChatActivo) {
        $timer.Stop()
        return
    }

    try {
        if (Test-Path $RutaArchivo) {
            $contenido = Get-Content -Path $RutaArchivo -Raw -ErrorAction SilentlyContinue
            if ($contenido -and $contenido -ne $script:UltimoContenido) {
                Actualizar-ChatConColores -contenido $contenido
                $script:UltimoContenido = $contenido
                # Notificación en título si la ventana no tiene foco
                if (-not $Form.ContainsFocus) {
                    $Form.Text = "(!) $TituloVentana"
                    $script:TieneNotificacion = $true
                }
            }
        }
    } catch {}
})

# Carga inicial
if (Test-Path $RutaArchivo) {
    try {
        $contenidoInicial = Get-Content -Path $RutaArchivo -Raw -ErrorAction SilentlyContinue
        if ($contenidoInicial) {
            Actualizar-ChatConColores -contenido $contenidoInicial
            $script:UltimoContenido = $contenidoInicial
        }
    } catch {}
} else {
    "" | Out-File -FilePath $RutaArchivo -Encoding utf8 -Force
}

# Mensaje de conexión
$intentos = 0
$maxIntentos = 5
while ($intentos -lt $maxIntentos) {
    try {
        $fecha = Get-Date -Format "HH:mm:ss"
        $lineaConexion = "[$fecha] *** $NombreEquipo se ha conectado ***"
        Add-Content -Path $RutaArchivo -Value $lineaConexion -Encoding utf8 -ErrorAction Stop
        break
    } catch {
        $intentos++
        Start-Sleep -Milliseconds 50
    }
}

# Iniciar timer y mostrar ventana
$timer.Enabled = $true
$Form.Show()
[System.Windows.Forms.Application]::Run($Form)



