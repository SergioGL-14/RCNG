# Lado remoto del chat.
# Se copia al %TEMP% del usuario y mantiene la conversación desde el propio
# equipo remoto, con limpieza automática al cerrar.

# Forzar STA para que WinForms funcione correctamente en el equipo remoto.
[System.Threading.Thread]::CurrentThread.ApartmentState = [System.Threading.ApartmentState]::STA

# Cargar las librerías de la ventana.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configuración básica de refresco.
$Intervalo = 250  # milisegundos

# Datos propios del lado remoto.
$NombreEquipo = $env:COMPUTERNAME
$RutaArchivo = Join-Path $PSScriptRoot "chat.txt"
$TituloVentana = "Chat - $NombreEquipo"

# Crear ventana
$Form = New-Object System.Windows.Forms.Form
$Form.Text = $TituloVentana
$Form.Size = New-Object System.Drawing.Size(420, 500)
$Form.StartPosition = "CenterScreen"
$Form.Topmost = $true
$Form.BackColor = [System.Drawing.Color]::White

# Ãrea de mensajes con colores por remitente.
$ChatBox = New-Object System.Windows.Forms.RichTextBox
$ChatBox.Multiline = $true
$ChatBox.Location = New-Object System.Drawing.Point(10, 10)
$ChatBox.Size = New-Object System.Drawing.Size(390, 330)
$ChatBox.ScrollBars = "Vertical"
$ChatBox.ReadOnly = $true
$ChatBox.ShortcutsEnabled = $false  # Deshabilitar copiar/pegar con Ctrl+C, Ctrl+V, etc.
$ChatBox.BackColor = [System.Drawing.Color]::White
$ChatBox.Font = New-Object System.Drawing.Font("Arial", 10)

# Campo de entrada
$InputBox = New-Object System.Windows.Forms.TextBox
$InputBox.Location = New-Object System.Drawing.Point(10, 352)
$InputBox.Size = New-Object System.Drawing.Size(235, 25)
$InputBox.Font = New-Object System.Drawing.Font("Arial", 10)
$InputBox.MaxLength = 300

# BotÃ³n enviar
$BotonEnviar = New-Object System.Windows.Forms.Button
$BotonEnviar.Location = New-Object System.Drawing.Point(255, 352)
$BotonEnviar.Size = New-Object System.Drawing.Size(72, 25)
$BotonEnviar.Text = "Enviar"

# BotÃ³n cerrar
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

# Guardar rutas auxiliares para limpieza y cierre coordinado.
$script:RutaScriptActual = $MyInvocation.MyCommand.Path
$script:RutaPidFile = Join-Path $PSScriptRoot "chat_usu.pid"
$script:RutaSenal  = Join-Path $PSScriptRoot "chat_close.signal"

# Registrar el PID permite que el lado soporte cierre esta ventana de forma limpia.
$PID | Out-File -FilePath $script:RutaPidFile -Encoding ASCII -Force -ErrorAction SilentlyContinue

# FunciÃ³n enviar mensaje
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

# FunciÃ³n cerrar chat (limpieza local con BAT)
function Cerrar-ChatOrdenado {
    # Detener timer
    $script:ChatActivo = $false
    $timer.Stop()

    # Eliminar archivos de control de esta sesiÃ³n
    Remove-Item $script:RutaPidFile -Force -ErrorAction SilentlyContinue
    Remove-Item $script:RutaSenal  -Force -ErrorAction SilentlyContinue

    # Enviar mensaje de desconexiÃ³n (con reintentos)
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
    
    # Limpieza LOCAL - Crear BAT que espera y borra archivos en %TEMP%
    $batPath = Join-Path $PSScriptRoot "limpiar_chat.bat"
    $miScript = $script:RutaScriptActual
    $miChatTxt = $RutaArchivo
    
    # BAT que espera hasta que los archivos estÃ©n liberados, luego los borra
    @"
@echo off
REM Esperar 3 segundos para que ambas ventanas se cierren
timeout /t 3 /nobreak >nul

REM Intentar borrar archivos en bucle (mÃ¡ximo 10 intentos)
set intentos=0
:retry
set /a intentos+=1
if %intentos% GTR 10 goto end

REM Borrar chat.txt
if exist "$miChatTxt" (
    del /F /Q "$miChatTxt" 2>nul
    if exist "$miChatTxt" (
        timeout /t 1 /nobreak >nul
        goto retry
    )
)

REM Borrar chat_usu.ps1
if exist "$miScript" (
    del /F /Q "$miScript" 2>nul
)

REM Auto-eliminarse
del /F /Q "$batPath" 2>nul

:end
"@ | Out-File -FilePath $batPath -Encoding ASCII -Force
    
    # Ejecutar el BAT en background
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batPath`"" -WindowStyle Hidden
    
    # Salir inmediatamente (PRIORITARIO)
    [System.Environment]::Exit(0)
}

$BotonCerrar.Add_Click({ Cerrar-ChatOrdenado })

# Evento al cerrar ventana con X (mismo cÃ³digo de limpieza)
$Form.Add_FormClosing({
    param($sender, $e)
    # Detener timer
    $script:ChatActivo = $false
    $timer.Stop()

    # Eliminar archivos de control de esta sesiÃ³n
    Remove-Item $script:RutaPidFile -Force -ErrorAction SilentlyContinue
    Remove-Item $script:RutaSenal  -Force -ErrorAction SilentlyContinue

    # Enviar mensaje de desconexiÃ³n (con reintentos)
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
    
    # Limpieza LOCAL - Crear BAT que espera y borra archivos en %TEMP%
    $batPath = Join-Path $PSScriptRoot "limpiar_chat.bat"
    $miScript = $script:RutaScriptActual
    $miChatTxt = $RutaArchivo
    
    # BAT que espera hasta que los archivos estÃ©n liberados, luego los borra
    @"
@echo off
REM Esperar 3 segundos para que ambas ventanas se cierren
timeout /t 3 /nobreak >nul

REM Intentar borrar archivos en bucle (mÃ¡ximo 10 intentos)
set intentos=0
:retry
set /a intentos+=1
if %intentos% GTR 10 goto end

REM Borrar chat.txt
if exist "$miChatTxt" (
    del /F /Q "$miChatTxt" 2>nul
    if exist "$miChatTxt" (
        timeout /t 1 /nobreak >nul
        goto retry
    )
)

REM Borrar chat_usu.ps1
if exist "$miScript" (
    del /F /Q "$miScript" 2>nul
)

REM Auto-eliminarse
del /F /Q "$batPath" 2>nul

:end
"@ | Out-File -FilePath $batPath -Encoding ASCII -Force
    
    # Ejecutar el BAT en background
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batPath`"" -WindowStyle Hidden
    
    # Salir inmediatamente (PRIORITARIO)
    [System.Environment]::Exit(0)
})

# FunciÃ³n para actualizar ChatBox con colores
function Actualizar-ChatConColores {
    param([string]$contenido)
    
    $ChatBox.Clear()
    $lineas = $contenido -split "`n"
    
    foreach ($linea in $lineas) {
        if ([string]::IsNullOrWhiteSpace($linea)) { continue }
        
        # Parsear lÃ­nea: [HH:mm:ss] NOMBRE: texto
        if ($linea -match '^\[(\d{2}:\d{2}:\d{2})\]\s+(\S+?):\s+(.*)$') {
            $hora = $matches[1]
            $nombre = $matches[2]
            $mensaje = $matches[3]
            
            # AÃ±adir hora (gris claro)
            $ChatBox.SelectionStart = $ChatBox.TextLength
            $ChatBox.SelectionColor = [System.Drawing.Color]::Gray
            $ChatBox.SelectionFont = New-Object System.Drawing.Font("Arial", 8)
            $ChatBox.AppendText("[$hora] ")
            
            # AÃ±adir nombre (azul para SOPORTE, negro para otros)
            $ChatBox.SelectionStart = $ChatBox.TextLength
            if ($nombre -eq "SOPORTE") {
                $ChatBox.SelectionColor = [System.Drawing.Color]::Blue
            } else {
                $ChatBox.SelectionColor = [System.Drawing.Color]::Black
            }
            $ChatBox.SelectionFont = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
            $ChatBox.AppendText("$nombre`: ")
            
            # AÃ±adir mensaje (negro, tamaÃ±o normal)
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
            # LÃ­nea sin formato reconocido
            $ChatBox.SelectionStart = $ChatBox.TextLength
            $ChatBox.SelectionColor = [System.Drawing.Color]::Black
            $ChatBox.SelectionFont = New-Object System.Drawing.Font("Arial", 10)
            $ChatBox.AppendText("$linea`n")
        }
    }
    
    $ChatBox.SelectionStart = $ChatBox.TextLength
    $ChatBox.ScrollToCaret()
}

# Timer para sincronizaciÃ³n
$script:UltimoContenido = ""
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $Intervalo

# Restaurar tÃ­tulo al recibir foco
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

    # Verificar seÃ±al de cierre remoto
    if (Test-Path $script:RutaSenal) {
        Remove-Item $script:RutaSenal -Force -ErrorAction SilentlyContinue
        Cerrar-ChatOrdenado
        return
    }

    try {
        if (Test-Path $RutaArchivo) {
            $contenido = Get-Content -Path $RutaArchivo -Raw -ErrorAction SilentlyContinue
            if ($contenido -and $contenido -ne $script:UltimoContenido) {
                Actualizar-ChatConColores -contenido $contenido
                $script:UltimoContenido = $contenido
                # NotificaciÃ³n en tÃ­tulo si la ventana no tiene foco
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

# Mensaje de conexiÃ³n
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





