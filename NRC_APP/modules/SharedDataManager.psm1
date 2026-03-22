# SharedDataManager centraliza todo lo que puede leerse o publicarse en el
# recurso compartido del proyecto: SQLite, JSON, scripts custom e iconos.
# Si el servidor no responde, la aplicacion sigue trabajando con la copia local.

#==================================================================
# BLOQUE: Configuración
#==================================================================
$script:DefaultSettings = [ordered]@{
    SharedServerBase      = '\\server\share\NRC_APP'
    ProxyPacUrl           = 'http://proxy.example.local/proxy.pac'
    PortalUrl             = 'https://portal.example.local'
    MailPortalUrl         = 'https://mail.example.local'
    WolCsvShare           = '\\server\share\network'
    WolCsvFileName        = 'network_inventory_sample.csv'
    DhcpServer            = 'DHCP-SERVER'
    SupportDisplayName    = 'NRC_APP Support'
    SupportEmail          = 'support@example.local'
    PrimaryGroupSearchBase   = 'OU=GrupoPrincipal,DC=example,DC=local'
    SecondaryGroupSearchBase = 'OU=GrupoSecundario,DC=example,DC=local'
}
$script:ServerBase      = $script:DefaultSettings.SharedServerBase
$script:ServerAvailable = $null   # $null = no testado aún
$script:LogCallback     = $null   # scriptblock: { param($msg) Add-Logs -text $msg }

function Get-AppSettingsPath {
    return Join-Path $Global:ScriptRoot 'database\appsettings.json'
}

function Get-DefaultAppSettings {
    return [ordered]@{
        SharedServerBase      = $script:DefaultSettings.SharedServerBase
        ProxyPacUrl           = $script:DefaultSettings.ProxyPacUrl
        PortalUrl             = $script:DefaultSettings.PortalUrl
        MailPortalUrl         = $script:DefaultSettings.MailPortalUrl
        WolCsvShare           = $script:DefaultSettings.WolCsvShare
        WolCsvFileName        = $script:DefaultSettings.WolCsvFileName
        DhcpServer            = $script:DefaultSettings.DhcpServer
        SupportDisplayName    = $script:DefaultSettings.SupportDisplayName
        SupportEmail          = $script:DefaultSettings.SupportEmail
        PrimaryGroupSearchBase   = $script:DefaultSettings.PrimaryGroupSearchBase
        SecondaryGroupSearchBase = $script:DefaultSettings.SecondaryGroupSearchBase
    }
}

function Get-AppSettings {
    $defaults = Get-DefaultAppSettings
    $path = Get-AppSettingsPath
    if (-not (Test-Path $path)) { return $defaults }

    try {
        $raw = Get-Content $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $defaults }
        $data = $raw | ConvertFrom-Json -ErrorAction Stop
        foreach ($key in $defaults.Keys) {
            if (-not $data.PSObject.Properties[$key] -or [string]::IsNullOrWhiteSpace([string]$data.$key)) {
                $data | Add-Member -NotePropertyName $key -NotePropertyValue $defaults[$key] -Force
            }
        }
        $result = [ordered]@{}
        foreach ($key in $defaults.Keys) {
            $result[$key] = [string]$data.$key
        }
        return $result
    } catch {
        return $defaults
    }
}

function Save-AppSettings {
    param([Parameter(Mandatory=$true)][hashtable]$Settings)

    $defaults = Get-DefaultAppSettings
    $merged = [ordered]@{}
    foreach ($key in $defaults.Keys) {
        $value = if ($Settings.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$Settings[$key])) {
            [string]$Settings[$key]
        } else {
            [string]$defaults[$key]
        }
        $merged[$key] = $value
    }

    $path = Get-AppSettingsPath
    $dir  = Split-Path $path -Parent
    if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    ($merged | ConvertTo-Json -Depth 3) | Out-File -FilePath $path -Encoding UTF8 -Force

    $script:ServerBase = $merged.SharedServerBase
    $script:ServerAvailable = $null
    return $merged
}

function Get-AppSettingValue {
    param(
        [Parameter(Mandatory=$true)][string]$Key,
        [string]$DefaultValue = ''
    )

    $settings = Get-AppSettings
    if ($settings.Contains($Key) -and -not [string]::IsNullOrWhiteSpace([string]$settings[$Key])) {
        return [string]$settings[$Key]
    }
    if (-not [string]::IsNullOrWhiteSpace($DefaultValue)) { return $DefaultValue }
    return ''
}

$script:ServerBase = Get-AppSettingValue -Key 'SharedServerBase' -DefaultValue $script:DefaultSettings.SharedServerBase

#==================================================================
# BLOQUE: Logging interno
#==================================================================

function Register-ServerLogCallback {
    <#
    .SYNOPSIS
        Registra el scriptblock que SharedDataManager usará para enviar mensajes al log de la UI cuando el formulario principal ya tenga Add-Logs listo.
    #>
    param([Parameter(Mandatory=$true)][scriptblock]$Callback)
    $script:LogCallback = $Callback
}

function script:Write-ServerLog {
    param([string]$Message)
    try {
        if ($script:LogCallback) { & $script:LogCallback $Message }
        else { Write-Output $Message }
    } catch { Write-Output $Message }
}

#==================================================================
# BLOQUE: Disponibilidad del servidor
#==================================================================

function Test-ServerAvailable {
    <#
    .SYNOPSIS
        Comprueba si la ruta UNC del servidor está accesible. Usa un ping corto para no bloquear la UI con esperas largas de SMB. El resultado queda cacheado; usa -Force para repetir la comprobación.
    #>
    param([switch]$Force)

    if (-not $Force -and $null -ne $script:ServerAvailable) {
        return $script:ServerAvailable
    }

    $script:ServerAvailable = $false
    try {
        # Extraer hostname del UNC (ej: \\servidor\share → servidor)
        $serverHost = [regex]::Match($script:ServerBase, '^\\\\([^\\]+)').Groups[1].Value
        if ([string]::IsNullOrWhiteSpace($serverHost)) {
            script:Write-ServerLog "⚠️ Servidor: no se pudo extraer hostname de '$script:ServerBase'"
            return $false
        }

        # Ping rápido con timeout 1 s — evita el bloqueo de 30-60 s del sondeo SMB
        $ping   = [System.Net.NetworkInformation.Ping]::new()
        $result = $ping.Send($serverHost, 1000)
        $ping.Dispose()

        if ($result.Status -ne [System.Net.NetworkInformation.IPStatus]::Success) {
            script:Write-ServerLog "⚠️ Servidor '$serverHost' no responde al ping → usando datos locales"
            return $false
        }

        # Ping OK → verificar que la carpeta compartida es accesible
        $script:ServerAvailable = [System.IO.Directory]::Exists($script:ServerBase)
        if ($script:ServerAvailable) {
            script:Write-ServerLog "✅ Recurso compartido accesible: $script:ServerBase"
        } else {
            script:Write-ServerLog "⚠️ Ping OK pero carpeta UNC no accesible: $script:ServerBase → usando datos locales"
        }
    } catch {
        $script:ServerAvailable = $false
        script:Write-ServerLog "⚠️ Error comprobando servidor: $($_.Exception.Message) → usando datos locales"
    }
    return $script:ServerAvailable
}

function Get-ServerBase {
    return $script:ServerBase
}

#==================================================================
# BLOQUE: Cadenas de conexión SQLite con soporte UNC
#==================================================================

function Open-SQLiteConnection {
    <#
    .SYNOPSIS
        Abre una conexión SQLite con ParseViaFramework como propiedad del objeto
        (necesario para rutas UNC). El flag debe ser propiedad, NO keyword de la
        cadena de conexión.
    #>
    param([Parameter(Mandatory=$true)][string]$DbPath)

    $connString = "Data Source=$DbPath;Version=3;Journal Mode=DELETE;BusyTimeout=5000;"
    $conn = New-Object System.Data.SQLite.SQLiteConnection -ArgumentList $connString
    $conn.ParseViaFramework = $true
    $conn.Open()
    return $conn
}

function Get-SQLiteConnString {
    <#
    .SYNOPSIS
        Devuelve la cadena de conexión base.
        NOTA: Para rutas UNC, usar Open-SQLiteConnection que establece
        ParseViaFramework como propiedad del objeto.
    #>
    param([string]$DbPath)
    return "Data Source=$DbPath;Version=3;Journal Mode=DELETE;BusyTimeout=5000;"
}

#==================================================================
# BLOQUE: Resolución de rutas activas (servidor o local)
#==================================================================

function Get-ActiveComputerDBPath {
    <#
    .SYNOPSIS
        Devuelve la ruta activa de ComputerNames.sqlite:
        servidor si está disponible, local como fallback.
    #>
    $serverPath = Join-Path $script:ServerBase 'ComputerNames.sqlite'
    if ((Test-ServerAvailable) -and (Test-Path $serverPath -ErrorAction SilentlyContinue)) {
        return $serverPath
    }
    return Join-Path $Global:ScriptRoot 'database\ComputerNames.sqlite'
}

#==================================================================
# BLOQUE: Lectura de JSONs (LOCAL siempre, servidor es fuente tras sync)
#==================================================================

function Get-ScriptsDbContent {
    <#
    .SYNOPSIS
        Lee scripts_db.json local. Si no existe, intenta el del servidor.
        Tras un sync completo, el local YA tiene la versión del servidor.
    #>
    $localPath = Join-Path $Global:ScriptRoot 'database\scripts_db.json'
    if (Test-Path $localPath) {
        $raw = Get-Content $localPath -Raw -Encoding UTF8
        if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
    }
    # Fallback: intentar directo del servidor si el local está vacío
    $serverPath = Join-Path $script:ServerBase 'scripts_db.json'
    if ((Test-ServerAvailable) -and (Test-Path $serverPath -ErrorAction SilentlyContinue)) {
        try { return Get-Content $serverPath -Raw -Encoding UTF8 } catch {}
    }
    return '[]'
}

function Get-AppsDbContent {
    <#
    .SYNOPSIS
        Lee apps_db.json local. Si no existe, intenta el del servidor.
    #>
    $localPath = Join-Path $Global:ScriptRoot 'database\apps_db.json'
    if (Test-Path $localPath) {
        $raw = Get-Content $localPath -Raw -Encoding UTF8
        if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
    }
    $serverPath = Join-Path $script:ServerBase 'apps_db.json'
    if ((Test-ServerAvailable) -and (Test-Path $serverPath -ErrorAction SilentlyContinue)) {
        try { return Get-Content $serverPath -Raw -Encoding UTF8 } catch {}
    }
    return '[]'
}

#==================================================================
# BLOQUE: Escritura de JSONs
#   - Save-*DbLocal   : solo guarda en local (uso normal)
#   - Sync-*JsonToServer : empuja el local al servidor (uso tras confirmación)
#   - Save-*DbBoth    : guarda en ambos (uso legacy / compatibilidad)
#==================================================================

function Save-ScriptsDbLocal {
    param([Parameter(Mandatory=$true)][string]$JsonContent)
    $localPath = Join-Path $Global:ScriptRoot 'database\scripts_db.json'
    $JsonContent | Out-File -FilePath $localPath -Encoding UTF8 -Force
}

function Save-AppsDbLocal {
    param([Parameter(Mandatory=$true)][string]$JsonContent)
    $localPath = Join-Path $Global:ScriptRoot 'database\apps_db.json'
    $JsonContent | Out-File -FilePath $localPath -Encoding UTF8 -Force
}

function Sync-ScriptsJsonToServer {
    <#
    .SYNOPSIS
        Empuja el scripts_db.json local al servidor.
        Llamar solo tras confirmación explícita del usuario.
    .OUTPUTS
        $true si se pudo guardar, $false en caso de error.
    #>
    if (-not (Test-ServerAvailable)) {
        script:Write-ServerLog "⚠️ Servidor no disponible: no se pudo replicar scripts_db.json"
        return $false
    }
    $localPath  = Join-Path $Global:ScriptRoot 'database\scripts_db.json'
    $serverPath = Join-Path $script:ServerBase 'scripts_db.json'
    try {
        Copy-Item -Path $localPath -Destination $serverPath -Force -ErrorAction Stop
        script:Write-ServerLog "✅ scripts_db.json replicado al servidor"
        return $true
    } catch {
        script:Write-ServerLog "⚠️ Error replicando scripts_db.json: $($_.Exception.Message)"
        return $false
    }
}

function Sync-AppsJsonToServer {
    <#
    .SYNOPSIS
        Empuja el apps_db.json local al servidor.
    .OUTPUTS
        $true si se pudo guardar, $false en caso de error.
    #>
    if (-not (Test-ServerAvailable)) {
        script:Write-ServerLog "⚠️ Servidor no disponible: no se pudo replicar apps_db.json"
        return $false
    }
    $localPath  = Join-Path $Global:ScriptRoot 'database\apps_db.json'
    $serverPath = Join-Path $script:ServerBase 'apps_db.json'
    try {
        Copy-Item -Path $localPath -Destination $serverPath -Force -ErrorAction Stop
        script:Write-ServerLog "✅ apps_db.json replicado al servidor"
        return $true
    } catch {
        script:Write-ServerLog "⚠️ Error replicando apps_db.json: $($_.Exception.Message)"
        return $false
    }
}

function Save-ScriptsDbBoth {
    param([Parameter(Mandatory=$true)][string]$JsonContent)
    Save-ScriptsDbLocal -JsonContent $JsonContent
    Sync-ScriptsJsonToServer | Out-Null
}

function Save-AppsDbBoth {
    param([Parameter(Mandatory=$true)][string]$JsonContent)
    Save-AppsDbLocal -JsonContent $JsonContent
    Sync-AppsJsonToServer | Out-Null
}

#==================================================================
# BLOQUE: Publicación al servidor (script + icono + SQLite DBs)
#==================================================================

function Sync-ComputerDbToServer {
    <#
    .SYNOPSIS
        Copia ComputerNames.sqlite local al servidor.
    .OUTPUTS
        $true si se pudo copiar, $false en caso de error.
    #>
    if (-not (Test-ServerAvailable)) {
        script:Write-ServerLog "⚠️ Servidor no disponible: no se pudo replicar ComputerNames.sqlite"
        return $false
    }
    $localPath  = Join-Path $Global:ScriptRoot 'database\ComputerNames.sqlite'
    $serverPath = Join-Path $script:ServerBase 'ComputerNames.sqlite'
    try {
        Copy-Item -Path $localPath -Destination $serverPath -Force -ErrorAction Stop
        script:Write-ServerLog "✅ ComputerNames.sqlite replicado al servidor"
        return $true
    } catch {
        script:Write-ServerLog "⚠️ Error replicando ComputerNames.sqlite: $($_.Exception.Message)"
        return $false
    }
}

function Publish-ScriptToServer {
    <#
    .SYNOPSIS
        Copia un archivo de script (.ps1/.bat) a la carpeta scripts\ del servidor.
        Solo para scripts custom (no built-in).
    .OUTPUTS
        $true si se copió con éxito, $false en caso contrario.
    #>
    param([Parameter(Mandatory=$true)][string]$LocalScriptPath)

    if (-not (Test-ServerAvailable)) { return $false }

    $serverScriptsDir = Join-Path $script:ServerBase 'scripts'
    if (-not (Test-Path $serverScriptsDir -ErrorAction SilentlyContinue)) {
        try { New-Item -Path $serverScriptsDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
        catch { return $false }
    }

    $fileName = [System.IO.Path]::GetFileName($LocalScriptPath)
    $dest     = Join-Path $serverScriptsDir $fileName
    try {
        Copy-Item -Path $LocalScriptPath -Destination $dest -Force -ErrorAction Stop
        script:Write-ServerLog "✅ Script '$fileName' publicado en el servidor"
        return $true
    } catch {
        script:Write-ServerLog "⚠️ No se pudo copiar '$fileName' al servidor: $($_.Exception.Message)"
        return $false
    }
}

function Publish-IconToServer {
    <#
    .SYNOPSIS
        Copia un archivo de icono a la carpeta ico\ del servidor.
    .OUTPUTS
        $true si se copió con éxito, $false en caso contrario.
    #>
    param([Parameter(Mandatory=$true)][string]$LocalIconPath)

    if ([string]::IsNullOrWhiteSpace($LocalIconPath)) { return $false }
    if (-not (Test-ServerAvailable)) { return $false }

    $serverIcoDir = Join-Path $script:ServerBase 'ico'
    if (-not (Test-Path $serverIcoDir -ErrorAction SilentlyContinue)) {
        try { New-Item -Path $serverIcoDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
        catch { return $false }
    }

    $fileName = [System.IO.Path]::GetFileName($LocalIconPath)
    $dest     = Join-Path $serverIcoDir $fileName
    try {
        Copy-Item -Path $LocalIconPath -Destination $dest -Force -ErrorAction Stop
        script:Write-ServerLog "✅ Icono '$fileName' publicado en el servidor"
        return $true
    } catch {
        script:Write-ServerLog "⚠️ No se pudo copiar icono '$fileName' al servidor: $($_.Exception.Message)"
        return $false
    }
}

function Remove-ScriptFromServer {
    <#
    .SYNOPSIS
        Elimina el archivo de script de la carpeta scripts\ del servidor.
    .OUTPUTS
        $true si se eliminó o no existía, $false si hubo error.
    #>
    param([Parameter(Mandatory=$true)][string]$FileName)

    if (-not (Test-ServerAvailable)) { return $false }
    $serverPath = Join-Path $script:ServerBase "scripts\$FileName"
    if (-not (Test-Path $serverPath -ErrorAction SilentlyContinue)) { return $true }
    try {
        Remove-Item -Path $serverPath -Force -ErrorAction Stop
        script:Write-ServerLog "✅ Script '$FileName' eliminado del servidor"
        return $true
    } catch {
        script:Write-ServerLog "⚠️ Error eliminando '$FileName' del servidor: $($_.Exception.Message)"
        return $false
    }
}

#==================================================================
# BLOQUE: Sincronización desde servidor a local
#==================================================================

function Sync-IconFromServer {
    <#
    .SYNOPSIS
        Si el icono no existe en icos\ local, lo descarga del servidor (ico\).
    #>
    param([string]$IconName)

    if ([string]::IsNullOrWhiteSpace($IconName)) { return }

    $localPath = Join-Path $Global:ScriptRoot "icos\$IconName"
    if (Test-Path $localPath) { return }   # Ya existe localmente

    if (-not (Test-ServerAvailable)) { return }

    $serverPath = Join-Path $script:ServerBase "ico\$IconName"
    if (-not (Test-Path $serverPath -ErrorAction SilentlyContinue)) { return }

    try {
        Copy-Item -Path $serverPath -Destination $localPath -Force -ErrorAction Stop
    } catch {}
}

function Sync-ScriptFromServer {
    <#
    .SYNOPSIS
        Si el script no existe en scripts\ local, lo descarga del servidor.
    .OUTPUTS
        $true si ya existía o se descargó correctamente, $false si no aplica.
    #>
    param([string]$FileName)

    if ([string]::IsNullOrWhiteSpace($FileName)) { return $false }

    $localPath = Join-Path $Global:ScriptRoot "scripts\$FileName"
    if (Test-Path $localPath) { return $true }   # Ya existe

    if (-not (Test-ServerAvailable)) { return $false }

    $serverPath = Join-Path $script:ServerBase "scripts\$FileName"
    if (-not (Test-Path $serverPath -ErrorAction SilentlyContinue)) { return $false }

    try {
        Copy-Item -Path $serverPath -Destination $localPath -Force -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

#==================================================================
# BLOQUE: Sincronización completa (botón "Actualizar Datos")
#==================================================================

function Invoke-FullSyncFromServer {
    <#
    .SYNOPSIS
        Sincronización completa desde el servidor:
          1. Re-comprueba disponibilidad del servidor.
          2. Copia ComputerNames.sqlite al local.
          3. Copia scripts_db.json y apps_db.json al local.
          4. Para cada script custom del JSON: descarga el .ps1/.bat si falta.
          5. Para cada script/app: descarga el ico si falta.
          6. Detecta scripts en local que ya no están en el servidor (RemovedFromServer).
          7. Detecta scripts cuyo archivo en el servidor es más nuevo que el local (UpdatedOnServer).
        Si el servidor no está disponible, informa y no hace nada.
    .OUTPUTS
        Hashtable con:
          Success          (bool)
          Message          (string)
          Copied           (string[])  - nombres de archivos copiados
          Errors           (string[])  - mensajes de error no fatales
          RemovedFromServer (object[]) - scripts custom en local que no están en server JSON
                                         Cada elemento: @{ Name, FileName, Id }
          UpdatedOnServer  (object[]) - scripts cuyo archivo en servidor es más nuevo
                                         Cada elemento: @{ Name, FileName, ServerPath, LocalPath }
    #>

    # Re-comprobar disponibilidad real
    if (-not (Test-ServerAvailable -Force)) {
        return @{
            Success           = $false
            Message           = "Servidor no accesible ($script:ServerBase). Se mantendrán los datos locales."
            Copied            = @()
            Errors            = @()
            RemovedFromServer = @()
            UpdatedOnServer   = @()
        }
    }

    $copied  = [System.Collections.Generic.List[string]]::new()
    $errors  = [System.Collections.Generic.List[string]]::new()
    $removed = [System.Collections.Generic.List[object]]::new()
    $updated = [System.Collections.Generic.List[object]]::new()

    # --- Snapshot ANTES de copiar el JSON ---
    # Capturar los scripts custom locales ANTES de sobreescribir con el del servidor
    $localScriptsDbPath = Join-Path $Global:ScriptRoot 'database\scripts_db.json'
    $preScripts = @()
    if (Test-Path $localScriptsDbPath) {
        try {
            $preRaw = Get-Content $localScriptsDbPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($preRaw -isnot [array]) { $preRaw = @($preRaw) }
            $preScripts = $preRaw | Where-Object { -not $_.BuiltIn -and -not [string]::IsNullOrWhiteSpace($_.FileName) }
        } catch {}
    }

    # --- 1) Bases de datos SQLite ---
    foreach ($dbFile in @('ComputerNames.sqlite')) {
        $src = Join-Path $script:ServerBase $dbFile
        $dst = Join-Path $Global:ScriptRoot "database\$dbFile"
        if (Test-Path $src -ErrorAction SilentlyContinue) {
            try {
                Copy-Item -Path $src -Destination $dst -Force -ErrorAction Stop
                $copied.Add($dbFile)
                script:Write-ServerLog "Copiado: $dbFile"
            } catch {
                $errors.Add("Error copiando $dbFile`: $($_.Exception.Message)")
            }
        }
    }

    # --- 2) Archivos JSON ---
    foreach ($jsonFile in @('scripts_db.json', 'apps_db.json')) {
        $src = Join-Path $script:ServerBase $jsonFile
        $dst = Join-Path $Global:ScriptRoot "database\$jsonFile"
        if (Test-Path $src -ErrorAction SilentlyContinue) {
            try {
                Copy-Item -Path $src -Destination $dst -Force -ErrorAction Stop
                $copied.Add($jsonFile)
                script:Write-ServerLog "Copiado: $jsonFile"
            } catch {
                $errors.Add("Error copiando $jsonFile`: $($_.Exception.Message)")
            }
        }
    }

    # --- 3) Scripts custom: descargar archivos faltantes + detectar actualizados ---
    $postScripts = @()
    if (Test-Path $localScriptsDbPath) {
        try {
            $postRaw = Get-Content $localScriptsDbPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($postRaw -isnot [array]) { $postRaw = @($postRaw) }
            $postScripts = $postRaw | Where-Object { -not $_.BuiltIn -and -not [string]::IsNullOrWhiteSpace($_.FileName) }

            foreach ($s in $postScripts) {
                $localFile  = Join-Path $Global:ScriptRoot "scripts\$($s.FileName)"
                $serverFile = Join-Path $script:ServerBase "scripts\$($s.FileName)"

                if (Test-Path $serverFile -ErrorAction SilentlyContinue) {
                    if (-not (Test-Path $localFile)) {
                        # Archivo no existe localmente → descargar
                        try {
                            Copy-Item $serverFile $localFile -Force -ErrorAction Stop
                            $copied.Add("script: $($s.FileName)")
                        } catch {
                            $errors.Add("Error descargando script $($s.FileName)`: $($_.Exception.Message)")
                        }
                    } else {
                        # Ambos existen → comparar fechas
                        $serverTime = (Get-Item $serverFile).LastWriteTime
                        $localTime  = (Get-Item $localFile).LastWriteTime
                        if ($serverTime -gt $localTime) {
                            $updated.Add([PSCustomObject]@{
                                Name       = $s.Name
                                FileName   = $s.FileName
                                ServerPath = $serverFile
                                LocalPath  = $localFile
                            })
                        }
                    }
                }

                # Icono del script
                if (-not [string]::IsNullOrWhiteSpace($s.IconFile)) {
                    Sync-IconFromServer -IconName $s.IconFile
                }
            }
        } catch {
            $errors.Add("Error procesando scripts_db.json: $($_.Exception.Message)")
        }
    }

    # --- 4) Detectar scripts removidos del servidor ---
    # Scripts que estaban en local ANTES del sync y ya NO están en el JSON del servidor
    $postFileNames = $postScripts | ForEach-Object { $_.FileName }
    foreach ($pre in $preScripts) {
        if ($postFileNames -notcontains $pre.FileName) {
            $removed.Add([PSCustomObject]@{
                Name     = $pre.Name
                FileName = $pre.FileName
                Id       = $pre.Id
            })
        }
    }

    # --- 5) Iconos de apps ---
    $localAppsDb = Join-Path $Global:ScriptRoot 'database\apps_db.json'
    if (Test-Path $localAppsDb) {
        try {
            $appsData = Get-Content $localAppsDb -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($appsData -isnot [array]) { $appsData = @($appsData) }
            foreach ($a in $appsData) {
                if (-not [string]::IsNullOrWhiteSpace($a.IconName)) {
                    Sync-IconFromServer -IconName $a.IconName
                }
            }
        } catch {
            $errors.Add("Error procesando apps_db.json: $($_.Exception.Message)")
        }
    }

    # --- Resultado ---
    $parts = @()
    if ($copied.Count -gt 0) { $parts += "Copiados: $($copied -join ', ')" }
    if ($removed.Count -gt 0) { $parts += "$($removed.Count) scripts eliminados del servidor" }
    if ($updated.Count -gt 0) { $parts += "$($updated.Count) scripts con versión más nueva en servidor" }
    $msg = if ($parts.Count -gt 0) { "Sincronización completada. $($parts -join '. ')." } else { "Sincronización completada. No había cambios." }
    if ($errors.Count -gt 0) { $msg += " Advertencias: $($errors -join '; ')." }

    script:Write-ServerLog $msg

    return @{
        Success           = $true
        Message           = $msg
        Copied            = $copied.ToArray()
        Errors            = $errors.ToArray()
        RemovedFromServer = $removed.ToArray()
        UpdatedOnServer   = $updated.ToArray()
    }
}

#==================================================================
# BLOQUE: Exportación
#==================================================================
Export-ModuleMember -Function @(
    'Register-ServerLogCallback',
    'Get-AppSettingsPath',
    'Get-DefaultAppSettings',
    'Get-AppSettings',
    'Save-AppSettings',
    'Get-AppSettingValue',
    'Test-ServerAvailable',
    'Get-ServerBase',
    'Get-SQLiteConnString',
    'Open-SQLiteConnection',
    'Get-ActiveComputerDBPath',
    'Get-ScriptsDbContent',
    'Get-AppsDbContent',
    'Save-ScriptsDbLocal',
    'Save-AppsDbLocal',
    'Save-ScriptsDbBoth',
    'Save-AppsDbBoth',
    'Sync-ScriptsJsonToServer',
    'Sync-AppsJsonToServer',
    'Sync-ComputerDbToServer',
    'Publish-ScriptToServer',
    'Publish-IconToServer',
    'Remove-ScriptFromServer',
    'Sync-IconFromServer',
    'Sync-ScriptFromServer',
    'Invoke-FullSyncFromServer'
)



