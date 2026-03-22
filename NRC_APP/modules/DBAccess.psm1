<#
    MÃ³dulo SQLite heredado para ComputerNames.sqlite.
    Se mantiene por compatibilidad, aunque el script principal ya inicializa
    su propia conexiÃ³n para el autocompletado y las bÃºsquedas rÃ¡pidas.
#>

function Initialize-DBConnection {
    param(
        [string]$ScriptRoot = (Split-Path -Parent $MyInvocation.MyCommand.Definition),
        [string]$DllRelative = '..\libs\System.Data.SQLite.dll',
        [string]$DbRelative = '..\database\ComputerNames.sqlite'
    )

    if ($script:DBConn) { return $script:DBConn }

    $dllPath = Join-Path $ScriptRoot $DllRelative
    $dbFile = Join-Path $ScriptRoot $DbRelative

    if (-not (Test-Path $dllPath)) { throw "DLL not found: $dllPath" }
    if (-not (Test-Path $dbFile)) { throw "DB not found: $dbFile" }

    Add-Type -Path $dllPath -ErrorAction Stop
    $connString = "Data Source=$dbFile;Version=3;"
    $conn = New-Object System.Data.SQLite.SQLiteConnection $connString
    $conn.Open()
    $script:DBConn = $conn
    return $script:DBConn
}

function Test-DBAvailable {
    try {
        $c = Initialize-DBConnection
        return $true
    } catch {
        return $false
    }
}

function Get-ComputerByFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Filter,
        [int]$Limit = 200
    )

    $conn = Initialize-DBConnection
    $cmd = $conn.CreateCommand()

   # Búsqueda LIKE sin distinguir mayúsculas, igual que en la app principal.
    $sql = @"
    SELECT ou, equipo, orig_line FROM computers
    WHERE UPPER(equipo) LIKE '%' || UPPER(@f) || '%'
       OR UPPER(ou) LIKE '%' || UPPER(@f) || '%'
    LIMIT @limit;
"@
    $cmd.CommandText = $sql
    $null = $cmd.Parameters.AddWithValue('@f', $Filter)
    $null = $cmd.Parameters.AddWithValue('@limit', [int]$Limit)

    $dt = New-Object System.Data.DataTable
    (New-Object System.Data.SQLite.SQLiteDataAdapter($cmd)).Fill($dt) | Out-Null

    $out = @()
    foreach ($row in $dt.Rows) {
        $obj = [PSCustomObject]@{
            OU       = $row['ou']
            Equipo   = $row['equipo']
            OrigLine = $row['orig_line']
        }
        $out += $obj
    }
    return $out
}

Export-ModuleMember -Function Initialize-DBConnection,Test-DBAvailable,Get-ComputerByFilter


