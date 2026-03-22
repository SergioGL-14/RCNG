<#
    Consultas DHCP por RSAT.
    Se usa como apoyo desde la app principal cuando hace falta buscar leases
    o reservas sin depender de WinRM.
#>

$script:DHCP_Server = 'DHCP-SERVER'

function script:Get-DefaultDhcpServer {
    if (Get-Command 'Get-AppSettingValue' -ErrorAction SilentlyContinue) {
        return (Get-AppSettingValue -Key 'DhcpServer' -DefaultValue $script:DHCP_Server)
    }
    return $script:DHCP_Server
}

# Busca leases en el servidor DHCP filtrando por nombre, IP o MAC.
# Recorre todos los scopes desde RSAT usando -ComputerName, sin WinRM.
function DHCP_QueryDirect {
    param(
        [string]$Filter,
        [string]$DHCPServer = $(script:Get-DefaultDhcpServer)
    )
    try {
        $results = @()
        $scopes = Get-DhcpServerv4Scope -ComputerName $DHCPServer -ErrorAction Stop
        foreach ($scope in $scopes) {
            $leases = Get-DhcpServerv4Lease -ComputerName $DHCPServer -ScopeId $scope.ScopeId -ErrorAction SilentlyContinue
            foreach ($l in $leases) {
                $match = ($l.HostName  -like "*$Filter*") -or
                         ($l.IPAddress -like "*$Filter*") -or
                         ($l.ClientId  -like "*$Filter*")
                if ($match) {
                    $results += [PSCustomObject]@{
                        HostName        = $l.HostName
                        IPAddress       = ($l.IPAddress -as [string])
                        ClientId        = $l.ClientId
                        ScopeId         = ($l.ScopeId -as [string])
                        LeaseExpiryTime = ($l.LeaseExpiryTime -as [string])
                        AddressState    = $l.AddressState
                    }
                }
            }
        }
        if ($results.Count -gt 0) { return $results }
        return $null
    } catch {
        return $null
    }
}

Export-ModuleMember -Function DHCP_QueryDirect


