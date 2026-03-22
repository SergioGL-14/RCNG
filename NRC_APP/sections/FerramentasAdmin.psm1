# Acciones del menu Administrador.
# Reune accesos a consolas del sistema y enlaces web configurables del entorno.
param()

function Invoke-Ferramentas_CommandPrompt {
    Start-Process cmd.exe -Verb runas
}

function Invoke-Ferramentas_Powershell {
    Start-Process powershell.exe -Verb runas
}

function Invoke-Ferramentas_PowershellISE {
    Start-Process powershell_ise.exe
}

function Invoke-Ferramentas_taskManager {
    Start-Process taskmgr
}

function Invoke-Ferramentas_services {
    Start-Process services.msc
}

function Invoke-Ferramentas_regedit {
    Start-Process regedit
}

function Invoke-Ferramentas_mmc {
    Start-Process mmc
}

function Invoke-Ferramentas_InternetExplorer {
    $url = 'https://portal.example.local'
    if (Get-Command 'Get-AppSettingValue' -ErrorAction SilentlyContinue) {
        $url = Get-AppSettingValue -Key 'PortalUrl' -DefaultValue $url
    }
    Start-Process "microsoft-edge:$url" -WindowStyle maximized
}

function Invoke-Ferramentas_GeneratePassword {
    $url = 'https://mail.example.local'
    if (Get-Command 'Get-AppSettingValue' -ErrorAction SilentlyContinue) {
        $url = Get-AppSettingValue -Key 'MailPortalUrl' -DefaultValue $url
    }
    Start-Process "microsoft-edge:$url" -WindowStyle maximized
}

function Invoke-Ferramentas_compmgmt {
    Get-ComputerTxtBox
    if ($Global:ComputerName -and ($Global:ComputerName -notmatch "(?i)^(localhost|\.|127\.0\.0\.1|$env:COMPUTERNAME)$")) {
        Start-Process compmgmt.msc "/computer:$Global:ComputerName"
    } else {
        Start-Process compmgmt.msc
    }
}

function Invoke-Ferramentas_DHCP {
    Get-ComputerTxtBox
    if ($Global:ComputerName -match "(?i)^(localhost|\.|127\.0\.0\.1|$env:COMPUTERNAME)$") {
        Start-Process dhcpmgmt.msc
    } else {
        Start-Process dhcpmgmt.msc "/computer:$Global:ComputerName"
    }
}

function Invoke-Ferramentas_TerminalAdmin {
    $mscPath = "C:\\Program Files\\Update Services\\AdministrationSnapin\\wsus.msc"
    if (Test-Path $mscPath) {
        Start-Process -FilePath $mscPath
    } else {
        Start-Process "mmc"
    }
}

function Invoke-Ferramentas_ADSearchDialog {
    Start-Process -FilePath "C:\\Windows\\system32\\dsa.msc"
}

function Invoke-Ferramentas_ADPrinters {
    Start-Process -FilePath "C:\\Windows\\system32\\printmanagement.msc"
}

function Invoke-Ferramentas_adExplorer {
    Add-Logs -text "Localhost - SysInternals AdExplorer"
    $command = Join-Path $Global:ScriptRoot 'tools\AdExplorer\AdExplorer.exe'
    Start-Process $command -WorkingDirectory (Join-Path $Global:ScriptRoot 'tools')
}

Export-ModuleMember -Function *
