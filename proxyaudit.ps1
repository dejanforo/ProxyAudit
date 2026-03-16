<#
.SYNOPSIS
    Audits proxy settings of all Exchange servers.

.DESCRIPTION
    This script audits the proxy settings of all Exchange servers in the environment.
    It collects both Exchange-specific proxy settings and WinHTTP proxy settings 
    and displays the results in a formatted table and CSV file.

.EXAMPLE
    .\proxyaudit.ps1
    This command runs the proxy audit and outputs the results to the console and a CSV file.
    The script does not require any parameters and will automatically gather information from all Exchange servers.

    The output of the script will look like this:

    Timestamp           ServerName Role    Status             ExchProxy                       ExchExceptions                         WinHTTPProxy           WinHTTPBypass
    ---------           ---------- ----    ------             ---------                       --------------                         ------------           -------------
    2026-03-13 15:25:53 SERVER01   Mailbox Online             http://proxy.domain.com:8080/   mgmt.net, intranet.com, domainA.com    proxy.domain.com:8080  *.mgmt.net;*.intranet.com;domainA.com;<local>
    2026-03-13 15:25:57 SERVER02   Edge    Unreachable (Edge) None                            None                                   N/A                    N/A
    2026-03-13 15:25:59 SERVER03   Edge    Unreachable (Edge) None                            None                                   N/A                    N/A
    2026-03-13 15:26:00 SERVER04   Mailbox Online             http://proxy.domain.com:8080/   mgmt.net, intranet.com, domainA.com    proxy.domain.com:8080  *.mgmt.net;*.intranet.com;domainA.com;<local>
    2026-03-13 15:26:00 SERVER05   Mailbox Online             http://proxy.domain.com:8080/   mgmt.net, intranet.com, domainA.com    proxy.domain.com:8080  *.mgmt.net;*.intranet.com;domainA.com;<local>

    Audit complete. Results saved to: D:\Scripts\proxyaudit\ProxyAudit_20260313_1526.tsv

.NOTES
    Author: Dejan Foro, Exchangemaster GmbH  
            e-mail: dejan.foro@exchangemaster.ch  
            web: https://www.exchangemaster.ch  
            Github: https://github.com/dejanforo  
            LinkedIn: https://www.linkedin.com/in/dejanforo/  
    Date: March 13, 2026
    Version: 7.0
    Requires: PowerShell 5.1 or 7.x and Exchange PSsnap-in installed.
#>

Clear-Host

if (-not (Get-Command "Get-ExchangeServer" -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Snapin -ErrorAction SilentlyContinue
}

# --- Improved Audit Logic ---
$AuditBlock = {
    param($ServerName, $Roles, $ExchProxy, $ExchExceptions)

    $WinHTTPRaw = netsh winhttp show proxy
    
    # Use Regex to grab everything after the " : " to handle URLs correctly
    $WinHTTPProxy = ($WinHTTPRaw | Where-Object { $_ -match "Proxy Server\(s\)\s+:\s+(.*)" })
    if ($Matches[1]) { $WinHTTPProxy = $Matches[1].Trim() } else { $WinHTTPProxy = "Direct Access" }

    $WinHTTPBypass = ($WinHTTPRaw | Where-Object { $_ -match "Bypass List\s+:\s+(.*)" })
    if ($Matches[1]) { $WinHTTPBypass = $Matches[1].Trim() } else { $WinHTTPBypass = "None" }

    [PSCustomObject]@{
        Timestamp       = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        ServerName      = $ServerName
        Role            = $Roles -join ", "
        Status          = "Online"
        ExchProxy       = $ExchProxy
        ExchExceptions  = $ExchExceptions
        WinHTTPProxy    = $WinHTTPProxy
        WinHTTPBypass   = $WinHTTPBypass
    }
}

Write-Host "Gathering server list..." -ForegroundColor Green
$ExchangeServers = Get-ExchangeServer | Sort-Object Name

$AuditResults = foreach ($Server in $ExchangeServers) {
    Write-Host "Processing $($Server.Name)..." -ForegroundColor Green
    $Roles = $Server.ServerRole
    $TargetFQDN = $Server.Fqdn
    
    $ExchProxy = if ($Server.InternetWebProxy) { $Server.InternetWebProxy.AbsoluteUri } else { "None" }
    $ExchExceptions = if ($Server.InternetWebProxyBypassList) { 
        ($Server.InternetWebProxyBypassList | ForEach-Object { $_.ToString() }) -join ", " 
    } else { "None" }

    $IsLocal = ($TargetFQDN -eq $env:COMPUTERNAME) -or 
               ($TargetFQDN -eq "$($env:COMPUTERNAME).$($env:USERDNSDOMAIN)") -or
               ($Server.Name -eq $env:COMPUTERNAME)

    if ($IsLocal) {
        & $AuditBlock $Server.Name $Roles $ExchProxy $ExchExceptions
    }
    elseif (Test-WSMan -ComputerName $TargetFQDN -ErrorAction SilentlyContinue) {
        Invoke-Command -ComputerName $TargetFQDN -ScriptBlock $AuditBlock -ArgumentList $Server.Name, $Roles, $ExchProxy, $ExchExceptions -ErrorAction SilentlyContinue
    }
    else {
        [PSCustomObject]@{
            Timestamp       = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            ServerName      = $Server.Name
            Role            = $Roles -join ", "
            Status          = if ($Roles -like "*Edge*") { "Unreachable (Edge)" } else { "Unreachable" }
            ExchProxy       = $ExchProxy
            ExchExceptions  = $ExchExceptions
            WinHTTPProxy    = "N/A"
            WinHTTPBypass   = "N/A"
        }
    }
}

Clear-Host 

# Display results cleanly (excluding technical remote session properties)
$AuditResults | Select-Object * -ExcludeProperty PSComputerName, RunspaceId | Format-Table -AutoSize

# Save to CSV
$FilePath = Join-Path $PSScriptRoot "ProxyAudit_$((Get-Date).ToString('yyyyMMdd_HHmm')).tsv"
$AuditResults | Select-Object * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -Delimiter "`t"

Write-Host "`nAudit complete. Results saved to: $FilePath" -ForegroundColor Cyan

# End of script