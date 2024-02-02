<#
        .SYNOPSIS
        Sync-SMTPRelayConnectorIPs.ps1

        .DESCRIPTION
        Script is used to synchronize receive connector IP addresses across all Exchange servers.

        .EXAMPLE
        C:\PS> Sync-SMTPRelayConnectorIPs.ps1

       .LINK
        BLOG: http://www.hcconsult.dk
        Github: https://github.com/codeandersen
        LinkedIn: https://www.linkedin.com/in/hanschrandersen/
        Twitter: @dk_hcandersen

        .CHANGELOG
        V1.00, 02/02/2024 - Initial version

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.
    #>

   #Parameters declaration
   $LogFile = "$($env:systemdrive)\scripts\Sync-SMTPRelayConnectorIPs\Sync-SMTPRelayConnectorIPs.log"
   $RelayConnector = "SMTP Relay"   
   $ExchServers = @("EXCH1","EXCH2")
   $FQDN = "FQDN.Doamin.com" 

#Starting transcript
Start-transcript -Path $LogFile


#Connect to Exchange Server
ForEach ($ExchServer in $ExchServers) {
    if($session){break}
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchServer.$FQDN/PowerShell/ -Authentication Kerberos
    $Connection = Import-PSSession $Session -DisableNameChecking
}

# Function to get "SMTP Relay" receive connectors from all Exchange servers
function Sync-SMTPRelayConnectorIPs {
    # Get a list of all Exchange servers
    $exchangeServers = Get-ExchangeServer

    # Identify the "SMTP Relay" connector with the latest change
    $latestConnectorInfo = $exchangeServers | ForEach-Object {
        Get-ReceiveConnector -Server $_.Name | Where-Object { $_.Name -eq $RelayConnector } |
        Select-Object @{Name='Server';Expression={$_.Server}}, Identity, WhenChanged, RemoteIPRanges
    } | Sort-Object WhenChanged -Descending | Select-Object -First 1

    if (-not $latestConnectorInfo) {
        Write-Host "No 'SMTP Relay' connector found."
        return
    }

    Write-Host "Synchronizing IP addresses based on the latest 'SMTP Relay' connector on $($latestConnectorInfo.Server)"

    # Synchronize IP addresses on all "SMTP Relay" connectors
    foreach ($server in $exchangeServers) {
        $connector = Get-ReceiveConnector -Server $server.Name | Where-Object { $_.Name -eq "SMTP Relay" }
        if ($connector) {
            # Compare and update IP ranges if necessary
            $currentIPRanges = $connector.RemoteIPRanges
            $latestIPRanges = $latestConnectorInfo.RemoteIPRanges

            $ipsToAdd = $latestIPRanges | Where-Object { $_ -notin $currentIPRanges }
            $ipsToRemove = $currentIPRanges | Where-Object { $_ -notin $latestIPRanges }

            if ($ipsToAdd) {
                Write-Host "Adding new IP ranges to $($connector.Identity): $($ipsToAdd -join ', ')"
                Set-ReceiveConnector -Identity $connector.Identity -RemoteIPRanges ($currentIPRanges + $ipsToAdd)
            }
            
            if ($ipsToRemove) {
                Write-Host "Removing outdated IP ranges from $($connector.Identity): $($ipsToRemove -join ', ')"
                $updatedIPRanges = $currentIPRanges | Where-Object { $_ -notin $ipsToRemove }
                Set-ReceiveConnector -Identity $connector.Identity -RemoteIPRanges $updatedIPRanges
            }            

            if (-not $ipsToAdd -and -not $ipsToRemove) {
                Write-Host "No changes required for $($connector.Identity)"
            }
        } else {
            Write-Host "No 'SMTP Relay' connector found on server $($server.Name)"
        }
    }
}

# Execute the synchronization function
Sync-SMTPRelayConnectorIPs


#Stopping transcript
Stop-transcript;
Exit