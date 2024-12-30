
<#
        .SYNOPSIS
        Gets PC information

        .DESCRIPTION
        Script is used for extracting information from a computer to use in an inventory system.

        .EXAMPLE
        C:\PS> PC_Inventory.ps1"

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        BLOG: http://www.hcconsult.dk
        Twitter: @dk_hcandersen

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>

$PC = Get-ComputerInfo
$User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$Nic = Get-NetAdapter -Physical
$DateTime = Get-Date -Format "HH:mm:ss dd.MM.yyyy"

#Monitor information
Add-Type -AssemblyName System.Windows.Forms
$Screens = [System.Windows.Forms.Screen]::AllScreens
$MonitorDetailsString = @()
foreach ($Screen in $Screens) {
    $MonitorDetailsString += "$($Screen.DeviceName): $($Screen.Bounds.Width) x $($Screen.Bounds.Height)"
}
$MonitorDetailsJoined = $MonitorDetailsString -join ", "

# Join multiple MAC addresses into a single string
$MACAddresses = ($nic.MacAddress -join ", ")

# Join multiple MAC addresses into a single string
$MonitorResolution = ($MonitorResolution -join ", ")

# Create a custom object with the required fields
$ExportObject = [PSCustomObject]@{
    ComputerName             = $PC.CsName
    UserName                 = $User
    ProcessorName            = $PC.CsProcessors.Name
    TotalPhysicalMemoryMB    = [Math]::Round($PC.CsTotalPhysicalMemory / 1MB, 0)
    ProcessorCores           = $PC.CsProcessors.NumberOfCores
    DriveCSizeGB             = [Math]::Round((Get-Volume -DriveLetter C).Size / 1GB, 0)
    PCSystemType             = $PC.CsPCSystemType
    Model                    = $PC.CsModel
    OSName                   = $PC.OsName
    OSVersion                = $PC.OsVersion
    OSBuild                  = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
    MACAddress               = $MACAddresses
    IPv4DefaultGateway       = $Nic.IPv4DefaultGateway.NextHop
    NumberOfMonitors         = [System.Windows.Forms.Screen]::AllScreens.Count
    MonitorsResolution       = $MonitorDetailsJoined
    DateTime                 = $DateTime
}



# Construct the export file name
$Computername = $PC.CsName
$DateTime = Get-Date -Format "yyyyMMddHHmmss"
$OutputCSVName = "$Computername" + "_" + $DateTime + ".csv"
$OutputfilePath = "\\AD.local\System\Info\Information\"
#$OutputfilePath = "C:\Github\On-Premises-Automation-Scripts\"
$OutputCSV = "$OutputfilePath" + "$OutputCSVName"

# Export to CSV
$ExportObject | Export-Csv -Path "$OutputCSV" -NoTypeInformation -Encoding UTF8 -Delimiter ";" 