<#
        .SYNOPSIS
        Cleanup OneDrive folders after migration to different tenant

        .DESCRIPTION
        Script is used for automating the decommisiong of user account in Active Directory when employess leaves a company .

        .EXAMPLE
        C:\PS> OneDriveCleanup.ps1

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        BLOG: http://www.hcconsult.dk
        Twitter: @dk_hcandersen

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>


#Parameters declaration
$Company = "HC Consult" # Change this to the company name that's in the OneDrive folder path
$rootPath = "$env:USERPROFILE" + "\"  # Change this to the drive you want to search
$CurrentDateTime = Get-Date -Format "MM-dd-yyyy_HH_mm"
$LogFile = "$env:LOCALAPPDATA\Temp\" + "$CurrentDateTime" + "_OneDriveCleanup.log"
$RegistryPath = "HKCU:\Software\Microsoft\OneDrive\Accounts\"
$OneDriveExe = "$env:ProgramFiles\" + "Microsoft OneDrive\OneDrive.exe"

#Initialize LogFile
Write-Output "$CurrentDateTime Running OneDrive Cleanup script." >> $LogFile

# Initialize folder search for Cubic OneDrive
$folders = Get-ChildItem -Path $rootPath -Directory

# Initialize registry search for Cubic OneDrives
$Subkeys = Get-ChildItem -Path $RegistryPath


#Write-Output "Folders found ""$result""" >> $LogFile


try {

    #Unlinking OneDrive account that contains Cubic
    foreach ($Subkey in $Subkeys) {
        $DisplayName = (Get-ItemProperty -Path $Subkey.PSPath -Name "Displayname" -ErrorAction SilentlyContinue).Displayname
    
        if ($DisplayName -like "*$Company*") {
            $FullOneDrivePath = $RegistryPath + $($Subkey.PSChildName)
    
            Write-Output "Removing OneDrive Account $Company in  $FullOneDrivePath" >> $LogFile
            Write-Output "Stopping OneDrive process before removal" >> $LogFile
            Stop-Process -name "Onedrive" -Force 
            Remove-Item -Path $FullOneDrivePath -Recurse -Force -Confirm:$false 
            Write-Output "Starting OneDrive" >> $LogFile
            Start-Process -FilePath "$OneDriveExe" 
            Start-Sleep -Seconds 5
            Start-Process -FilePath "$OneDriveExe" -ArgumentList "/configure_business"
            Start-Sleep -Seconds 5
            Stop-Process -name "Onedrive" -Force 
            Start-Process -FilePath "$OneDriveExe"
            Write-Output "OneDrive account $Company in  $FullOneDrivePath removed" >> $LogFile
    }
}   
    #OneDrive folder cleanup
    Write-Host "Searching for OneDrive folder" >> $LogFile
    $result = $folders | Where-Object {$_.Name  -like "*OneDrive*" -and $_.Name  -like "*$Company*" }

    if ($result.Count -gt 0) {
        Write-Output "Starting OneDrive folder cleanup process....." >> $LogFile
        Start-Sleep -Seconds 30
        $result | ForEach-Object {
        Write-Output "Deleting Onedrive folder ""$_""" >> $LogFile
        Remove-Item -Recurse -Force $_ >> $LogFile
        }
    } else {
        Write-Output "No matching OneDrive folders found." >> $LogFile
    }
}
catch {
    Write-Output "Error: There was an error running the script error is: $_" >> $LogFile
}

