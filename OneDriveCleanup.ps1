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
$Company = "Cubic" # Change this to the company name that's in the OneDrive folder path
$rootPath = "$env:USERPROFILE" + "\"  # Change this to the drive you want to search
$CurrentDateTime = Get-Date -Format "MM-dd-yyyy_HH_mm"
$LogFile = "$env:LOCALAPPDATA\Temp\" + "$CurrentDateTime" + "_OneDriveCleanup.log"

#Initialize LogFile
Write-Output "$CurrentDateTime Running OneDrive Cleanup script." >> $LogFile

# Use Get-ChildItem to list directories at the root of the specified drive
$folders = Get-ChildItem -Path $rootPath -Directory

# Use Where-Object to filter the folders that contain both "onedrive" and "cubic" in their names
Write-Host "Searching for OneDrive folder to delete" >> $LogFile
$result = $folders | Where-Object {$_.Name  -like "*OneDrive*" -and $_.Name  -like "*$Company*" }
Write-Output "Folders found ""$result""" >> $LogFile


try {
    if ($result.Count -gt 0) {
        Write-Output "Starting OneDrive Cleanup process....." >> $LogFile
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