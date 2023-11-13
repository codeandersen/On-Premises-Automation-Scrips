<#
        .SYNOPSIS
        Cleanup OneDrive known folders backup after failed migration.

        .DESCRIPTION
        Script is used for cleaning up after failed known folder move. It will move files that was accidently move to OneDrive back to local desktop, document and picture folder. 
        The it will delete the desktop, document and picture folder from OneDrive.
        Last it will setup OneDrive for known folder backup.

        .EXAMPLE
        C:\PS> OneDriveKFMClean.ps1

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
$rootPath = "$env:USERPROFILE" + "\"  
$CurrentDateTime = Get-Date -Format "MM-dd-yyyy_HH_mm"
$LogFile = "$env:LOCALAPPDATA\Temp\" + "$CurrentDateTime" + "_OneDriveKFMCleanup.log"
$LogFileRobocopy = "$env:LOCALAPPDATA\Temp\" + "$CurrentDateTime" + "_OneDriveKFMCleanupRobocopy"
$RegistryPath = "HKCU:\Software\Microsoft\OneDrive\Accounts\"
$OneDriveExe = "$env:ProgramFiles\" + "Microsoft OneDrive\OneDrive.exe"
#$sidToCheck = "S-1-1-0"  # SID for "Everyone" group
#$EveryoneGroup = New-Object System.Security.Principal.SecurityIdentifier($sidToCheck) 
#$EveryoneGroupName = $EveryoneGroup.Translate([System.Security.Principal.NTAccount]).Value #Conversion of SID to name

#Initialize LogFile
Write-Output "$CurrentDateTime Running OneDrive KFM Backup Cleanup script." >> $LogFile

# Initialize folder search for Company OneDrive
$folders = Get-ChildItem -Path $rootPath -Directory

# Initialize registry search for company OneDrives
$Subkeys = Get-ChildItem -Path $RegistryPath


try {
    
    <#
    #Unlinking OneDrive account that contains company name variable
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
    } #>  

    #OneDrive folder cleanup
    Write-Output "Searching for OneDrive folder" >> $LogFile
    $result = $folders | Where-Object { $_.Name -like "*OneDrive*" -and $_.Name -like "*$Company*" }

    if ($result.Count -gt 0) {
        Write-Output "Starting OneDrive folder cleanup process....." >> $LogFile
        
        robocopy "$($result.FullName)\Documents" "$env:userprofile\Documents" /E /DCOPY:DAT /XO /R:100 /W:3 /LOG:"$env:LOCALAPPDATA\Temp\robocopylog_doc_sw.txt"
        robocopy "$($result.FullName)\Desktop" "$env:userprofile\Desktop" /E /DCOPY:DAT /XO /R:100 /W:3 /LOG:"$env:LOCALAPPDATA\Temp\robocopylog_desk_sw.txt"
        robocopy "$($result.FullName)\Pictures" "$env:userprofile\Pictures" /E /DCOPY:DAT /XO /R:100 /W:3 /LOG:"$env:LOCALAPPDATA\Temp\robocopylog_pic_sw.txt"

        pause

        #OneDrive folder cleanup
        $result | ForEach-Object {
            Write-Output "Deleting Onedrive folder: `"$($result.FullName)`"" >> $LogFile
            Remove-Item -Recurse -Force -Path  "$($result.FullName)\Documents" >> $LogFile
            Remove-Item -Recurse -Force -Path  "$($result.FullName)\Desktop" >> $LogFile
            Remove-Item -Recurse -Force -Path  "$($result.FullName)\Pictures" >> $LogFile
        } 
    }
    else {
        Write-Output "No matching OneDrive folders found." >> $LogFile
    }
}
catch {
    Write-Output "Error: There was an error running the script error is: `"$($result.FullName)`"" >> $LogFile
}