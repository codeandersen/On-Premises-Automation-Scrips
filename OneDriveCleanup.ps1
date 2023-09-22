<#
        .SYNOPSIS
        Cleanup OneDrive folders after migration to different tenant

        .DESCRIPTION
        Script is used for unlinking OneDrive account and cleanup OneDrive folder after tenant migration .

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
$Company = "Firma" # Change this to the company name that's in the OneDrive folder path
$rootPath = "$env:USERPROFILE" + "\"  
$CurrentDateTime = Get-Date -Format "MM-dd-yyyy_HH_mm"
$LogFile = "$env:LOCALAPPDATA\Temp\" + "$CurrentDateTime" + "_OneDriveCleanup.log"
$RegistryPath = "HKCU:\Software\Microsoft\OneDrive\Accounts\"
$OneDriveExe = "$env:ProgramFiles\" + "Microsoft OneDrive\OneDrive.exe"
$sidToCheck = "S-1-1-0"  # SID for "Everyone" group
$EveryoneGroup = New-Object System.Security.Principal.SecurityIdentifier($sidToCheck) 
$EveryoneGroupName = $EveryoneGroup.Translate([System.Security.Principal.NTAccount]).Value #Conversion of SID to name

#Initialize LogFile
Write-Output "$CurrentDateTime Running OneDrive Cleanup script." >> $LogFile

# Initialize folder search for company OneDrive
$folders = Get-ChildItem -Path $rootPath -Directory

# Initialize registry search for company OneDrives
$Subkeys = Get-ChildItem -Path $RegistryPath


try {
    
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
    }   
    #OneDrive folder cleanup
    Write-Output "Searching for OneDrive folder" >> $LogFile
    $result = $folders | Where-Object { $_.Name -like "*OneDrive*" -and $_.Name -like "*$Company*" }

    if ($result.Count -gt 0) {
        Write-Output "Starting OneDrive folder cleanup process....." >> $LogFile
        Write-Output "Check if everyone deny permission is present....." >> $LogFile

        # Get the current ACL for the folder
        $acl = Get-Acl -Path "$($result.FullName)"
        # Check ACL for Deny entry and everyone group
        $denyPermissions = $acl.Access | Where-Object { $_.AccessControlType -eq "Deny" -and $_.IdentityReference.Value -eq $EveryoneGroupName }
        if ($denyPermissions.Count -gt 0) {
            # Remove all deny permissions for the specified SID
            foreach ($denyPermission in $denyPermissions) {
                $acl.RemoveAccessRule($denyPermission)
            }
        
            # Set the modified ACL back to the folder
            Set-Acl -Path "$($result.FullName)" -AclObject $acl
            Write-Output  "All deny permissions for  '$EveryoneGroupName' have been removed from '$folderPath'." >> $LogFile
        }
        else {
            Write-Output  "No deny permissions found for '$EveryoneGroupName' in `"$($result.FullName)`" ." >> $LogFile
        }

        #OneDrive folder deletion
        $result | ForEach-Object {
            Write-Output "Deleting Onedrive folder..... `"$($result.FullName)`"" >> $LogFile
            Remove-Item -Recurse -Force -Path  "$($result.FullName)" >> $LogFile
        }
    }
    else {
        Write-Output "No matching OneDrive folders found." >> $LogFile
    }
}
catch {
    Write-Output "Error: There was an error running the script error is: `"$($result.FullName)`"" >> $LogFile
}