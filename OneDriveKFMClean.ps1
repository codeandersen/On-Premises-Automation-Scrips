<#
        .SYNOPSIS
        Cleanup OneDrive known folders backup after failed migration.

        .DESCRIPTION
        Script is used for cleaning up after failed known folder move. It will move files that was accidently move to OneDrive back to local desktop, document and picture folder. 
        Then it will delete the desktop, document and picture folder from OneDrive.

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
$Company = "Company Name" # Change this to the company name that's in the OneDrive folder path
$CompanyRegName = "STARK" # Used for setting variable after successfull run
$rootPath = "$env:USERPROFILE" + "\"  
$CurrentDateTime = Get-Date -Format "MM-dd-yyyy_HH_mm"
$LogFile = "$env:LOCALAPPDATA\Temp\" + "$CurrentDateTime" + "_OneDriveKFMCleanup.log"
$RegistryPath = "HKCU:\Software\Microsoft\OneDrive\Accounts\"
$OneDriveExe = "$env:ProgramFiles\" + "Microsoft OneDrive\OneDrive.exe"
$Documents = "Dokumenter"
$Pictures = "Billeder"
$Desktop = "Skrivebord"
    
#Initialize LogFile
Write-Output "$CurrentDateTime Running OneDrive KFM Backup Cleanup script." >> $LogFile
    
# Initialize folder search for Company OneDrive
$folders = Get-ChildItem -Path $rootPath -Directory
    
# Initialize registry search for company OneDrives
$Subkeys = Get-ChildItem -Path $RegistryPath
    
    
try {
       
    $status = Get-ItemProperty -Path "HKCU:\$CompanyRegName" -Name "OneDriveCleanupDone" -ErrorAction Continue
        
    #Checks if script has already been executed
    If (!($status)) {
    
        #OneDrive folder cleanup
        Write-Output "Searching for OneDrive folder" >> $LogFile
        $result = $folders | Where-Object { $_.Name -like "*OneDrive*" -and $_.Name -like "*$Company*" }
    
        if ($result.Count -gt 0) {
            Write-Output "Starting OneDrive folder cleanup process....." >> $LogFile
              
            robocopy "$($result.FullName)\$Documents" "$env:userprofile\Documents" /E /DCOPY:DAT /XO /R:100 /W:3 /LOG:"$env:LOCALAPPDATA\Temp\robocopylog_doc_sw.txt"
            robocopy "$($result.FullName)\$Desktop" "$env:userprofile\Desktop" /E /DCOPY:DAT /XO /R:100 /W:3 /LOG:"$env:LOCALAPPDATA\Temp\robocopylog_desk_sw.txt"
            robocopy "$($result.FullName)\$Pictures" "$env:userprofile\Pictures" /E /DCOPY:DAT /XO /R:100 /W:3 /LOG:"$env:LOCALAPPDATA\Temp\robocopylog_pic_sw.txt"
    
            pause
            #OneDrive folder cleanup
            $result | ForEach-Object {
                Write-Output "Deleting Onedrive folder: `"$($result.FullName)`"" >> $LogFile
                Remove-Item -Recurse -Force -Path  "$($result.FullName)\$Documents" >> $LogFile
                Remove-Item -Recurse -Force -Path  "$($result.FullName)\$Desktop" >> $LogFile
                Remove-Item -Recurse -Force -Path  "$($result.FullName)\$Pictures" >> $LogFile
                Start-sleep -Seconds 10
                Stop-Process -Name OneDrive
                New-Item  -Path "HKCU:\$CompanyRegName" -Force
                New-ItemProperty -Path "HKCU:\$CompanyRegName" -Name "OneDriveCleanupDone" -Value "True" -Force
                Stop-Process -Name OneDrive
                Start-sleep -Seconds 5
                Start-Process $OneDriveExe
            }
        }
        else {
            Write-Output "No matching OneDrive folders found." >> $LogFile
        }
    
    
    
    }
    
       
}
catch {
    Write-Output "Error: There was an error running the script error is: `"$($result.FullName)`"" >> $LogFile
}