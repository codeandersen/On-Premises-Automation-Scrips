    <#
        .SYNOPSIS
        Deletes user account from importing a CSV file

        .DESCRIPTION
        Script is used for automating the decommisiong of user account in Active Directory when employess leaves a company .

        .PARAMETER Name
        -CSVFile
            Specifies the location of the Csv file to be imported for disabling accounts.
            
            Required?                    true
               Position?                    0
               Default value
               Accept pipeline input?       false
               Accept wildcard characters?

        .PARAMETER Extension
        Specifies the extension. "Csv" is the default.

               Required?                    True
               Position?                    1
               Default value
               Accept pipeline input?       false
               Accept wildcard characters?

        .EXAMPLE
        C:\PS> DeleteAccounts.ps1 -Csvfile "C:\temp\DeleteAccounts.csv"

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        BLOG: http://www.hcconsult.dk
        Twitter: @dk_hcandersen

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>
    Param(
        [Parameter(Mandatory=$True,
        Position=0,
        HelpMessage="Enter path and filename of Csv file")]
        [String]$csvfile
        )

    #Parameters declaration
    $LogFile = "$($env:systemdrive)\logs\DeleteAccounts_log.txt"
    $SmtpServer = "serv09.tnm.local"
    $MailFrom = "ad@tmg.dk"
    $MailTo = "hca@apento.com"
    $MailSubject = "Deletion of user report"
    $NewFileNameAfterJob = "UserSlet.old"
    
    #Import Active Directory modules.
    import-module ActiveDirectory

    #Starting transcript
    start-transcript -Path $LogFile
    
    If(!(Test-Path $csvfile))
        {
            write-output "Error: CSV file doesn't exist. Executed on $((get-date).DateTime)"
            #Stopping transcript
            stop-transcript;
            Exit
        }
    
    try {
        #Imports the Csv file
        $csv = Import-Csv "$csvfile" -Header UserLogonName -Delimiter ";" | Select-Object -Skip 1

        #Loops through every user and deactivates them
        ForEach($item in $csv)
        {
        $UserPrincipalName = $($item.UserLogonName)
        Get-ADUser -Filter{UserPrincipalName -eq $UserPrincipalName} | Remove-ADUser -Confirm:$False
        write-output "User: $UserPrincipalName has been deleted from Active Directory"     
        }

        #Email is sent with information about users that have been deactivated
        Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject "$MailSubject" -Body "The following users in the attached file has been deleted" -Attachments "$csvfile" -SmtpServer $SmtpServer -UseSsl
        Rename-Item -Path $csvfile -NewName "$NewFileNameAfterJob"

    }

    #Catch if deactivating failed. Logged to file and email sent.
    catch {
            write-output "Error: Executed the delete user account script on $((get-date).DateTime) with the error $_"
            Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject 'Errror running deletion of user account script' -Body "Error: Executed the delete user accounts script on $((get-date).DateTime) with the error $_" -SmtpServer $SmtpServer -UseSsl      
    }
        

    #Stopping transcript
    stop-transcript;
    
    exit