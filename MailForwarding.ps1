<#
        .SYNOPSIS
        Forwards mail from a disabled user to another from a CSV file

        .DESCRIPTION
        Script is used for automating the mail forwarding of user account in Exchange Online when employess leaves a company .

        .PARAMETER Name
        -CSVFile
            Specifies the location of the Csv file to be imported for mail forwarding user account.
            
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
        C:\PS> MailForwarding.ps1 -Csvfile "C:\temp\Userforwarding.csv"

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
    $LogFile = "$($env:systemdrive)\logs\Mailforwarding_log.txt"
    $SmtpServer = "serv09.tnm.local"
    $MailFrom = "ad@tmg.dk"
    $MailTo = "hca@apento.com"
    $MailSubject = "Mail forwarding of disabled users report"
    $NewFileNameAfterJob = "UserDeak.old"

    #Import Active Directory modules.
    import-module ActiveDirectory

    # Import Exchange Online CMDlets
    $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
    . "$MFAExchangeModule"

    # Connect to Exchange Online
    $AdminUserName = "Scripts@toemmergaarden.onmicrosoft.com"
    $password = Get-Content C:\Scripts\AdminCred-SetOutOfOfficeDisabledUsers.txt | ConvertTo-SecureString
    $Credentials = New-Object -typename System.Management.Automation.PSCredential -ArgumentList $AdminUserName,$password
    Connect-EXOPSSession -Credential $Credentials

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
            $csv = Import-Csv "$csvfile" -Header UserLogonName,Ledermail -Delimiter ";" | Select-Object -Skip 1
    
            #Loops through every user and deactivates them
            ForEach($item in $csv)
            {
            $UserPrincipalName = $($item.UserLogonName)
            $UserLederMail = $($item.Ledermail)
            Get-ADUser -Filter{UserPrincipalName -eq $UserPrincipalName} | Set-Mailbox -ForwardingSMTPAddress "$UserLederMail"
            write-output "User: $UserPrincipalName mail has been forwarded to $UserLederMail"     
            }
    
            #Email is sent with information about users that have been deactivated
            Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject "$MailSubject" -Body "The following users in the attached file mail have been forwarded" -Attachments "$csvfile" -SmtpServer $SmtpServer -UseSsl
            Remove-Item -Path "$NewFileNameAfterJob" -Confirm:$false -Force
            Rename-Item -Path $csvfile -NewName "$NewFileNameAfterJob"
        }

        #Catch if deactivating failed. Logged to file and email sent.
        catch {
            write-output "Error: Executed the mail forwarding user script on $((get-date).DateTime) with the error $_"
            Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject 'Error the mail forwarding user script' -Body "Error: Executed the mail forwarding user script on $((get-date).DateTime) with the error $_" -SmtpServer $SmtpServer -UseSsl      
}
    

#Stopping transcript
stop-transcript;

exit