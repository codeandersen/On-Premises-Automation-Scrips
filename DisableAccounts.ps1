    <#
        .SYNOPSIS
        Disabling user account from importing a CSV file

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
        C:\PS> DisableAccounts.ps1 -Csvfile "C:\temp\DisableAccounts.csv"

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
    $LogFile = "$($env:windir)\temp\DisableAccounts_log.txt"
    $SmtpServer = "msonline-dk.mail.protection.outlook.com"
    $MailFrom = "deaktivering@msonline.dk"
    $MailTo = "hca@apento.com"
    $MailSubject = "Deativation of user report"
    
    #Import Active Directory modules.
    import-module ActiveDirectory

    #Starting transcript
    start-transcript -Path $LogFile
    
    try {
        #Imports the Csv file
        $csv = Import-Csv "$csvfile" -Header UserLogonName

        #Loops through every user and deactivates them
        ForEach($item in $csv)
        {
        $UserLogonName = $($item.UserLogonName)
        #Disable-ADAccount -Identity $UserLogonName
        write-output "User: $UserLogonName has been disabled from Active Directory"        
        }

        #Email is sent with information about users that have been deactivated
        Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject "$MailSubject" -Body "The following users in the attached file has been disabled" -Attachments "$csvfile" -SmtpServer $SmtpServer -UseSsl

    }

    #Catch if deactivating failed. Logged to file and email sent.
    catch {
        write-output "Error: Executed the disable user accounts script on $((get-date).DateTime) with the error $_"
        Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject 'Error running deactivation script' -Body "Error: Executed the disable user accounts script on $((get-date).DateTime) with the error $_" -SmtpServer $SmtpServer -UseSsl
    }
        

    #Stopping transcript
    stop-transcript;
    
    exit