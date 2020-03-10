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

    start-transcript -Path "$($env:windir)\temp\DisableAccounts_log.txt"

    try {}
    catch {write-output "Executed the disable user accounts script on $((get-date).DateTime) with the error $_" >> "$($env:windir)\temp\DisableAccounts_log.txt"}

    stop-transcript;

    exit