<#
        .SYNOPSIS
        Add user accounts from a CSV file to an Active Directory group

        .DESCRIPTION
        Script is used for automating the addition of users to groups in Active Directory.

        .PARAMETER Name
        -CSVFile
            Specifies the location of the Csv file.
            
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
        C:\PS> AddToGroup.ps1 -Csvfile "C:\temp\UserAccounts.csv"

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
        Position=0
        #HelpMessage="Enter path and filename of CSV file"
        )]
        [String]$csvfile,
        [String]$group
        )


#Parameters declaration
$LogFile = "$($env:systemdrive)\logs\AddToGroup_log.txt"
$SmtpServer = "serv09.tnm.local"
$MailFrom = "ad@tmg.dk"
$MailTo = "hca@apento.com"
$MailSubject = "Add User to AD group report"
$NewFileNameAfterJob = "\\serv06\scripts\AD\Users\UserOpret.old"
$OUTopLevelPath = "OU=Afdelinger,OU=Domain Users,DC=tnm,DC=local"

# Import active directory module for running AD cmdlets
Import-Module activedirectory

  #Starting transcript
  start-transcript -Path $LogFile
    
  If(!(Test-Path $csvfile))
      {
          write-output "Error: CSV file doesn't exist. Executed on $((get-date).DateTime)"
          #Stopping transcript
          stop-transcript;
          Exit
      } 