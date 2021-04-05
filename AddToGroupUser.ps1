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
        C:\PS> AddToGroupUser.ps1 -Csvfile "C:\temp\UserAccounts.csv"

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
        HelpMessage="Enter path and filename of CSV file"
        )]
        [String]$csvfile
        )

    Function AddToGroup{
       
                $GroupExists = Get-ADGroup -Filter { Name -eq $Group }              

                If($GroupExists)
                {
                    Add-ADGroupMember -Identity $Group -Members $Accountname 
                    Write-host "Added user $accountname to $group"
	            }
                Else
                {
                    write-output "Group $Group doesn't exist. User $Accountname hasn't been added."
	            }               
            
        }

    Function EmptyGroup{
       
                $GroupExists = Get-ADGroup -Filter { Name -eq $GroupToEmpty }              

                If($GroupExists)
                {
                    Get-ADGroupMember -Identity $GroupToEmpty | ForEach-Object {Remove-ADGroupMember $GroupToEmpty $_ -Confirm:$false}
                    Write-host "Removed users from $GroupToEmpty"
	            }
                Else
                {
                    write-output "Group $GroupToEmpty doesn't exist"
                    Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject "Add user to group: Error emptying group $GroupToEmpty" -SmtpServer $SmtpServer -UseSsl     
                    stop-transcript;
                    exit 
	            }               
            
        }

#Parameters declaration
$csvfirstline = Import-csv "$csvfile" -Header 'Account','Group' -Delimiter ";" | Select-Object -Skip 1 -First 1
$GroupToEmpty = $csvfirstline.Group
$LogFile = "$($env:systemdrive)\logs\AddToGroupUser_log_$GroupToEmpty.txt"
$SmtpServer = "serv09.tnm.local"
$MailFrom = "ad@tmg.dk"
$MailTo = "brku@tmg.dk"
$MailSubject = "Add users to AD group report"
#$csvfile = "\\serv06\scripts\AD\GPO\Firewall_Full.csv"


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

    Try {    
        #Imports the Csv file
        $csv = Import-Csv "$csvfile" -Header 'Account','Group' -Delimiter ";" | Select-Object -Skip 1
        #$csv = Import-csv "\\serv06\scripts\AD\GPO\Firewall_Full.csv" -Header 'Account','Group' -Delimiter ";" | Select-Object -Skip 1

        #Empty group before adding users
        $csvfirstline = Import-csv "$csvfile" -Header 'Account','Group' -Delimiter ";" | Select-Object -Skip 1 -First 1
        $GroupToEmpty = $csvfirstline.Group
        EmptyGroup

        #Loops through every user and deactivates them
        ForEach($account in $csv)
        {
            #Read user data from each field in each row and assign the data to a variable as below
            $Accountname 	= $account."Account"
            $Group 	        = $account."Group"
                         
		    AddToGroup
	                     
        }
        
        #Email is sent with information about users that have been created
        write-output "Users has been added to group $((get-date).DateTime)"
        Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject "$MailSubject" -Body "Add users to group: The following users in the attached file has been added to a group" -Attachments "$csvfile" -SmtpServer $SmtpServer -UseSsl
        Remove-Item -Path "$csvfile" -Confirm:$false -Verbose
}
#Catch if user creation fails. Logged to file and email sent.
    catch {
            write-output "Error: Executed the add users to group script on $((get-date).DateTime) with the error $_"
            Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject 'Add users to group: Errror running add users to group script' -Body "Error: Executed the add users to group script on $((get-date).DateTime) with the error $_" -SmtpServer $SmtpServer -UseSsl      
          }           





    #Stopping transcript
    stop-transcript;
    exit