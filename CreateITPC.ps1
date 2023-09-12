<#
        .SYNOPSIS
        Creates IT PC accounts from a CSV file

        .DESCRIPTION
        Script is used for automating the creation of IT PC accounts in Active Directory.

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
        C:\PS> CreateITPCAccounts.ps1 -Csvfile "C:\temp\ITPCAccounts.csv"

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
        HelpMessage="Enter path and filename of CSV file")]
        [String]$csvfile
        )



Function CreateUser{
 Write-output "Create user 1"
    New-ADUser `
            -SamAccountName $Username `
            -UserPrincipalName "$upn" `
            -Name "$DisplayName" `
            -Enabled $True `
            -DisplayName "$DisplayName" `
            -Path $OU `
            -AccountPassword (convertto-securestring $Password -AsPlainText -Force) -ChangePasswordAtLogon $False `
            -PasswordNeverExpires $True `
            -CannotChangePassword $True
Write-output "Create user 2"
 }


Function AddToGroups{

            foreach ($Group in $Groups)
            {
                $GroupExists = Get-ADGroup -Filter { Name -eq $Group }              
                
                If($GroupExists)
                {
                    Add-ADGroupMember -Identity $Group -Members $Username
	            }
                Else
                {
                    write-output "Group $Group doesn't exist"
                    # Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject "User creation: Error adding user $Username to group $Group" -SmtpServer $SmtpServer -UseSsl      
	            }               
            }
}


# Import active directory module for running AD cmdlets
Import-Module activedirectory

#Parameters declaration
$LogFile = "$($env:systemdrive)\logs\CreateITPCAccounts_log.txt"
$SmtpServer = "serv09.tnm.local"
$MailFrom = "ad@tmg.dk"
$MailTo = "brku@tmg.dk"
$MailSubject = "Create IT PC accounts report"
$NewFileNameAfterJob = "\\serv06\scripts\AD\Users\UserITPC.old"
$OUTopLevelPath = "OU=802.1x,OU=Domain Users,DC=tnm,DC=local"


#Starting transcript
start-transcript -Path $LogFile

If(!(Test-Path -Path $csvfile))
        {
            write-output "Error: CSV file doesn't exist. Executed on $((get-date).DateTime)"
            write-output "$csvfile"
            #Stopping transcript
            stop-transcript;
            Exit
        }
    
try {    
        #Imports the Csv file
        $csv = Import-Csv "$csvfile" -Header 'Mac','Ousti','Navn' -Delimiter ";" | Select-Object -Skip 1
        #$csv = Import-csv "\\serv06\scripts\AD\Users\UserOpret.csv" -Header 'ADOU','First name','Last name','Initialer','Display name','Telephone','Email','Street','City','Zip','User','Password','Home','Pager','Mobile','Fax','Job Title','Department','Company','Member of' -Delimiter ";" | Select-Object -Skip 1

        #Loops through every user and deactivates them
        ForEach($itpc in $csv)
        {
            #Read user data from each field in each row and assign the data to a variable as below
            $OU 		= "OU=" + $itpc.Ousti + ",$OUTopLevelPath"
            $Firstname 	= $itpc.Mac
	        $Lastname 	= $itpc.Mac
            $Username 	= $itpc.Mac
            $UPN  = $itpc.Mac + "@tmg.dk"
	        $DisplayName 	= $itpc."Navn"  
            $CN         = "CN=" + "$Displayname" + "," + $OU
            $Password 	= $itpc.Mac
            
            If(!(Get-ADUser -F {SamAccountName -eq $Username}) -and (!(Get-ADUser -F * -Searchbase $OU | where name -eq $DisplayName))) 
            {         
                CreateUser 
                Write-output "Created ITPC Account $Firstname"           
	        }  

            If((Get-ADUser -F {SamAccountName -eq $Username}) -and (!(Get-ADUser -F * -Searchbase $OU | where name -eq $DisplayName))) 
            {         
                Set-ADUser -Identity $Username -DisplayName "$DisplayName"
                #Rename-ADObject -Identity $CN -NewName "$DisplayName"
                $OLDCNDATA = Get-ADUser -Identity $Username -Properties DistinguishedName  | Select DistinguishedName 
                $DistinguishedName = $OLDCNDATA.DistinguishedName
                Rename-ADObject $DistinguishedName -NewName $DisplayName
                Write-output "Updated display name on $Firstname with new display name $DisplayName"
                
                    
	        } 
                       
        }
        
        #Email is sent with information about users that have been created
        write-output "IT PC script finished processing IT PC provisioning on $((get-date).DateTime)"
        #Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject "$MailSubject" -Body "User creation: The following users in the attached file has been created" -Attachments "$csvfile" -SmtpServer $SmtpServer -UseSsl
        #Remove-Item -Path "$NewFileNameAfterJob" -Confirm:$false -Verbose
        #Rename-Item -Path $csvfile -NewName "$NewFileNameAfterJob"
}
#Catch if user creation fails. Logged to file and email sent.
catch {
            write-output "Error: Executed the create IT PC accounts script on $((get-date).DateTime) with the error $_"
            #Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject 'IT PC account creation: Errror running creation of IT PC account script' -Body "Error: Executed the IT PC account script on $((get-date).DateTime) with the error $_" -SmtpServer $SmtpServer -UseSsl      
    }           





    #Stopping transcript
    stop-transcript;
    exit