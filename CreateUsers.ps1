<#
        .SYNOPSIS
        Creates user accounts from a CSV file

        .DESCRIPTION
        Script is used for automating the creation user accounts in Active Directory. 

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
        C:\PS> CreateUserAccounts.ps1 -Csvfile "C:\temp\UserAccounts.csv"

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
 
    New-ADUser `
            -SamAccountName $Username `
            -UserPrincipalName "$upn" `
            -Name "$DisplayName" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Enabled $True `
            -DisplayName "$DisplayName" `
            -Path $OU `
            -City $city `
            -PostalCode $Zip `
            -HomePhone $HomePhone `
            -Company $company `
            -StreetAddress $street `
            -OfficePhone $telephone `
            -EmailAddress $email `
            -Title $jobtitle `
            -MobilePhone $Mobile `
            -Fax $Fax `
            -Department $department `
            -AccountPassword (convertto-securestring $Password -AsPlainText -Force) -ChangePasswordAtLogon $False `
            -PasswordNeverExpires $True `
            -CannotChangePassword $True `
            -OtherAttributes @{
                                pager = $Pager
                              }
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
                    Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject "User creation: Error adding user $Username to group $Group" -SmtpServer $SmtpServer -UseSsl      
	            }               
            }
}


# Import active directory module for running AD cmdlets
Import-Module activedirectory

#Parameters declaration
$LogFile = "$($env:systemdrive)\logs\CreateUserAccounts_log.txt"
$SmtpServer = "serv09.tnm.local"
$MailFrom = "ad@tmg.dk"
$MailTo = "brku@tmg.dk"
$MailSubject = "Create User accounts report"
$NewFileNameAfterJob = "\\serv06\scripts\AD\Users\UserOpret.old"
$OUTopLevelPath = "OU=Afdelinger,OU=Domain Users,DC=tnm,DC=local"

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
        $csv = Import-Csv "$csvfile" -Header 'ADOU','First name','Last name','Initialer','Display name','Telephone','Email','Street','City','Zip','User','Password','Home','Pager','Mobile','Fax','Job Title','Department','Company','Member of' -Delimiter ";" | Select-Object -Skip 1
        #$csv = Import-csv "\\serv06\scripts\AD\Users\UserOpret.csv" -Header 'ADOU','First name','Last name','Initialer','Display name','Telephone','Email','Street','City','Zip','User','Password','Home','Pager','Mobile','Fax','Job Title','Department','Company','Member of' -Delimiter ";" | Select-Object -Skip 1

        #Loops through every user and deactivates them
        ForEach($User in $csv)
        {
            #Read user data from each field in each row and assign the data to a variable as below
            $OU 		= "OU=" + $User.ADOU + ",$OUTopLevelPath"
            $Firstname 	= $User."First name"
	        $Lastname 	= $User."Last Name"
            $Username 	= $User.Initialer
            $UPN  = $User.Initialer + "@tmg.dk"
	        $DisplayName 	= $User."Display name"    
            $Telephone 	= $User.Telephone
            $Email 	= $User.Email
            $Street 	= $User.Street
            $City 	= $User.City
            $Zip 	= $User.Zip
            #$User 	= $User.User
            $Password 	= $User.Password
            $HomePhone 	= $User.Home
            $Pager 	= $User.Pager
            $Mobile = $User.Mobile
            $Fax = $User.Fax
            $JobTitle = $User."Job Title"
            $Department = $User.Department
            $Company = $User.Company
            $Groups = $User."Member of" -split ","
            
            If(!(Get-ADUser -F {SamAccountName -eq $Username}))
            {
                CreateUser       
		        AddToGroups
	        }             
        }
        
        #Email is sent with information about users that have been created
        write-output "Users has been created on $((get-date).DateTime)"
        Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject "$MailSubject" -Body "User creation: The following users in the attached file has been created" -Attachments "$csvfile" -SmtpServer $SmtpServer -UseSsl
        Remove-Item -Path "$NewFileNameAfterJob" -Confirm:$false -Verbose
        Rename-Item -Path $csvfile -NewName "$NewFileNameAfterJob"
}
#Catch if user creation fails. Logged to file and email sent.
catch {
            write-output "Error: Executed the create user accounts script on $((get-date).DateTime) with the error $_"
            Send-MailMessage -From "$MailFrom" -To "$MailTo" -Subject 'User creation: Errror running creation of user account script' -Body "Error: Executed the creation user accounts script on $((get-date).DateTime) with the error $_" -SmtpServer $SmtpServer -UseSsl      
    }           





    #Stopping transcript
    stop-transcript;
    exit
    