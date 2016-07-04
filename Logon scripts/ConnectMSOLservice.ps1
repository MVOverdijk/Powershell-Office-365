####################################
# Script: ConnectMSOLservice.ps1   #
# Version: 1.0                     #
# Rapid Circle (c) 2015            #
# by Mark Overdijk                 #
####################################
set-executionpolicy remotesigned
Import-Module MSOnline
$O365Cred = Get-Credential
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $O365Session
Connect-MsolService –Credential $O365Cred