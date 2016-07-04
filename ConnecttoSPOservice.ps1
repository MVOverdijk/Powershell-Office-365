####################################
# Script: ConnecttoSPOservice.ps1  #
# Version: 1.0                     #
# Rapid Circle (c) 2015            #
# by Mark Overdijk                 #
####################################
set-executionpolicy remotesigned
write-host "Only fill out the Tenant name, as in only the Uppercase name in: https://TENANT.sharepoint.com"
$Tenant = Read-Host -Prompt "Tenant"
$TenantURL = "https://$Tenant-admin.sharepoint.com"
write-host $tenantURL

$error.clear()
try { Connect-SPOService -Url $TenantURL -credential (Get-Credential)}
catch { 
Write-host $error -foregroundcolor red }
if (!$error) {
Write-host "Login succesful. You are now connected to SPO $TenantURL" -foregroundcolor green
}