####################################
# Script: PublishFilesSPO.ps1      #
# Version: 2.0                     #
# Rapid Circle (c) 2016            #
# by Mark Overdijk & Massimo Prota #
####################################

# Clear the screen
Clear-Host

# Add Wave16 references to SharePoint client assemblies - required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

# Parameters
# Specify the subsite URL where the list/library resides
$SiteUrl = "https://TENANT.sharepoint.com/"
# Title of the List/Library  
$ListName = "LISTNAME"
# Username with sufficient publish/approve permissions
$UserName = "USERNAME"
# User will be prompted for password

# Set Transcript file name
$Now = Get-date -UFormat %Y%m%d_%H%M%S
$File = "PublishFilesSPO_$Now.txt"
#Start Transcript
Start-Transcript -path $File | out-null

# Display the data to the user
Write-Host "/// Values entered for use in script ///" -foregroundcolor cyan
Write-Host "Site: " -foregroundcolor white -nonewline; Write-Host $SiteUrl -foregroundcolor green
Write-Host "List name: " -foregroundcolor white -nonewline; Write-Host $ListName -foregroundcolor green
Write-Host "Useraccount: " -foregroundcolor white -nonewline; Write-Host $UserName -foregroundcolor green
# Prompt User for Password
$SecurePassword = Read-Host -Prompt "Password" -AsSecureString
Write-Host "All files in " -foregroundcolor white -nonewline; Write-Host $ListName -foregroundcolor green -nonewline; Write-Host " on site " -foregroundcolor white -nonewline; Write-Host $SiteUrl -foregroundcolor green -nonewline; Write-Host " will be published by UserName "  -foregroundcolor white -nonewline; Write-Host $UserName  -foregroundcolor green
Write-Host " "

# Prompt to confirm
Write-Host "Are these values correct? (Y/N) " -foregroundcolor yellow -nonewline; $confirmation = Read-Host

# Run script when user confirms
 if ($confirmation -eq 'y') {
 
# Bind to site collection
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
$Context.Credentials = $credentials

# Bind to list 
$list = $Context.Web.Lists.GetByTitle($ListName)
# Query for All Items 
$query = New-Object Microsoft.SharePoint.Client.CamlQuery
$query.ViewXml = " "  
$collListItem = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
$Context.Load($List)
$Context.Load($collListItem)
$Context.ExecuteQuery()

# Go through process for all items
foreach ($ListItem in $collListItem){
# Adding spacer
 Write-Host " "
 Write-Host "/////////////////////////////////////////////////////////////"
 Write-Host " "
# Write the Item ID, the FileName and the Modified date for each items which is will be published 
 Write-Host "Working on file: " -foregroundcolor yellow -nonewline; Write-Host $ListItem.Id, $ListItem["FileLeafRef"], $ListItem["Modified"]
 
# Un-comment below "if" when you want to add a filter which files will be published 
# Fill out the details which files should be skipped. Example will skip all files which where modifed last < 31-jan-2015
#
# if (
# $ListItem["Modified"] -lt "01/31/2015 00:00:00 AM"){
# Write-Host "This item was last modified before January 31st 2015" -foregroundcolor red
# Write-Host "Skip file" -foregroundcolor red
# continue
# }

# Check if file is checked out by checking if the "CheckedOut By" column does not equal empty
if ($ListItem["CheckoutUser"] -ne $null){
# Item is not checked out, Check in process is applied
	Write-Host "File: " $ListItem["FileLeafRef"] "is checked out." -ForegroundColor Cyan
	$listItem.File.CheckIn("Auto check-in by PowerShell script", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
	Write-Host "- File Checked in" -ForegroundColor Green
}
# Publishing the file
Write-Host "Publishing file:" $ListItem["FileLeafRef"] -ForegroundColor Cyan
$listItem.File.Publish("Auto publish by PowerShell script")
Write-Host "- File Published" -ForegroundColor Green

# Check if the file is approved by checking if the "Approval status" column does not equal "0" (= Approved)
if ($List.EnableModeration -eq $true){
# if Content Approval is enabled, the file will be approved
if ($ListItem["_ModerationStatus"] -ne '0'){
# File is not approved, approval process is applied
	Write-Host "File:" $ListItem["FileLeafRef"] "needs approval" -ForegroundColor Cyan
	$listItem.File.Approve("Auto approval by PowerShell script")
	Write-Host "- File Approved" -ForegroundColor Green
}
else {
Write-Host "- File has already been Approved" -ForegroundColor Green
}
}
$Context.Load($listItem) 
$Context.ExecuteQuery()
}
# Adding footer
 Write-Host " "
 Write-Host "/////////////////////////////////////////////////////////////"
 Write-Host " "
 Write-Host "Script is done" -ForegroundColor Green
 Write-Host "Files have been published/approved" -ForegroundColor Green
 Write-Host "Thank you for using PublishFilesSPO.ps1 by Rapid Circle" -foregroundcolor cyan
 Write-Host " "
 }
# Stop script when user doesn't confirm
else {
 Write-Host " "
 Write-Host "Script cancelled by user" -foregroundcolor red
 Write-Host " "
 }
 Stop-Transcript | out-null
##############################
# Rapid Circle               #
# http://www.rapidcircle.com #
##############################