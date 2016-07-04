###############################################################################################################
# Script: AddGeolocationOffice365.ps1                                                                         #
# Version: 1.0                                                                                                #
# Rapid Circle (c) 2015                                                                                       #
# by Mark Overdijk                                                                                            #
# How to use guide: http://www.rapidcircle.com/walkthrough-add-geolocation-column-to-your-list-in-office-365/ #
###############################################################################################################
set-executionpolicy Unrestricted
<# Clear the screen #>
Clear-Host
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
<# Get User input #>
$SiteURL = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Site URL, example: https://yourtenant.sharepoint.com/sites/yoursite", "URL", "")
$Login = [Microsoft.VisualBasic.Interaction]::InputBox("Office 365 Username, example: youradmin@yourtenant.onmicrosoft.com", "Username", "")
$ListName = [Microsoft.VisualBasic.Interaction]::InputBox("List name to add Geolocation column", "ListName", "")
$ColumnName = [Microsoft.VisualBasic.Interaction]::InputBox("Column name for the Geolocation column", "ColumnName", "")
$BingMapsKey = [Microsoft.VisualBasic.Interaction]::InputBox("Bing Maps key", "Key", "")
<# Show results #>
Write-Host "/// Values entered for use in script ///" -foregroundcolor magenta
Write-Host "Site: " -foregroundcolor white -nonewline; Write-Host $SiteURL -foregroundcolor green
Write-Host "Useraccount: " -foregroundcolor white -nonewline; Write-Host $Login -foregroundcolor green
Write-Host "List name: " -foregroundcolor white -nonewline; Write-Host $ListName -foregroundcolor green
Write-Host "Geolocation column name: " -foregroundcolor white -nonewline; Write-Host $ColumnName -foregroundcolor green
Write-Host "Bing Maps key: " -foregroundcolor white -nonewline; Write-Host $BingMapsKey -foregroundcolor green
Write-Host " "
<# Confirm before proceed #>
Write-Host "Are these values correct? (Y/N) " -foregroundcolor yellow -nonewline; $confirmation = Read-Host
if ($confirmation -eq 'y') {
$WebUrl = $SiteURL
$EmailAddress = $Login
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
$Credentials = Get-Credential -UserName $EmailAddress -Message "Please enter your Office 365 Password"
$Context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($EmailAddress,$Credentials.Password)
$List = $Context.Web.Lists.GetByTitle("$ListName")
$FieldXml = "<Field Type='Geolocation' DisplayName='$ColumnName'/>"
$Option=[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView
$List.Fields.AddFieldAsXml($fieldxml,$true,$option)
$Context.Load($list)
$Context.ExecuteQuery()
$web = $Context.Web
$web.AllProperties["BING_MAPS_KEY"] = $BingMapsKey
$web.Update()
$Context.ExecuteQuery()
$Context.Dispose()
Write-Host " "
Write-Host "Done!" -foregroundcolor green
Write-Host " "
}
else {
Write-Host " "
Write-Host "Script cancelled" -foregroundcolor red
Write-Host " "
}
