#Script variables
$WebUrl = "https://spd1806.sharepoint.com/sites/development/"

#Adding references to SharePoint client assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#$Username = Read-Host -Prompt "Enter your username"
$Username = "adam.smith@spd1806.onmicrosoft.com"
$Password = Read-Host -Prompt "Enter your password" -AsSecureString
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)

#Bind to CSOM
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
$Context.Credentials = $Creds
$Site = $Context.Site
$Context.Load($Site)
$Context.ExecuteQuery()

Write-Host $Site.Id

Read-Host