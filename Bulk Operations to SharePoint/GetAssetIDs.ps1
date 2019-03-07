#Script variables
$WebUrl = "https://teradata.sharepoint.com/sites/compas"

#Adding references to SharePoint client assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
$Username = Read-Host -Prompt "Enter your username"
$Password = Read-Host -Prompt "Enter your password" -AsSecureString
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)

#Bind to CSOM
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
$Context.Credentials = $Creds
$Site = $Context.Site
$Context.Load($Site)
$Context.ExecuteQuery()


$listName = "DocumentAssets"
$list = $Context.Web.Lists.GetByTitle($listName)
#$qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
$qry = [Microsoft.SharePoint.Client.CamlQuery]
$qry.Query = "<Where><Lt><FieldRef Name='AssetStatus'/><Value Type='String' IncludeTimeValue='False'></Value></Lt></Where>"
$items = $list.GetItems($qry)
$Context.Load($items)
$Context.ExecuteQuery()


foreach($item in $items) {
    Write-Host $item["Asset_x0020_ID"]
}

Read-Host

$Context.Dispose()