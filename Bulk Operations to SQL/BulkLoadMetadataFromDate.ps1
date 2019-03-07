# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll"
 
# Specify tenant admin and site URL
$SiteUrl = "https://teradata.sharepoint.com/sites/COMPAS/"
$ListName = "DocumentAssets"
$UserName = "adam.smith@teradata.com"
$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString

# Bind to site collection
$ClientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
$ClientContext.Credentials = $credentials
$ClientContext.ExecuteQuery()
 
# Get List
$List = $ClientContext.Web.Lists.GetByTitle($ListName)
 
$ClientContext.Load($List)
$ClientContext.ExecuteQuery()

$qry = new-object Microsoft.SharePoint.Client.CamlQuery
$qry.ViewXml = "<View><Query><Where><Geq><FieldRef Name='Modified' /><Value Type='DateTime'>2017-07-19-08:00</Value></Geq></Where></Query></View>"

$ListItems = $List.GetItems($qry)
$ClientContext.Load($ListItems)
$ClientContext.ExecuteQuery()
 
    $csvStr = @()
    $webRequestUrl = "https://teradatadamservices-prod.azurewebsites.net/integrationsvc/Asset/Add"

foreach($item in $ListItems) {
    $itemId = $item.Id

    Write-Progress -Activity "Updating" -Status $itemId

    $currReqUrl = $webRequestUrl + "?itemId=$itemId&listName=$ListName"
    Invoke-WebRequest -Uri $currReqUrl -Method Post
}
    $csvStr | Export-Csv C:\Users\AS255108\Desktop\updated_workflows.csv
