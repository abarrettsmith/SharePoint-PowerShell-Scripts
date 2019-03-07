#Script variables
$WebUrl = "https://teradata.sharepoint.com/sites/compas"
$listName = "DocumentAssets"

#Adding references to SharePoint client assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

$Username = "adam.smith@teradata.com"
$Password = Read-Host -Prompt "Enter your password" -AsSecureString
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)

#Bind to CSOM
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
$Context.Credentials = $Creds
$Site = $Context.Site
$Context.Load($Site)
$Context.ExecuteQuery()

$list = $Context.Web.Lists.GetByTitle($listName)
$qry = new-object Microsoft.SharePoint.Client.CamlQuery
$qry.ViewXml = "<Where><Neq><FieldRef Name='ScrubLevel'/><Value Type='String'></Value></Neq></Where>"
$items = $list.GetItems($qry)
$Context.Load($items)
$Context.ExecuteQuery()

Write-Host $items.Count

$counter = 0
$csvStr = @()

For ($i=0; $i -lt $items.Count; $i++) {
    
    $percComplete = ($j/$item.Count*100)
    $percComplete = [math]::Round($percComplete,4)
    
    Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete
    
    Write-Host $item["Asset_x0020_ID"]
    
    $props = New-Object PSObject -Property @{
        "Asset ID" = $items[$i]["Asset_x0020_ID"]
    }
    $csvStr += $props

    $counter++;
}

$csvStr | Export-Csv C:\Users\AS255108\Desktop\SuspendedWorkflows_uat.csv
Write-Host $counter "Items Found"

$Context.Dispose()