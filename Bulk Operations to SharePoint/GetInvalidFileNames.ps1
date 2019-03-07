# Script variables
$WebUrl = "https://teradata.sharepoint.com/sites/compasdev"

# Adding references to SharePoint client assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#$Username = Read-Host -Prompt "Enter your username"
$Username = "adam.smith@teradata.com"
$Password = Read-Host -Prompt "Enter your password" -AsSecureString
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)

# Bind to CSOM
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
$Context.Credentials = $Creds
$Site = $Context.Site
$Context.Load($Site)
$Context.ExecuteQuery()

# Load list
$listName = "DocumentAssets3"
$list = $Context.Web.Lists.GetByTitle($listName)

$qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()

$items = $list.GetItems($qry)
$Context.Load($items)
$Context.ExecuteQuery()

Write-Host $items.Count "items found."

$csvStr = @()
for($i=0; $i -lt $items.Count; $i++) {

    $percComplete = ($i/$items.Count*100)
    $percComplete = [math]::Round($percComplete,4)
    Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete

    # Set variables
    $itemId = $items[$i].Id
    $assetId = $items[$i]["Asset_x0020_ID"]
    $fileName = $items[$i]["FileLeafRef"]
    
    if ($fileName -like "*"+$assetId+"*") {
        
        Write-Host "Item" $itemId ":" $assetId $fileName "- valid"

    } else {
        
        Write-Host "Item" $itemId ":" $assetId $fileName "- invalid"
        
        $props = New-Object PSObject -Property @{
            "Item ID" = $itemId
        }
        $csvStr += $props
    }
}
$csvStr | Export-Csv C:\Users\AS255108\Desktop\Invalid_FileName.csv

$Context.Dispose()