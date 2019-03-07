# Script variables
$WebUrl = "https://teradata.sharepoint.com/sites/compas"

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
$listName = "Content Ops Work Queue"
$list = $Context.Web.Lists.GetByTitle($listName)
$qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
$items = $list.GetItems($qry)
$Context.Load($items)
$Context.ExecuteQuery()

Write-Host $items.Count


$getCompositeIdUrl = "https://teradatadamservices-uat.azurewebsites.net/integrationsvc/Composite/GetCompositeId"
$addPendingCompositeUrl = "https://teradatadamservices-uat.azurewebsites.net/integrationsvc/Composite/AddPendingComposite"

$csvStr = @()
for($i=0; $i -lt $items.Count; $i++) {

    $percComplete = ($i/$items.Count*100)
    $percComplete = [math]::Round($percComplete,4)
    Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete

    # Set variables
    $itemId = $items[$i].Id
    $assetId = $items[$i]["Asset_x0020_ID"]
    
    # Get Composite ID
    $currReqUrl = $getCompositeIdUrl + "?assetId=$assetId"
    $dbItemId = Invoke-WebRequest -Uri $currReqUrl -Method Get

    # If error free add pending composite asset, otherwise write to CSV
    if ($dbItemId -notcontains "error") {
        
        $currReqUrl = $addPendingCompositeUrl + "?itemId=$itemId&listName=$listName&DBItemId=$dbItemId"
        Invoke-WebRequest -Uri $currReqUrl -Method Post

    } else {

        $props = New-Object PSObject -Property @{
            "Item ID" = $itemId
        }

    }
    
    #Start-Sleep -s 30
}
$csvStr | Export-Csv C:\Users\AS255108\Desktop\MissingSQLRow.csv


Write-Host "Script complete."
Read-Host

$Context.Dispose()