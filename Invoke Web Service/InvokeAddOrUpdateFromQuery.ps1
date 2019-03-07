#Adding references to SharePoint client assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

#Script variables
$WebUrl = "https://teradata.sharepoint.com/sites/compas"
$webRequestUrl = "https://teradatadamservices-prod.azurewebsites.net/integrationsvc/Asset/Add"
$listNames = @("DocumentAssets", "DocumentAssets2", "DocumentAssets3", "ExternalAssets", "MigratedURLAssets", "MigratedAssetFiles")

# Date and time of COMP-723 script start
$dateToQuery = "2018-03-28T08:30:00Z"

# Queries set for dynamic use to prevent view limit error
$queryAll = "<View><Query><Where><Geq><FieldRef Name='Modified' /><Value Type='DateTime' IncludeTimeValue='TRUE'>" + $dateToQuery + "</Value></Geq></Where></Query></View>"
$camlQueries = @(@{name="migratedNew"; qry="<View><Query><Where><And><Geq><FieldRef Name='Modified' /><Value Type='DateTime' IncludeTimeValue='TRUE'>" + $dateToQuery + "</Value></Geq><Geq><FieldRef Name='ActivationDate' /><Value Type='DateTime'>2014-12-01</Value></Geq></And></Where></Query></View>"},
                 @{name="migratedOld"; qry="<View><Query><Where><And><And><Geq><FieldRef Name='Modified' /><Value Type='DateTime' IncludeTimeValue='TRUE'>" + $dateToQuery + "</Value></Geq><Geq><FieldRef Name='ActivationDate' /><Value Type='DateTime'>2008-12-31</Value></Geq></And><Leq><FieldRef Name='ActivationDate' /><Value Type='DateTime'>2014-11-30</Value></Leq></And></Where></Query></View>"},
                 @{name="migratedOldest"; qry="<View><Query><Where><And><Geq><FieldRef Name='Modified' /><Value Type='DateTime' IncludeTimeValue='TRUE'>" + $dateToQuery + "</Value></Geq><Leq><FieldRef Name='ActivationDate' /><Value Type='DateTime'>2008-12-30</Value></Leq></And></Where></Query></View>"})

# Get login
$Username = "adam.smith@teradata.com"
$Password = Read-Host -Prompt "Enter your password" -AsSecureString
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)

#Bind to CSOM
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
$Context.Credentials = $Creds
$Site = $Context.Site
$Context.Load($Site)
$Context.ExecuteQuery()

# Invoke the web service for each item found
Function InvokeForItems($qry) {

    $items = $list.GetItems($qry)
    $Context.Load($items)
    $Context.ExecuteQuery()

    Write-Host $items.Count "items found in" $listName

    For ($i=0; $i -lt $items.Count; $i++) {
    
        $percComplete = ($i/$items.Count*100)
        $percComplete = [math]::Round($percComplete,4)
    
        Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete in $listName Completed" -PercentComplete $percComplete
    
        $assetId = $items[$i]["Asset_x0020_ID"]
        $itemId = $items[$i].Id

        $item = $list.GetItemById($itemId)
        $Context.Load($item)
        $Context.ExecuteQuery()

        $currReqUrl = $webRequestUrl + "?itemId=$itemId&listName=$listName"
        Invoke-WebRequest -Uri $currReqUrl -Method Post

        Start-Sleep -Seconds 3

        $counter++;
    }
}

$counter = 0

# Perform check in each Asset list
foreach ($listName in $listNames) {

    Write-Host "Loading items from" $listName
    
    $list = $Context.Web.Lists.GetByTitle($listName)
    $Context.Load($list)
    $Context.ExecuteQuery()

    $qry = new-object Microsoft.SharePoint.Client.CamlQuery

    if ($listName -ne "MigratedAssetFiles") {
        
        $qry.ViewXml = $queryAll

        InvokeForItems($qry)

    } else {
        For ($q=0; $q -lt $camlQueries.Count; $q++) {
            $qry.ViewXml = $camlQueries[$q]["qry"]

            InvokeForItems($qry)
        }
    }    
}

Write-Host $counter "Items Updated"

$Context.Dispose()