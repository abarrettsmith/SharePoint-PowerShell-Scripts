$itemId = 72
$listName = "ExternalAssets"

$webRequestUrl = "https://teradatadamservices-prod.azurewebsites.net/integrationsvc/Asset/Add"
    
    Write-Progress -Activity "Updating" -Status $i

    $currReqUrl = $webRequestUrl + "?itemId=$i&listName=$listName"
    Invoke-WebRequest -Uri $currReqUrl -Method Post
    