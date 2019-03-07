#Script variables
$WebUrl = "https://teradata.sharepoint.com/sites/compas"
$csv = Import-Csv C:\Users\AS255108\Desktop\fixThisShit.csv

#Set web service url
$webRequestUrl = "https://teradatadamservices-prod.azurewebsites.net/integrationsvc/Asset/Add"

#Adding references to SharePoint client assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll"

$Username = "adam.smith@teradata.com"
$Password = Read-Host -Prompt "Enter your password" -AsSecureString
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)

#Bind to CSOM
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
$Context.Credentials = $Creds
$Site = $Context.Site
$Context.Load($Site)
$Context.ExecuteQuery()

$assets = $csv | Select 'AssetID'

For ($j=0; $j -lt $assets.Count; $j++){

        #Load list
        $assetId = $assets[$j].'AssetId'

        $List = $ClientContext.Web.Lists.GetByTitle($ListName)
        $ClientContext.Load($List)
        $ClientContext.ExecuteQuery()

        # Query "content Ops Work Queue"
        $qry = new-object Microsoft.SharePoint.Client.CamlQuery
        $qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Asset_x0020_ID' /><Value Type='Text'>" + $assetId + "</Value></Eq></Where></Query></View>"

        $list = $Context.Web.Lists.GetByTitle("DocumentAssets3")

        $item = $list.GetItemById($itemId)
        $Context.Load($item)
        $Context.ExecuteQuery()

            Write-Progress -Activity "Updating" -Status $assetId

            $currReqUrl = $webRequestUrl + "?itemId=$itemId&listName=$listName"
            Invoke-WebRequest -Uri $currReqUrl -Method Post
        } else {
            $csvStr | Export-Csv C:\Users\AS255108\Desktop\Error.csv

            Read-Host

            $Context.Dispose()
        }

    $csvStr | Export-Csv C:\Users\AS255108\Desktop\fixedShit.csv