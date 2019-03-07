# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll"
 
# Specify tenant admin and site URL
$SiteUrl = "https://teradata.sharepoint.com/sites/COMPASDEV/"
$ListName = "DocumentAssets2"
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
$qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Asset_x0020_ID' /><Value Type='Text'>DA000070</Value></Eq></Where></Query></View>"

# Get Items
$ListItems = $List.GetItems($qry)
$ClientContext.Load($ListItems)
$ClientContext.ExecuteQuery()
 
# Count Items
$counter = 0;

    # Loop Through Items
    For ($j=0; $j -lt $ListItems.Count; $j++){
        
        $counter += 1
        $percComplete = ($j/$ListItems.Count*100)
        $percComplete = [math]::Round($percComplete,4)
        Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete

        Write-Host "Before Item"
        Write-Host $ListItems[$j]["FileLeafRef"]
        
        $ListItems[$j]["FileLeafRef"] = $ListItems[$j]["FileLeafRef"] + " - " + $ListItems[$j]["Asset_x0020_ID"]

        Write-Host "After Item"
        Write-Host $ListItems[$j]["FileLeafRef"]

        #$ListItems[$j].Update()
        #$ClientContext.ExecuteQuery()
    }

Write-Host "Item Counter:" $counter
