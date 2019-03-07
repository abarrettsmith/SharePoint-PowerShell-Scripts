# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll")
 
# Specify tenant admin and site URL
$SiteUrl = "https://teradata.sharepoint.com/sites/COMPAS/"
$ListName = "Content Ops Work Queue"
$UserName = "adam.smith@teradata.com"
$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
 
# Connect to site
$ClientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
$ClientContext.Credentials = $credentials
$ClientContext.ExecuteQuery()
 
# Get List
$List = $ClientContext.Web.Lists.GetByTitle($ListName)
$ClientContext.Load($List)
$ClientContext.ExecuteQuery()

# Query entire list
$qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()

#$qry = new-object Microsoft.SharePoint.Client.CamlQuery
#$qry.ViewXml = "<View><Query><Where><Geq><FieldRef Name='Created' /><Value Type='DateTime'>2017-07-01</Value></Geq></Where></Query></View>"

# Load list items
$ListItems = $List.GetItems($qry)
$ClientContext.Load($ListItems)
$ClientContext.ExecuteQuery()

$csvStr = @()
For ($i=0; $i -lt $ListItems.Count; $i++) {

    $percComplete = ($i/$ListItems.Count*100)
    $percComplete = [math]::Round($percComplete,4)
    Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete

    $assetId = $ListItems[$i]["Asset_x0020_ID"]

    # Get List
    $ArchiveList = $ClientContext.Web.Lists.GetByTitle("Content Operations Archive")
    $ClientContext.Load($ArchiveList)
    $ClientContext.ExecuteQuery()

    $Archiveqry = new-object Microsoft.SharePoint.Client.CamlQuery
    $Archiveqry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Asset_x0020_ID' /><Value Type='Text'>$assetId</Value></Eq></Where></Query></View>"

    Write-Host "Searching for asset " $assetId

    # Load list items
    $ArchiveListItems = $ArchiveList.GetItems($Archiveqry)
    $ClientContext.Load($ArchiveListItems)
    $ClientContext.ExecuteQuery()

    For ($j=0; $j -lt $ArchiveListItems.Count; $j++) {
        
        Write-Host "Asset ID matched archived item:" $ArchiveListItems[$i]["Asset_x0020_ID"]

        $props = New-Object PSObject -Property @{
            "Asset ID" = $ArchiveListItems[$j]["Asset_x0020_ID"]
        }
        $csvStr += $props
    }
    
}
$csvStr | Export-Csv c:\Users\AS255108\Desktop\COWQArchived.csv