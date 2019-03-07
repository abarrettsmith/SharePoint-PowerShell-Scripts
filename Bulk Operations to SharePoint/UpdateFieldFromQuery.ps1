# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll")
 
# Specify tenant admin and site URL
$SiteUrl = "https://teradata.sharepoint.com/sites/COMPAS/"
$ListName = "DocumentAssets3"
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

# Query "content Ops Work Queue"
$qry = new-object Microsoft.SharePoint.Client.CamlQuery
$qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='FormFlow' /><Value Type='Text'>No Form Application</Value></Eq></Where></Query></View>"

# Query entire list
#$qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()

$ListItems = $List.GetItems($qry)
$ClientContext.Load($ListItems)
$ClientContext.ExecuteQuery()

# Loop List Items to Start Workflow
For ($j=0; $j -lt $ListItems.Count; $j++){

    # Calculate Completion
    $percComplete = ($j/$ListItems.Count*100)
    $percComplete = [math]::Round($percComplete,4)
    Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete

    # Load Item
    $itemId = $ListItems[$j].Id
    $ListItem = $List.GetItemById($itemId)
    $ClientContext.Load($ListItem)
    $ClientContext.ExecuteQuery()

    #$AppArray = @("RTIM","CIM","DCM")

    # Update with correct Form ID
    #$ListItem["Applications"] = $AppArray #"RTIM", "CIM", "DCM"

    $ListItem["FormFlow"] = "No"
    $ListItem.Update()
    $ClientContext.ExecuteQuery()

    Write-Host $itemId "updated"
}