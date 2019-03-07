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

$csv = Import-Csv c:\Users\AS255108\Desktop\ContentOpsArchive.csv
$assets = $csv | Select 'Asset ID'

$csvStr = @()
# Loop List Items to Start Workflow
For ($j=0; $j -lt $assets.Count; $j++){

    $percComplete = ($j/$assets.Count*100)
    $percComplete = [math]::Round($percComplete,4)
    Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete

    try {
        $itemId = $assets[$j].'Asset ID'.Substring(2)
    
        $ListItem = $List.GetItemById($itemId)
        $ClientContext.Load($ListItem)
        $ClientContext.ExecuteQuery()

        $ListItem.DeleteObject()
        $ClientContext.ExecuteQuery()
        
        $props = New-Object PSObject -Property @{
            "Asset ID" = $assets[$j].'Asset ID'
        }
    }
    catch {
        
    }
    $csvStr += $props
}
$csvStr | Export-Csv c:\Users\AS255108\Desktop\SubmissionsArchived.csv