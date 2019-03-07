# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll")
 
# Specify tenant admin and site URL
$SiteUrl = "https://teradata.sharepoint.com/sites/COMPAS/"
$ListName = "Content Ops Work Queue"
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
 
# Retrieve WorkflowService related objects
$WorkflowServicesManager = New-Object Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager($ClientContext, $ClientContext.Web)
$WorkflowSubscriptionService = $WorkflowServicesManager.GetWorkflowSubscriptionService()
$WorkflowInstanceService = $WorkflowServicesManager.GetWorkflowInstanceService()
$ClientContext.Load($WorkflowServicesManager)
$ClientContext.Load($WorkflowSubscriptionService)
$ClientContext.Load($WorkflowInstanceService)
$ClientContext.ExecuteQuery()

# Get WorkflowAssociations with List
$WorkflowAssociations = $WorkflowSubscriptionService.EnumerateSubscriptionsByList($List.Id)
$ClientContext.Load($WorkflowAssociations)
$ClientContext.ExecuteQuery()
 
# Prepare Start Workflow Payload
$Dict = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'

$csv = Import-Csv c:\Users\AS255108\Desktop\SubmissionsToCheck.csv
$assets = $csv | Select 'Asset ID'

$csvStr = @()
# Loop List Items to Start Workflow
For ($j=0; $j -lt $assets.Count; $j++){

    $percComplete = ($j/$assets.Count*100)
    $percComplete = [math]::Round($percComplete,4)
    Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete

    $itemId = $assets[$j].'Asset ID'.Substring(2)

    $ListItem = $List.GetItemById($itemId)
    $ClientContext.Load($ListItem)
    $ClientContext.ExecuteQuery()

    $itemWfInstances = $WorkflowInstanceService.EnumerateInstancesForListItem($List.Id, $ListItem.Id)
    $ClientContext.Load($itemWfInstances)
    $ClientContext.ExecuteQuery()
    for ($k=0;$k -lt $itemWfInstances.Count;$k++)
        {
            if ($itemWfInstances[$k].Status -eq "Suspended" -and $itemWfInstances[$k].WorkflowSubscriptionId -eq "8ca76f14-7408-4793-bf9b-d0e1c2061f02") {
                   
                    try {
                        $counter++;
                        Write-Host "Subcription ID" $itemWfInstances[$k].WorkflowSubscriptionId
                        
                        $props = New-Object PSObject -Property @{
                            "Asset ID" = $ListItem["Asset_x0020_ID"]
                            "WF ID" = $itemWfInstances[$k].WorkflowSubscriptionId
                        }
                        $csvStr += $props
                    } catch {
                        Write-Host "Error in workflow on" $ListItem.Id "Details: $_"
                    }
                }
           }
       }
    $csvStr | Export-Csv C:\Users\AS255108\Desktop\SuspendedOpen.csv
    Write-Host $counter "Items Found"