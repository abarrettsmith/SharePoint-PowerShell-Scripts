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


$csv = Import-Csv c:\Users\AS255108\Desktop\SuspendedWorkflows_prod.csv
$assets = $csv | Select 'Asset ID', 'WF ID'

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

    # Start workflow
    $workflowId = $WorkflowAssociations | Where-Object {$_.Id -eq $assets[$j].'WF Id'}
    $startMsg = [string]::Format("Starting workflow, on ListItemId {0}", $assets[$j].'Asset ID')

    Write-Host $startMsg
        
    $Action = $WorkflowInstanceService.StartWorkflowOnListItem($workflowId, $assets[$j].'Asset ID'.Substring(2), $Dict)
    $ClientContext.ExecuteQuery()
        
    $props = New-Object PSObject -Property @{
        "Item ID" = $assets[$j].'Asset ID'
    }
    $csvStr += $props

    # Sleep for 30 seconds
    Start-Sleep -s 30
}
$csvStr | Export-Csv c:\Users\AS255108\Desktop\PS_Scripts\Workflows\RestartedWorkflows.csv