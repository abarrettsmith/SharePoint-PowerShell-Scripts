# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll"
 
# Specify tenant admin and site URL
$SiteUrl = "https://teradata.sharepoint.com/sites/COMPASUAT/"
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
 
$ListItems = $List.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
$ClientContext.Load($ListItems)
$ClientContext.ExecuteQuery()
 
# Create WorkflowServicesManager instance
$WorkflowServicesManager = New-Object Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager($ClientContext, $ClientContext.Web)
 
# Connect to WorkflowSubscriptionService
$WorkflowSubscriptionService = $WorkflowServicesManager.GetWorkflowSubscriptionService()
 
# Connect WorkflowInstanceService instance
$WorkflowInstanceService = $WorkflowServicesManager.GetWorkflowInstanceService()
 
$ClientContext.Load($WorkflowServicesManager)
$ClientContext.Load($WorkflowSubscriptionService)
$ClientContext.Load($WorkflowInstanceService)
$ClientContext.ExecuteQuery()
 
# Get WorkflowAssociations with List
$WorkflowAssociations = $WorkflowSubscriptionService.EnumerateSubscriptionsByList($List.Id)
$ClientContext.Load($WorkflowAssociations)
$ClientContext.ExecuteQuery()
 
# Prepare Terminate Workflow Payload
$EmptyObject = New-Object System.Object
$Dict = New-Object 'System.Collections.Generic.Dictionary[System.String,System.Object]'

# Loop Terminate Workflow
function csvLoop{
    $csvStr = @()
    For ($j=0; $j -lt $ListItems.Count; $j++){
    $percComplete = ($j/$ListItems.Count*100)
    $percComplete = [math]::Round($percComplete,4)
    Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete
        If (!$ListItems[$j]["Form_x0020_Submission_x0020_Proc"].Description.Contains("Complete")){
            $current = $ListItems[$j]                      
            $props = New-Object PSObject -Property @{
                "Asset ID" = $current["Asset_x0020_ID"]
                "Modified" = $current["Modified"]
                "Workflow Status" = $current["Form_x0020_Submission_x0020_Proc"].Description
            }
            $csvStr += $props
        }
    }
    $csvStr | Export-Csv C:\Users\AS255108\Desktop\PS_Scripts\WorkflowRestart\NotCompletedWorkflows.csv
}
Write-Host "Ready to Loop"