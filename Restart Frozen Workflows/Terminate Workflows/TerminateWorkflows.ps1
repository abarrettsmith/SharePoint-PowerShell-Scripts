# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll"
 
# Specify tenant admin and site URL
$SiteUrl = "https://teradata.sharepoint.com/sites/COMPAS/"
$ListName = "ExternalAssets"
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

    $counter = 0;
    $csvStr = @()
    # Loop Terminate Workflow
    For ($j=0; $j -lt $ListItems.Count; $j++){
 
       If ($ListItems[$j]["Reuse_x0020_Credit_x0020_Panel"] -notcontains '<AssetID type="System.String"></AssetID>'){
           $itemWfInstances = $WorkflowInstanceService.EnumerateInstancesForListItem($List.Id, $ListItems[$j].Id)
           $ClientContext.Load($itemWfInstances)
           $ClientContext.ExecuteQuery()
           for ($k=0;$k -lt $itemWfInstances.Count;$k++)
           {
                try {
                    $counter++;
                    $WorkflowInstanceService.TerminateWorkflow($itemWfInstances[$k])
                    Write-Host "Worfklow terminated on" $ListItems[$j].Id
                    
                    $ListItems[$j].Update()
                    $ClientContext.ExecuteQuery()

                    $props = New-Object PSObject -Property @{
                        "Asset ID" = $ListItems[$j]["Asset_x0020_ID"]
                    }
                    $csvStr += $props
                } catch {
                    Write-Host "Error terminating workflow on" $ListItems[$j].Id "Details: $_"
                }
           }
       }
    }
    $csvStr | Export-Csv C:\Users\AS255108\Desktop\WorkflowRestart\WorkflowsTermianted_UAT.csv
    Write-Host $counter

Write-Host "Ready to Loop"
