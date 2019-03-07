# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll"
 
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

# Query "Migrated Asset Files"
#$qry = new-object Microsoft.SharePoint.Client.CamlQuery
#$qry.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ActivationDate' /><Value Type='DateTime'>2014-12-01</Value></Geq></Where></Query></View>" # Newest
#$qry.ViewXml = "<View><Query><Where><And><Leq><FieldRef Name='ActivationDate' /><Value Type='DateTime'>2014-11-20</Value></Leq><Geq><FieldRef Name='ActivationDate' /><Value Type='DateTime'>2008-12-31</Value></Geq></And></Where></Query></View>" # Old
#$qry.ViewXml = "<View><Query><Where><Leq><FieldRef Name='ActivationDate' /><Value Type='DateTime'>2008-12-30</Value></Leq></Where></Query></View>" # Oldest

# Query "content Ops Work Queue"
#$qry = new-object Microsoft.SharePoint.Client.CamlQuery
#$qry.ViewXml = "<View><Query><Where><Leq><FieldRef Name='Created' /><Value Type='DateTime'>2017-05-01</Value></Leq></Where></Query></View>"
#$qry.ViewXml = "<View><Query><Where><Geq><FieldRef Name='Created' /><Value Type='DateTime'>2017-09-01</Value></Geq></Where></Query></View>"
#$qry.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Asset_x0020_ID' /><Value Type='Text'>AS3604</Value></Eq></Where></Query></View>"

# Query entire list
$qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()

$ListItems = $List.GetItems($qry)
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

        $percComplete = ($j/$ListItems.Count*100)
        $percComplete = [math]::Round($percComplete,4)
        Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete
 
           $itemWfInstances = $WorkflowInstanceService.EnumerateInstancesForListItem($List.Id, $ListItems[$j].Id)
           $ClientContext.Load($itemWfInstances)
           $ClientContext.ExecuteQuery()
           for ($k=0;$k -lt $itemWfInstances.Count;$k++)
           {
                if ($itemWfInstances[$k].Status -eq "Suspended" -and ($itemWfInstances[$k].WorkflowSubscriptionId -eq "8ca76f14-7408-4793-bf9b-d0e1c2061f02" -or $itemWfInstances[$k].WorkflowSubscriptionId -eq "4ed0a944-3b01-44e4-a1af-1d7e54cae951")) {
                   
                    try {
                        $counter++;
                        #$WorkflowInstanceService.TerminateWorkflow($itemWfInstances[$k])
                        #Write-Host "Worfklow terminated on" $ListItems[$j].Id
                        Write-Host "Subcription ID" $itemWfInstances[$k].WorkflowSubscriptionId
                        
                        #$ListItems[$j].Update()
                        #$ClientContext.ExecuteQuery()

                        $props = New-Object PSObject -Property @{
                            "Asset ID" = $ListItems[$j]["Asset_x0020_ID"]
                            "WF ID" = $itemWfInstances[$k].WorkflowSubscriptionId
                        }
                        $csvStr += $props
                    } catch {
                        Write-Host "Error terminating workflow on" $ListItems[$j].Id "Details: $_"
                    }
                }
           }
       }
    $csvStr | Export-Csv C:\Users\AS255108\Desktop\SuspendedWorkflows_prod.csv
    Write-Host $counter "Items Found"
