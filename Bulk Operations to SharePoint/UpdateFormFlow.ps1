# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll"
 
# Specify tenant admin and site URL
$SiteUrl = "https://teradata.sharepoint.com/sites/COMPASDEV/"
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

$qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()

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

        # Set Fields
        $DocumentPanel = $ListItems[$j]["Document_x0020_Panel"]
        $IPDocumentPanel = $ListItems[$j]["IP_x0020_Document_x0020_Panel"]
        $URLPane = $ListItems[$j]["URL_x0020_Pane"]
        $IPURLPane = $ListItems[$j]["IP_x0020_URL_x0020_Pane"]

        # If 'No' side of form
        if ($IPDocumentPanel -contains '<NewAssetTitle type="System.String"></NewAssetTitle>' -and $IPDocumentPanel -contains '<EnhAssetTitle type="System.String"></EnhAssetTitle>' -and $IPURLPane -contains '<NewAssetURL type="System.String"></NewAssetURL>' -and $IPURLPane -contains '<EnhAssetURL type="System.String"></EnhAssetURL>') {
            
            $ListItems[$j]["FormFlow"] = "No"
        }
        # If 'Yes' side of form
        elseif ($DocumentPanel -contains '<NewRequestedTitle type="System.String"></NewRequestedTitle>' -and $URLPane -contains '<NewAssetURL type="System.String"></NewAssetURL>') {
            
            $ListItems[$j]["FormFlow"] = "Yes"
        }

        $ListItems[$j].Update()
        $ClientContext.ExecuteQuery()

        Write-Host $ListItems[$j]["Asset_x0020_ID"] "updated."
    }

Write-Host $counter "Items Updated"
