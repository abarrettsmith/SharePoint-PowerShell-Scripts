# Add Wave16 references to SharePoint client assemblies and authenticate to Office 365 site - required for CSOM
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")
Add-Type -Path (Resolve-Path "$env:CommonProgramFiles\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll")
 
# Specify tenant admin and site URL
$SiteUrl = "https://teradata.sharepoint.com/sites/COMPAS/"
$UserName = "adam.smith@teradata.com"
$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
 
# Connect to site
$ClientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
$ClientContext.Credentials = $credentials
$ClientContext.ExecuteQuery()

# Create SQL connection
$sqlCon = new-object "System.data.sqlclient.SQLconnection"

# Set connection string
# PROD
$sqlCon.ConnectionString =("Data Source=susazur1001.database.windows.net;Initial Catalog=COMPAS; User ID=COMPASADM;password=Nyirah24%^&*;MultipleActiveResultSets=True;")
# UAT
#$sqlCon.ConnectionString =("Data Source=susazur1000.database.windows.net;Initial Catalog=IP_REPOSITORY; User ID=IPREUSESVC;Password=Nyirah24!@#$;MultipleActiveResultSets=True;")
# DEV
#$sqlCon.ConnectionString =("Data Source=susazur1000.database.windows.net;Initial Catalog=IP_REPOSITORY_DEV; User ID=IPREUSESVC;Password=Nyirah24!@#$;MultipleActiveResultSets=True;")

# Open connection
$sqlCon.Open()

Write-Host "Connected to SQL"

$csv = Import-Csv c:\Users\AS255108\Desktop\EmptyFileNames.csv
$assets = $csv | Select 'AssetId', 'SPListName', 'SPItemId'
$csvStr = @()

Write-Host $assets

# Loop List Items to Start Workflow
For ($j=0; $j -lt $assets.Count + 1; $j++){

    # Print progress
    $percComplete = ($j/($assets.Count+1)*100)
    $percComplete = [math]::Round($percComplete,4)
    Write-Progress -Activity "Sorting Through Items" -Status "%$percComplete Completed" -PercentComplete $percComplete

    # Set variables
    $listName = $assets[$j].'SPListName'
    $itemId = $assets[$j].'SPItemId'
    $assetId = $assets[$j].'AssetId'

    Write-Host "ListName: $listName"
    Write-Host "itemId: $itemId"
    Write-Host "assetId: $assetId"

    # Get List
    $List = $ClientContext.Web.Lists.GetByTitle($listName)
    $ClientContext.Load($List)
    $ClientContext.ExecuteQuery()
    
    # Get Item
    $ListItem = $List.GetItemById($itemId)
    $ClientContext.Load($ListItem)
    $ClientContext.ExecuteQuery()


    if ($listName -ne "ExternalAssets") {
        # Get file name
        $fileName = $ListItem["FileLeafRef"]

        # Update SQL row
        $sqlCmd = new-object "System.data.sqlclient.sqlcommand"
        $sqlCmd.Connection = $sqlCon
        $sqlCmd.CommandTimeout = 600000
        $sqlCmd.CommandText = "UPDATE dbo.Assets SET [FileName] = '$fileName' WHERE AssetId = '$AssetId'"

        $rowsAffected = $sqlCmd.ExecuteNonQuery()
    }

    Write-Host "Updated - $assetId"
}