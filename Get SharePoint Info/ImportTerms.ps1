#Script variables
$Terms = Import-CSV C:\Users\AS255108\Desktop\PowerShellScripts\Terms.csv
$Site = "https://teradata.sharepoint.com"
$GroupName = "Sandbox"
$TermSetName = "Taxonomy"

#Adding references to SharePoint client assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"
$Username = Read-Host -Prompt "Enter your username"
$Password = Read-Host -Prompt "Enter your password" -AsSecureString
$Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username,$Password)

#Bind to MMS
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($Site)
$Context.Credentials = $Creds
$MMS = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($Context)
$Context.Load($MMS)
$Context.ExecuteQuery()

#Retrieve Term Stores
$TermStores = $MMS.TermStores
$Context.Load($TermStores)
$Context.ExecuteQuery()

#Bind to Term Store
$TermStore = $TermStores[0]
$Context.Load($TermStore)
$Context.ExecuteQuery()

#Bind to Group
$Group = $TermStore.Groups.GetByName($GroupName)
$Context.Load($Group)
$Context.ExecuteQuery()

#Bind to Term Set
$TermSet = $Group.TermSets.GetByName($TermSetName)
$Context.Load($TermSet)
$Context.ExecuteQuery()

#Define Node Variables
$Didget1 = 0
$Didget2 = 0
$Didget3 = 0
$Didget4 = 0
$Didget5 = 0
$ErrorCount = 0

#Add Node Values to Term Store
ForEach ($Term in $Terms) {
    
    #Get Number of . in Node
    $NumberOfPeriods = ($Term.Node.Split([string[]]@("."),[StringSplitOptions]"None")).Count - 1

    #Assign value appropriately
    if ($NumberOfPeriods -eq 0) {
        
        $Didget1 += 1

        $Level1TermOrder += $Term.Name

        Write-Host "Creating Node" $Term.Node "as" $Term.Name

        #Add value to term store
        $Level1TermAdd = $TermSet.CreateTerm(($Term.Name), 1033, [System.Guid]::NewGuid().ToString())
        $Context.Load($Level1TermAdd)
        $Context.ExecuteQuery()
    }
    elseif ($NumberOfPeriods -eq 1) {
        
        $Didget2 += 1

        $Level2TermOrder += $Term.Name

        Write-Host "Creating Node" $Term.Node "as" $Term.Name

        #Add value to term store
        $Level2TermAdd = $Level1TermAdd.CreateTerm(($Term.Name), 1033, [System.Guid]::NewGuid().ToString())
        $Context.Load($Level2TermAdd)
        $Context.ExecuteQuery()
    }
    elseif ($NumberOfPeriods -eq 2) {

        $Didget3 += 1
        $Level3TermOrder += $Term.Name

        Write-Host "Creating Node" $Term.Node "as" $Term.Name

        #Add value to term store
        $Level3TermAdd = $Level2TermAdd.CreateTerm(($Term.Name), 1033, [System.Guid]::NewGuid().ToString())
        $Context.Load($Level3TermAdd)
        $Context.ExecuteQuery()
    }
    elseif ($NumberOfPeriods -eq 3) {
        
        $Didget4 += 1

        $Level4TermOrder += $Term.Name

        Write-Host "Creating Node" $Term.Node "as" $Term.Name

        #Add value to term store
        $Level4TermAdd = $Level3TermAdd.CreateTerm(($Term.Name), 1033, [System.Guid]::NewGuid().ToString())
        $Context.Load($Level4TermAdd)
        $Context.ExecuteQuery()
    } 
    elseif ($NumberOfPeriods -eq 4) {
        
        $Didget5 += 1

        $Level5TermOrder += $Term.Name

        Write-Host "Creating Node" $Term.Node "as" $Term.Name

        #Add value to term store
        $Level5TermAdd = $Level4TermAdd.CreateTerm(($Term.Name), 1033, [System.Guid]::NewGuid().ToString())
        $Context.Load($Level5TermAdd)
        $Context.ExecuteQuery()
    }
    else {
        #Notify of error
        Write-Host "Something went wrong"
        Write-Host $Term.Node
        Write-Host $Term.Name
        $ErrorCount += 1
    }


}

Write-Host ""
Write-Host "Import Complete."

Read-Host