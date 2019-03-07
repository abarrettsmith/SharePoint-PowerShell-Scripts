#Script variables
$CSVTerms = Import-CSV C:\Users\AS255108\Desktop\PowerShellScripts\Terms.csv
$Site = "https://teradata.sharepoint.com/sites/itdev"
$GroupName = "Taxonomy"
$TermSetName = "V8"

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

#Bind to Terms
$Terms = $TermSet.Terms
$Context.Load($Terms)
$Context.ExecuteQuery()

#Create the file and add headings
$OutputFile = "TaxonomyCrossReference.csv"
$file = New-Object System.IO.StreamWriter($OutputFile)
$file.Writeline("Term,GUID,Node");

$csvVal = $CSVTerms
$pattern = '[^a-zA-Z0-9]'

foreach($row in $csvVal) 
{
    $row.Name = $row.Name-replace $pattern, '';
}

#Get Node from original taxonomy
function GetNode($TermName) {
    
    $TermName = $TermName -replace $pattern, ''
    
    $TermRow = $csvVal|where{$_."Name" -clike $TermName -and $_."Name"} 

    $TermRow.Node
    return
}

#Recursive function to get terms
function GetTerms([Microsoft.SharePoint.Client.Taxonomy.Term] $term) {

    $SubTerms = $term.Terms;
    $Context.Load($SubTerms);
    $Context.ExecuteQuery();

    Foreach ($SubTerm in $SubTerms) {
        
        $SubTermNode = GetNode($SubTerm.Name);
        
        $file.Writeline($SubTerm.Name.Replace(",", "") + "," + $SubTerm.Id + "," + $SubTermNode);
        Write-Host $SubTerm.Name - $SubTermNode

        
        GetTerms($SubTerm);
    }
}

#Get terms and write them to the file
Foreach ($Term in $Terms) {

    $TermNode = GetNode($Term.Name);

    $file.WriteLine($Term.Name.Replace(",", "") + "," + $Term.Id + "," + $TermNode);
    Write-Host $Term.Name - $TermNode

    GetTerms($Term);
}

$file.Flush();
$file.Close();
Write-Host ""
Write-Host "Export Complete."
Read-Host