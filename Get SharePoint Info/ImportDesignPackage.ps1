function Import-SPDesignPackage {
    #written by Ingo Karstein (http://blog.karstein-consulting.com)
    # v1.0

    #You can copy this function to your own script file or use the file as PowerShell module

    #See ... for details

    [CmdLetBinding(DefaultParameterSetName="Default")]
    param(
        [parameter(Mandatory=$true, Position=0, ParameterSetName="Default", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $SiteUrl="",

        [parameter(Mandatory=$true, Position=0, ParameterSetName="Site", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Microsoft.SharePoint.SPSite]
        $Site=$null,

        [parameter(Mandatory=$true, Position=1, ParameterSetName="Default", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$true, Position=1, ParameterSetName="Site", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $ImportFileName = "",

        [parameter(Mandatory=$true, Position=2, ParameterSetName="Default", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$true, Position=2, ParameterSetName="Site", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [bool]
        $Apply = $false,

        [parameter(Mandatory=$false, Position=3, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, Position=3, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [string]
        $PackageName = "",


        [parameter(Mandatory=$false, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [int]
        $MajorVersion = 1,

        [parameter(Mandatory=$false, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [int]
        $MinorVersion = 0
    )

    begin {
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Publishing") | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName("System.Net") | Out-Null
    }

    process {
        $localSite = $null
        $localUrl = ""

        if( $PSCmdlet.ParameterSetName -like "Default" ) {
            $localSite = Get-SPSite $SiteUrl -ErrorAction 0
            $localUrl = $SiteUrl
        } else {
            $localSite = $site
            $localUrl = $site.Url
        }
        
        $resultObject = new-object System.Management.Automation.PSObject
        $resultObject | Add-Member -MemberType NoteProperty -Name "SiteUrl" -Value $localUrl -Force
        $resultObject | Add-Member -MemberType NoteProperty -Name "Success" -Value $false -Force

        try {
            if( !(Test-Path $ImportFileName) ) {
                $resultObject | Add-Member -MemberType NoteProperty -Name "InputFileFound" -Value $false -Force
                return
            } else {
                $resultObject | Add-Member -MemberType NoteProperty -Name "InputFileFound" -Value $true -Force
            }

            if( [System.IO.Path]::GetExtension($ImportFileName) -ne ".wsp" ) {
                $resultObject | Add-Member -MemberType NoteProperty -Name "InputFileExtensionValid" -Value $false -Force
                return
            } else {
                $resultObject | Add-Member -MemberType NoteProperty -Name "InputFileExtensionValid" -Value $true -Force
            }

            if($localSite -ne $null ) {
                $resultObject | Add-Member -MemberType NoteProperty -Name "SiteFound" -Value $true -Force

                $file = [System.IO.Path]::GetFileName($ImportFileName)

                if( [string]::IsNullOrEmpty($PackageName) ) {
                    $PackageName = [System.IO.Path]::GetFileNameWithoutExtension($ImportFileName)
                }

                $solutionFileName = "{0}-v{1}.{2}.wsp" -f ($PackageName, $MajorVersion, $MinorVersion)

                $webFile = $null
                $webFile = $localSite.RootWeb.GetFile("_catalogs/solutions/" + $solutionFileName)

                if( $webFile -ne $null -and $webFile.Exists ) {
                    $resultObject | Add-Member -MemberType NoteProperty -Name "PackageAlreadyExists" -Value $true -Force
                    return
                } else {
                    $resultObject | Add-Member -MemberType NoteProperty -Name "PackageAlreadyExists" -Value $false -Force
                }
                
                $package = new-object Microsoft.SharePoint.Publishing.DesignPackageInfo($file, [Guid]::Empty, $MajorVersion, $MinorVersion)

                $spfolder = $null
                $inputStream = $null

                $installDone = $false

                try {
                    $spfolder = $localSite.RootWeb.RootFolder.SubFolders.Add("tmp_importspdesignpackage_15494B80-89A0-44FF-BA6C-208CB6A053D0")
                    
                    $inputStream = [System.IO.File]::OpenRead($ImportFileName)
                    
                    $spfile = $spfolder.Files.Add($file, $inputStream, $true)
                    
                    $inputStream.Close()
                    $inputStream = $null

                    [Microsoft.SharePoint.Publishing.DesignPackage]::Install($localSite, $package, $spfile.Url)
                    $resultObject | Add-Member -MemberType NoteProperty -Name "InstallError" -Value [System.Exception]$null -Force
                    $installDone = $true
                } catch {
                    $resultObject | Add-Member -MemberType NoteProperty -Name "InstallError" -Value $_.Exception -Force
                } finally {
                    if( $spfolder -ne $null ) { $spfolder.Delete() }
                    if( $inputStream -ne $null ) { $inputStream.Close() }
                }
                
                if( $installDone ) {
                    if( $Apply ) {
                        try {
                            [Microsoft.SharePoint.Publishing.DesignPackage]::Apply($localSite, $package)
                            $resultObject | Add-Member -MemberType NoteProperty -Name "ApplyError" -Value [System.Exception]$null -Force
                            $resultObject.Success = $true
                        } catch {
                            $resultObject | Add-Member -MemberType NoteProperty -Name "ApplyError" -Value $_.Exception -Force
                        }
                    } else {
                        $resultObject.Success = $true
                    }
                }
            } else {
                $resultObject | Add-Member -MemberType NoteProperty -Name "SiteFound" -Value $false -Force
            }

        } finally {
            if( $PSCmdlet.ParameterSetName -like "Default*" ) {
                if( $localSite -ne $null ) {
                    $localSite.Dispose()
                }
            } else {
                if( $disposeSiteObject -eq $true -and $site -ne $null -and $site -is [Microsoft.SharePoint.SPSite] ) {
                    $site.Dispose()
                }
            }
            $resultObject
        }
    }

    end {}
}

function New-ObjectFromHashtable {
    #written by Ingo Karstein (http://blog.karstein-consulting.com)
    # v1.0

    #Use this function to convert a hashtable to a PowerShell object ("PSObject"), e.g. for using hashtables for property name binding in
    # PowerShell pipelines

    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [Hashtable]
        $Hashtable
    )

    begin {
        $results = @()
    }

    process {
        $r = new-object System.Management.Automation.PSObject
        $Hashtable.Keys | % {
            $key = $_
            $value = $Hashtable[$key]
            $r | Add-Member -MemberType NoteProperty -Name $key -Value $value -Force
        }

        $results += $r
    }

    end {
        $results
    }

}

$r = Import-SPDesignPackage -SiteUrl "http://sharepoint.local/publishingxyz" -ImportFileName "C:\temp\publishing2.wsp" -PackageName "P2" -Apply $true

(
    @{ SiteUrl = "http://sharepoint.local/publishing";
       ImportFileName = "C:\temp\publishing2.wsp";
       PackageName = "P2";
       Apply=$true
    },
    @{ SiteUrl = "http://sharepoint.local/sites/publishing2";
       ImportFileName = "C:\temp\publishing2.wsp";
       PackageName = "P2";
       Apply=$true
    }
) | New-ObjectFromHashtable | Import-SPDesignPackage