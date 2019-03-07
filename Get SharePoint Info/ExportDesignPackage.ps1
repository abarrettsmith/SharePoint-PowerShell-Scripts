Add-PSSnapin "Microsoft.SharePoint.PowerShell" -EA 0; cls
Set-ExecutionPolicy Unrestricted

function Export-SPDesignPackage {
    #written by Ingo Karstein (http://blog.karstein-consulting.com)  
    # v1.0

    #You can copy this function to your own script file or use the file as PowerShell module

    #See ... for details

    [CmdLetBinding(DefaultParameterSetName="Default")]
    param(
        [parameter(Mandatory=$true, Position=0, ParameterSetName="Default", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string]
        $SiteUrl="https://teradata.sharepoint.com/IPUAT/",

        [parameter(Mandatory=$true, Position=0, ParameterSetName="Site", ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Microsoft.SharePoint.SPSite]
        $Site=$null,

        [parameter(Mandatory=$false, Position=1, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, Position=1, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [string]
        $ExportFileName = "IPUAT",

        [parameter(Mandatory=$false, Position=1, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, Position=1, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [string]
        $ExportFolder = "C:\Users\AS255108\Desktop\DesignPackage",

        [parameter(Mandatory=$false, Position=1, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, Position=1, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [string]
        $UseTempFileForExportWithExtension = "",

        [parameter(Mandatory=$false, Position=2, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, Position=2, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [string]
        $PackageName = "IPUATDesignPackage",

        [parameter(Mandatory=$false, Position=3, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, Position=3, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [switch]
        $IncludeSearchConfig = $false,

        [parameter(Mandatory=$false, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [switch]
        $DisposeSiteObject = $true,

        [parameter(Mandatory=$false, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [System.Management.Automation.PSCredential]
        $DownloadCredentials = $Null,

        [parameter(Mandatory=$false, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [switch]
        $OverwriteExistingFiles = $false,

        [parameter(Mandatory=$false, ParameterSetName="Default", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [parameter(Mandatory=$false, ParameterSetName="Site", ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true)]
        [switch]
        $UseExportFileNumbering = $false
    )

    begin {
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Publishing") | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName("System.Net") | Out-Null

        $a = 0
        if( !([string]::IsNullOrEmpty($ExportFileName)) ) { $a++ }
        if( !([string]::IsNullOrEmpty($ExportFolder)) ) { $a++ }
        if( !([string]::IsNullOrEmpty($UseTempFileForExportWithExtension)) ) { $a++ }

        if( $a -gt 1 ) {
            $e = new-object System.Exception("Cannot use parameters ""ExportFileName"", ""ExportFolder"" and/or ""UseTempFileForExportWithExtension"" side by side. Please choose one!")
            $err  =new-object System.Management.Automation.ErrorRecord($e, "Cannot use parameters ""ExportFileName"", ""ExportFolder"" and/or ""UseTempFileForExportWithExtension"" side by side. Please choose one!", "InvalidArgument", $null)
            $PSCmdlet.ThrowTerminatingError($err)
        }

        if(!([string]::IsNullOrEmpty($ExportFolder))) {
            if( !(Test-Path $ExportFolder -PathType Container) ) {
                $e = new-object System.Exception("The folder specified in parameter ""ExportFolder"" does not exist!")
                $err  =new-object System.Management.Automation.ErrorRecord($e, "The folder specified in parameter ""ExportFolder"" does not exist!", "InvalidArgument", $null)
                $PSCmdlet.ThrowTerminatingError($err)
            }
        }

        if(!([string]::IsNullOrEmpty($ExportFileName))) {
            $path = split-path $ExportFileName
            if( !(Test-Path $path -PathType Container) ) {
                $e = new-object System.Exception("The folder of the filename specified in parameter ""ExportFileName"" does not exist!")
                $err  =new-object System.Management.Automation.ErrorRecord($e, "The folder of the filename specified in parameter ""ExportFileName"" does not exist!", "InvalidArgument", $null)
                $PSCmdlet.ThrowTerminatingError($err)
            }
        }

        $count = 0
    }

    process {
        $count++

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
            if($localSite -ne $null ) {
                $resultObject | Add-Member -MemberType NoteProperty -Name "SiteFound" -Value $true -Force

                try {
                    if( !([string]::IsNullOrEmpty($PackageName)) ) {
                        $package = [Microsoft.SharePoint.Publishing.DesignPackage]::Export($localSite, $PackageName, $IncludeSearchConfig)
                    } else {
                        $package = [Microsoft.SharePoint.Publishing.DesignPackage]::Export($localSite, $IncludeSearchConfig)
                    }
                    $resultObject | Add-Member -MemberType NoteProperty -Name "ExportError" -Value [System.Exception]$null -Force
                } catch {
                    $resultObject | Add-Member -MemberType NoteProperty -Name "ExportError" -Value $_.Exception -Force
                    return
                }

                $packageFileName = "{0}-{1}.{2}.wsp" -f ($package.PackageName, $package.MajorVersion, $package.MinorVersion)

                $resultObject | Add-Member -MemberType NoteProperty -Name "PackageFileName" -Value $packageFileName -Force
                $resultObject | Add-Member -MemberType NoteProperty -Name "PackageName" -Value ($package.PackageName) -Force
                $resultObject | Add-Member -MemberType NoteProperty -Name "PackageMajorVersion" -Value ($package.MajorVersion) -Force
                $resultObject | Add-Member -MemberType NoteProperty -Name "PackageMinorVersion" -Value ($package.MinorVersion) -Force

                $wc = new-object System.Net.WebClient

                $cred = [System.Net.CredentialCache]::DefaultNetworkCredentials
                if( $DownloadCredentials -ne $null ) {
                    $cred = $DownloadCredentials.GetNetworkCredential()
                }

                $wc.Credentials = $cred

                $downloadUrl = $localSite.Url.TrimEnd("/") + "/_catalogs/solutions/" + $packageFileName

                $localFile = ""
                if( !([string]::IsNullOrEmpty($ExportFolder)) ) {
                    $localFile = join-path $ExportFolder $packageFileName
                } else {
                    if( !([string]::IsNullOrEmpty($ExportFileName)) ) {
                        if( $UseExportFileNumbering ) {
                            $path = split-path $ExportFileName
                            $fn = [System.IO.Path]::GetFileNameWithoutExtension($ExportFileName)
                            $ext = [System.IO.Path]::GetExtension($ExportFileName)

                            $localFile = join-path $path ("{0}-{1}{2}" -f @($fn, $count, $ext))
                        } else {
                            $localFile = $ExportFileName
                        }
                    } else {
                        $localFile = join-path ([System.IO.Path]::GetTempPath()) $packageFileName
                    }
                }

                if( !([string]::IsNullOrEmpty($UseTempFileForExportWithExtension)) ) {
                    $localFile = [System.IO.Path]::GetTempFileName() + $UseTempFileForExportWithExtension
                }

                if( Test-Path $localFile ) {
                    if( $OverwriteExistingFiles ) {
                        Remove-Item $localFile -Confirm:$false -Force
                        $resultObject | Add-Member -MemberType NoteProperty -Name "ExportFileOverridden" -Value $true -Force 
                    }
                } else {
                    $resultObject | Add-Member -MemberType NoteProperty -Name "ExportFileOverridden" -Value $false -Force 
                }


                if( !(Test-Path $localFile) ) {
                    $resultObject | Add-Member -MemberType NoteProperty -Name "ExportFile" -Value $localFile -Force

                    try {
                        $wc.DownloadFile($downloadUrl, $localFile)
                        $resultObject | Add-Member -MemberType NoteProperty -Name "DownloadError" -Value [System.Exception]$null -Force
                        $resultObject.Success = $true
                    } catch {
                        $resultObject | Add-Member -MemberType NoteProperty -Name "DownloadError" -Value $_.Exception -Force
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

    begin {}

    process {
        $r = new-object System.Management.Automation.PSObject
        $Hashtable.Keys | % {
            $key = $_
            $value = $Hashtable[$key]
            $r | Add-Member -MemberType NoteProperty -Name $key -Value $value -Force
        }

        return $r
    }

    end {}

}

##########################
# Some samples

$cred = new-object System.Management.Automation.PSCredential( "domain\spfarm", (ConvertTo-SecureString -AsPlainText "Passw0rd" -Force))

$site1 = get-spsite "http://sharepoint.local/publishing"
$site2 = get-spsite "http://sharepoint.local/sites/publishing2"

$site1, $site2 | Export-SPDesignPackage -UseTempFileForExportWithExtension ".wsp" -DownloadCredentials $cred -PackageName "test"

$site1, $site2 | Export-SPDesignPackage -ExportFileName "C:\temp\Package.wsp" -UseExportFileNumbering -IncludeSearchConfig -DisposeSiteObject -OverwriteExistingFiles


(
    @{PackageName="P1"; ExportFileName="C:\temp\p1.wsp"; SiteUrl="http://sharepoint.local/publishing"},
    @{PackageName="P2"; ExportFileName="C:\temp\p2.wsp"; SiteUrl="http://sharepoint.local/sites/publishing2"}
) | New-ObjectFromHashtable | Export-SPDesignPackage

$site2 | Export-SPDesignPackage -ExportFileName "C:\temp\publishing2.wsp"  -IncludeSearchConfig -DisposeSiteObject -OverwriteExistingFiles