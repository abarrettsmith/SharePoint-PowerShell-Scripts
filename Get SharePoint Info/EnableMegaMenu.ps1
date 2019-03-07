Install-Module SharePointPnPPowerShellOnline

Connect-PnPOnline -Url https://wizbick.sharepoint.com
$web = Get-PnPWeb
$web.MegaMenuEnabled = $true
$web.Update()
Invoke-PnPQuery