Set-ExecutionPolicy RemoteSigned
$UserCredential = Get-Credential
$ConnectionStatus = 'Success'

try { 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
Remove-PSSession $Session
}
catch {
$ConnectionStatus = 'Error'
}

Read-Host 'Connection Status: ' $ConnectionStatus