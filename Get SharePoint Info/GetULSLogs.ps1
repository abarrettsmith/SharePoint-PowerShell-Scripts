Connect-SPOService -Url https://teradata.sharepoint.com/ -Credential adam.smith@teradata.com

$endTimeinUTC = Get-SPOTenantLogLastAvailableTimeInUtc
$startTimeinUTC = $endTimeinUTC.AddDays(-14)
$tenantlogs = Get-SPOTenantLogEntry -StartTimeinUtc $startTimeinUTC -EndTimeInUtc $endTimeinUTC -CorrelationId 86d4b59d-30ca-3000-9fd1-d3e6aa0a0cc0