param(
    [Parameter(Mandatory = $true)][string]$SmtpServer,
    [Parameter(Mandatory = $true)][string]$From,
    [Parameter(Mandatory = $true)][string]$To,
    [int]$Port = 25
)

Send-MailMessage -SmtpServer $SmtpServer -Port $Port -From $From -To $To `
    -Subject "Test mail from $env:COMPUTERNAME" `
    -Body "Test sent at $(Get-Date) from $env:COMPUTERNAME as $From." `
    -ErrorAction Stop

Write-Host "Sent OK: $From -> $To via ${SmtpServer}:$Port"
