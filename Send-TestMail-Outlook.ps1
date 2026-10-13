param(
    [Parameter(Mandatory = $true)][string]$From,   # the group / shared mailbox address
    [Parameter(Mandatory = $true)][string]$To,     # use your OWN address so the mail comes back to your inbox
    [int]$WaitSeconds = 60
)

$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")

# --- 1. what Outlook knows about its accounts ---
Write-Host "--- Outlook accounts ---"
$types = @{ 0 = "Exchange"; 1 = "IMAP"; 2 = "POP3"; 3 = "HTTP"; 4 = "EAS" }
foreach ($a in $ns.Accounts) {
    $line = "{0} | {1} | type={2}" -f $a.DisplayName, $a.SmtpAddress, $types[[int]$a.AccountType]
    try { if ($a.ExchangeMailboxServerName) { $line += " | ExchangeServer=" + $a.ExchangeMailboxServerName } } catch {}
    $line
}

# --- 2. send ---
$subject = "Test mail from $env:COMPUTERNAME $(Get-Date -Format 'HHmmss')"
$m = $ol.CreateItem(0)
$m.To      = $To
$m.Subject = $subject
$m.Body    = "Test sent at $(Get-Date) from $env:COMPUTERNAME on behalf of $From."
$m.SentOnBehalfOfName = $From
$m.Send()
Write-Host "Sent via Outlook: $From -> $To"

# --- 3. read the Received: headers from the copy that arrives in the inbox ---
Write-Host "--- Waiting up to $WaitSeconds s for it to arrive, to read the server headers ---"
$PR_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
$inbox = $ns.GetDefaultFolder(6)
$deadline = (Get-Date).AddSeconds($WaitSeconds)
$found = $null
while (-not $found -and (Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 3
    $found = $inbox.Items | Where-Object { $_.Subject -eq $subject } | Select-Object -First 1
}
if (-not $found) { Write-Warning "Mail not seen in the inbox yet - open it manually and use File > Properties > Internet headers."; return }

$hdr = $found.PropertyAccessor.GetProperty($PR_HEADERS)
$show = $false
foreach ($l in ($hdr -split "`r?`n")) {
    if ($l -match '^\S') { $show = ($l -match '^(Received|From|X-MS-Exchange-Organization-AuthSource|Return-Path):') }
    if ($show) { $l }   # continuation lines (leading whitespace) stay with their header, so the "by <server>" part is printed
}
