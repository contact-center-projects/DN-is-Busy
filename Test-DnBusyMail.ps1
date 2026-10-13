<#
.SYNOPSIS
    Feasibility test: can this server send mail as the group address, without Outlook?
    Standalone - does NOT touch the DnBusy monitor service.

.EXAMPLE
    # 1. network check only (nothing is sent)
    .\Test-DnBusyMail.ps1 -SmtpServer smtp.mycompany.local -TestConnectionOnly

    # 2. send one test mail to yourself, "from" the group address (anonymous relay, port 25)
    .\Test-DnBusyMail.ps1 -SmtpServer smtp.mycompany.local -From "ccgroup@mycompany.com" -To "me@mycompany.com"

    # 3. authenticated submission with STARTTLS
    .\Test-DnBusyMail.ps1 -SmtpServer smtp.mycompany.local -Port 587 -UseSsl -Credential (Get-Credential) `
        -From "ccgroup@mycompany.com" -To "me@mycompany.com"

    # 4. also look up an agent's email in AD (read-only)
    .\Test-DnBusyMail.ps1 -LookupAgent k000001 -SmtpServer x -TestConnectionOnly
#>
[CmdletBinding()]
param(
    [string]$SmtpServer,
    [int]$Port = 25,
    [switch]$UseSsl,
    [System.Management.Automation.PSCredential]$Credential,
    [string]$From,
    [string]$To,
    [string]$Subject = "DnBusy mail test",
    [string]$LookupAgent,
    [switch]$TestConnectionOnly
)

$ErrorActionPreference = "Stop"

if ($LookupAgent) {
    Write-Host "--- AD lookup for '$LookupAgent' ---"
    $ds = New-Object System.DirectoryServices.DirectorySearcher
    $ds.Filter = "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$LookupAgent)(employeeID=$LookupAgent)(employeeNumber=$LookupAgent)))"
    [void]$ds.PropertiesToLoad.AddRange(@("sAMAccountName", "mail", "displayName", "employeeID"))
    $res = $ds.FindAll()
    if ($res.Count -eq 0) { Write-Warning "No AD user found by sAMAccountName / employeeID / employeeNumber." }
    foreach ($r in $res) {
        "{0} | {1} | {2} | employeeID={3}" -f $r.Properties["samaccountname"], $r.Properties["displayname"], $r.Properties["mail"], $r.Properties["employeeid"]
    }
}

if (-not $SmtpServer) { return }

Write-Host "--- Network check: $SmtpServer`:$Port ---"
$t = Test-NetConnection -ComputerName $SmtpServer -Port $Port -WarningAction SilentlyContinue
"{0}:{1} reachable = {2}" -f $SmtpServer, $Port, $t.TcpTestSucceeded
if ($TestConnectionOnly -or -not $t.TcpTestSucceeded) { return }

if (-not $From -or -not $To) { throw "Specify -From (the group address) and -To (your own address)." }

Write-Host "--- Sending test mail: $From -> $To ---"
$msg = New-Object System.Net.Mail.MailMessage($From, $To)
$msg.Subject = $Subject
$msg.Body    = "Test from $env:COMPUTERNAME at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'). Sent without Outlook as $From."

$smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
$smtp.EnableSsl = [bool]$UseSsl
if ($Credential) { $smtp.Credentials = $Credential.GetNetworkCredential() } else { $smtp.UseDefaultCredentials = $false }
try {
    $smtp.Send($msg)
    Write-Host "SENT OK - check the inbox of $To (and its spam folder)."
}
catch {
    Write-Warning "Send failed: $($_.Exception.Message)"
    if ($_.Exception.InnerException) { Write-Warning $_.Exception.InnerException.Message }
}
finally { $msg.Dispose(); $smtp.Dispose() }
