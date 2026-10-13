$nssm    = "C:\Program Files\nssm-2.24\nssm-2.24\win64\nssm.exe"
$svc     = "DnBusyMail"
$appDir  = "D:\GCTI\DnBusyMail"                      # copy Watch-DnBusyMail.ps1 here
$script  = "$appDir\Watch-DnBusyMail.ps1"
$logDir  = "$appDir\logs"

# Domain account the service runs as. It needs: read on the share, Send As on -From, AD read (for -SendToAgents),
# and (if -SqlServer is used) SELECT on GCTI_Confserv.dbo.cfg_person. LocalSystem will NOT work with EWS / the share.
$svcUser = "DOMAIN\svc_dnbusymail"                   # <-- change

# Same arguments you tested by hand. Must use -UseEws (or -SmtpServer): Outlook COM does not work in a service.
$scriptArgs = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" ' -f $script) +
              '-SharePath "\\PRIMARY-SERVER\DnBusy","\\BACKUP-SERVER\DnBusy" ' +     # <-- the output share of each monitor service
              '-From "support-group@example.com" ' +
              '-TestTo "test-user@example.com" ' +
              '-SendToAgents -UseEws ' +
              '-Cc "manager1@example.com","support-group@example.com" ' +
              '-IsccTo "ops1@example.com","ops2@example.com" -IsccCc "teamlead@example.com"'

if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

# --- Install (only if the service does not exist yet) ---
if (-not (Get-Service -Name $svc -ErrorAction SilentlyContinue)) {
    & $nssm install $svc "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
}

# --- Configure (safe to re-run) ---
& $nssm set $svc Application   "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
& $nssm set $svc AppParameters $scriptArgs
& $nssm set $svc AppDirectory  $appDir
& $nssm set $svc Start         SERVICE_AUTO_START
& $nssm set $svc AppExit       Default Restart
& $nssm set $svc AppStdout     "$logDir\service_stdout.log"
& $nssm set $svc AppStderr     "$logDir\service_stderr.log"
& $nssm set $svc AppRotateFiles 1
& $nssm set $svc AppRotateBytes 10485760

# Run as the domain account (password is asked here, not stored in this file)
$cred = Get-Credential -UserName $svcUser -Message "Password for the service account"
& $nssm set $svc ObjectName $cred.UserName $cred.GetNetworkCredential().Password

# --- Start and verify ---
if ((Get-Service $svc).Status -eq 'Running') { Restart-Service $svc } else { Start-Service $svc }
Start-Sleep -Seconds 2
Get-Service $svc
& $nssm get $svc AppParameters
