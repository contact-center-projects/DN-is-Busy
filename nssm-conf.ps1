$nssm    = "C:\Program Files\nssm-2.24\nssm-2.24\win64\nssm.exe"
$svc     = "DnBusyMonitor-Primary"
$appDir  = "D:\GCTI\DNisBusyMonitor-Primary"
$script  = "$appDir\Watch-DnBusyEvents.v2.ps1"
$outDir  = "D:\Genesys_Logs\DNisBusyMonitor-Primary"

$scriptArgs = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" ' -f $script) +
              '-LogFolder "D:\Genesys_Logs\UniversalRoutingServer_Backup" ' +
              '-FilePattern "UniversalRoutingServer_Backup*.log" ' +
              ('-OutputFolder "{0}" ' -f $outDir) +
              '-OutputPrefix "dnbusy" -RetentionDays 7 -PollIntervalSeconds 2'

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
& $nssm set $svc AppStdout     "$outDir\service_primary_stdout.log"
& $nssm set $svc AppStderr     "$outDir\service_primary_stderr.log"

# --- Start and verify ---
if ((Get-Service $svc).Status -eq 'Running') { Restart-Service $svc } else { Start-Service $svc }
Start-Sleep -Seconds 2
Get-Service $svc
& $nssm get $svc AppParameters
