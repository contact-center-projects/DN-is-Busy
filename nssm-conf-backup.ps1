# Run on the BACKUP server. Same as nssm-conf.ps1, with the service/folder names for the backup server.
$nssm    = "C:\Program Files\nssm-2.24\nssm-2.24\win64\nssm.exe"
$svc     = "DnBusyMonitor-Backup"
$appDir  = "D:\GCTI\DNisBusyMonitor-Backup"          # copy Watch-DnBusyEvents.v2.ps1 here
$script  = "$appDir\Watch-DnBusyEvents.v2.ps1"
$outDir  = "D:\Genesys_Logs\DNisBusyMonitor-Backup"  # share this folder so the mail script can read it

# <-- check these two on the backup server: where its URS writes the logs, and the file name pattern
$logDir  = "D:\Genesys_Logs\UniversalRoutingServer_Backup"
$pattern = "UniversalRoutingServer_Backup*.log"

$scriptArgs = ('-NoProfile -ExecutionPolicy Bypass -File "{0}" ' -f $script) +
              ('-LogFolder "{0}" ' -f $logDir) +
              ('-FilePattern "{0}" ' -f $pattern) +
              ('-OutputFolder "{0}" ' -f $outDir) +
              '-OutputPrefix "dnbusy" -RetentionDays 7 -PollIntervalSeconds 2'

foreach ($d in $appDir, $outDir) { if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }

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
& $nssm set $svc AppStdout     "$outDir\service_backup_stdout.log"
& $nssm set $svc AppStderr     "$outDir\service_backup_stderr.log"

# --- Start and verify ---
if ((Get-Service $svc).Status -eq 'Running') { Restart-Service $svc } else { Start-Service $svc }
Start-Sleep -Seconds 2
Get-Service $svc
& $nssm get $svc AppParameters
