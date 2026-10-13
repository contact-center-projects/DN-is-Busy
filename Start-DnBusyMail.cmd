@echo off
rem Starts the DN-is-busy mail script from a share. No admin rights, no service.
rem Do not start it twice (two copies would race on the same state file).

set "SCRIPT=\\FILE-SERVER\Share\DnBusyMail\Watch-DnBusyMail.ps1"
set "STATE=\\FILE-SERVER\Share\DnBusyMail\state"
rem Host of the Genesys Config Server database (maps numeric agent values / employee_id to user names).
set "SQLSERVER=PUT-SQL-HOST-HERE"

start "DnBusyMail" powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -SharePath "\\PRIMARY-SERVER\DNBusyMonitor-Primary","\\BACKUP-SERVER\DNBusyMonitor-Backup" ^
  -From "support-group@example.com" ^
  -TestTo "test-user@example.com" ^
  -SendToAgents -UseEws ^
  -DedupeMinutes 60 ^
  -CcManager ^
  -SqlServer "%SQLSERVER%" ^
  -Cc "cc-user@example.com","support-group@example.com" ^
  -IsccTo "iscc-owner1@example.com","iscc-owner2@example.com" ^
  -IsccCc "support-group@example.com" ^
  -StateFolder "%STATE%"
