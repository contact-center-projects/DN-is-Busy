# DN is Busy

PowerShell tooling that detects Genesys Universal Routing Server (URS) "DN is busy" errors in the URS log files and notifies the affected agent (or a support group for ISCC transfer errors) by mail.

All server names, addresses and internal links in this repository are placeholders (`PRIMARY-SERVER`, `example.com`, ...). Replace them with your own values.

## Components

| File | Purpose |
|---|---|
| `Watch-DnBusyEvents.v2.ps1` | Monitor: reads the rotating URS logs and writes one `DNBUSY\|...` record per event to a daily file |
| `Watch-DnBusyMail.ps1` | Mail: reads those records and sends one mail per event (threshold, dedupe, manager in Cc, DB/AD/EWS address lookup) |
| `Start-DnBusyMail.cmd` | Starts the mail script from a share (no admin rights, no service) |
| `nssm-conf.ps1`, `nssm-conf-backup.ps1`, `nssm-mail-conf.ps1` | Install the scripts as Windows services with NSSM |
| `Send-TestMail*.ps1`, `Test-DnBusyMail.ps1` | Small helpers to test mail sending (SMTP, EWS, Outlook) and AD lookup |
| `Watch-DnBusyEvents.ps1` | First version of the monitor (superseded by v2) |

## Documentation

- [ARCHITECTURE.md](ARCHITECTURE.md) – full description (English)
- [ARCHITECTURE.he.md](ARCHITECTURE.he.md) – full description (Hebrew)
- [ARCHITECTURE-flow.md](ARCHITECTURE-flow.md) – the flow in one diagram

## Quick start

```powershell
# 1. Monitor (on the URS server)
.\Watch-DnBusyEvents.v2.ps1 -LogFolder "D:\Genesys_Logs\UniversalRoutingServer_Backup" -OutputFolder "D:\Monitoring\DnBusy"

# 2. Mail, safe mode: everything goes to -TestTo
.\Watch-DnBusyMail.ps1 -SharePath "\\PRIMARY-SERVER\DnBusy" -From "support-group@example.com" -TestTo "test-user@example.com" -UseEws

# 3. Real sending: add -SendToAgents (and -Cc, -CcManager, -IsccTo, -IsccCc as needed)
```

The mail wording inside `Watch-DnBusyMail.ps1` (`Send-EventMail`) is in Hebrew.
