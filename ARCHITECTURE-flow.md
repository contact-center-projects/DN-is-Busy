# DN is Busy – Simple Flow

(The detailed version is in [ARCHITECTURE.md](ARCHITECTURE.md); Hebrew: [ARCHITECTURE.he.md](ARCHITECTURE.he.md).)

```mermaid
flowchart TD
    A["Genesys URS<br/>writes log files"] --> B["Monitor service (Primary and Backup)<br/>Watch-DnBusyEvents.v2.ps1<br/>finds 'DN is busy' errors"]
    B --> C[("dnbusy_yyyyMMdd.log<br/>one DNBUSY| line per error")]
    C --> D["Mail script<br/>Watch-DnBusyMail.ps1<br/>reads the lines"]
    D --> T{"Same DN twice<br/>within 10 minutes?"}
    T -->|no| W["wait"]
    T -->|yes| E{"Which type<br/>of error?"}
    E -->|"Agent error<br/>has AGENT"| F["Find agent's address<br/>(DB, AD, EWS)"]
    F --> G["Mail to the agent<br/>+ Cc + manager"]
    E -->|"ISCC error<br/>no AGENT"| H["Mail to the ISCC list<br/>IsccTo + IsccCc"]
```

Same flow as text:

```
URS logs
   |
   v
Monitor service (finds "DN is busy")
   |
   v
dnbusy_yyyyMMdd.log   (one DNBUSY| line per error)
   |
   v
Mail script  --- same DN twice within 10 minutes? ---
   |
   +--- Agent error (has AGENT) ---> find agent's address (DB / AD / EWS) ---> mail to AGENT (+ Cc + manager)
   |
   +--- ISCC error (no AGENT)   ---> mail to IsccTo (+ IsccCc)
```

If the agent's address is not found, or `-SendToAgents` is not used, the mail goes to `-TestTo` instead.
