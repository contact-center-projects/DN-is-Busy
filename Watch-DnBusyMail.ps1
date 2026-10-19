<#
.SYNOPSIS
    Runs on a PC that has Outlook. Tails the daily DNBUSY file over a share and sends one mail per event.

.DESCRIPTION
    Safe by default: every mail goes to -TestTo. Only with -SendToAgents does it try to mail the
    agent (resolved through Outlook's address book); if an agent cannot be resolved it falls back to -TestTo.
    Retries of the same call (same ConnID + Agent) within -DedupeMinutes produce only one mail.

.EXAMPLE
    .\Watch-DnBusyMail.ps1 -SharePath "\\PRIMARY-SERVER\DnBusy" -From "support-group@example.com" -TestTo "test-user@example.com"
#>
[CmdletBinding()]     # enables -Verbose (agent/manager lookup details are shown only with it)
param(
    [Parameter(Mandatory = $true)][string[]]$SharePath,   # one or more output shares (e.g. primary and backup server)
    [Parameter(Mandatory = $true)][string]$From,
    [Parameter(Mandatory = $true)][string]$TestTo,
    [switch]$SendToAgents,
    [string[]]$Cc,                       # CC on mails that go to an agent (separate with , or ;). Default: -From
    [switch]$CcManager,                  # also CC the agent's manager (AD "manager" attribute), in addition to -Cc
    [string[]]$IsccTo,                  # recipients of ISCC (no-agent) mails; used only with -SendToAgents, otherwise they go to -TestTo
    [string[]]$IsccCc,                   # CC on ISCC mails
    [switch]$UseEws,                     # send over Exchange Web Services (HTTPS) with the Windows login - no Outlook, no SMTP
    [string]$EwsUrl = "https://mail.example.com/EWS/Exchange.asmx",
    [string]$SmtpServer,                 # or send by SMTP (no Outlook needed, works with nobody logged on)
    [int]$Port = 25,
    [switch]$UseSsl,
    [switch]$UseCurrentUser,             # authenticate to SMTP as the account the task runs as
    [string]$Prefix = "dnbusy",
    [string]$SqlServer,                  # Config Server DB host; when set, numeric agent values (employee_id) are mapped to user_name
    [string]$SqlDatabase = "GCTI_Confserv",
    [string]$FormLink = "http://forms.example.com/office-form",   # link to the "טופס בית-משרד"; pass -FormLink '' to omit it
    [int]$PollSeconds = 10,
    [int]$MinOccurrences = 2,            # mail only after the same DN is seen this many times within -WithinMinutes (1 = mail on the first occurrence)
    [int]$WithinMinutes = 10,
    [int]$DedupeMinutes = 0,            # 0 = one mail per agent per day; >0 = one mail per agent per that many minutes
    [int]$BurstSeconds = 0,             # ignore repeated detections for the same DN within this many seconds of the first (0 = disabled)
    [string]$StateFolder = "$env:ProgramData\DnBusyMail"   # remembers who was already mailed, so a restart does not re-send
)

$ErrorActionPreference = "Stop"

function ConvertTo-AddressList($Value) {
    @($Value | ForEach-Object { $_ -split '[,;]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}
$SharePath = ConvertTo-AddressList $SharePath     # under "powershell -File" a list arrives as one "a,b" string
if (-not $PSBoundParameters.ContainsKey('Cc')) { $Cc = @($From) }
$script:AgentCc = ConvertTo-AddressList $Cc      # not read as $Cc inside functions: PowerShell names are case-insensitive, it would clash with the local $cc
$IsccTo  = ConvertTo-AddressList $IsccTo
$IsccCc  = ConvertTo-AddressList $IsccCc

$script:Path   = $null
$script:Pos    = 0
$script:Sent   = @{}     # agent -> last mail time
$script:Outlook = $null

function Get-StateFile { Join-Path $StateFolder ("sent_{0}.txt" -f (Get-Date).ToString("yyyyMMdd")) }

function Load-State {
    $h = @{}
    $p = Get-StateFile
    if (Test-Path -LiteralPath $p) {
        foreach ($line in Get-Content -LiteralPath $p) {
            $i = $line.LastIndexOf('|')
            if ($i -gt 0) {
                $t = [datetime]::MinValue
                if ([datetime]::TryParseExact($line.Substring($i + 1), "yyyy-MM-dd HH:mm:ss", $null, [System.Globalization.DateTimeStyles]::None, [ref]$t)) { $h[$line.Substring(0, $i)] = $t }
            }
        }
    }
    return $h
}

function Save-State([string]$Key, [datetime]$When) {
    if (-not (Test-Path $StateFolder)) { New-Item -ItemType Directory -Path $StateFolder -Force | Out-Null }
    Add-Content -LiteralPath (Get-StateFile) -Value ("{0}|{1}" -f $Key, $When.ToString("yyyy-MM-dd HH:mm:ss"))
    # keep only the last 7 days of state files
    Get-ChildItem $StateFolder -Filter "sent_*.txt" -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) } | Remove-Item -Force -ErrorAction SilentlyContinue
}

$script:UserCache = @{}   # employee_id -> @{ Name; At }

# Agent values that are plain numbers are employee_ids; map them to user_name via GCTI_Confserv.dbo.cfg_person.
# Values like k000001 are already user names. Falls back to the original value if not found / not unique / DB error.
function Resolve-UserName([string]$Agent) {
    if (-not $SqlServer -or $Agent -notmatch '^\d+$') { return $Agent }
    $c = $script:UserCache[$Agent]
    if ($c -and ((Get-Date) - $c.At).TotalMinutes -lt 60) { return $c.Name }
    $name = $Agent
    try {
        $cn = New-Object System.Data.SqlClient.SqlConnection("Server=$SqlServer;Database=$SqlDatabase;Integrated Security=SSPI;Connect Timeout=5;Application Name=DnBusyMail")
        $cn.Open()
        try {
            $cmd = $cn.CreateCommand()
            $cmd.CommandText = "SELECT DISTINCT user_name FROM dbo.cfg_person WHERE employee_id = @id"
            $cmd.CommandTimeout = 10
            [void]$cmd.Parameters.AddWithValue("@id", $Agent)
            $names = @()
            $rd = $cmd.ExecuteReader()
            while ($rd.Read()) { $names += [string]$rd["user_name"] }
            $rd.Close()
            if ($names.Count -eq 1) { $name = $names[0] }
            else { Write-Warning "employee_id $Agent matched $($names.Count) users in cfg_person - not resolved" }
        } finally { $cn.Close() }
        $script:UserCache[$Agent] = @{ Name = $name; At = Get-Date }     # cache only successful lookups
    }
    catch { Write-Warning "DB lookup for '$Agent' failed: $($_.Exception.Message)" }
    return $name
}

function Get-Outlook {
    if (-not $script:Outlook) { $script:Outlook = New-Object -ComObject Outlook.Application }
    return $script:Outlook
}

# Manager's address from AD: find the agent by mail address, read "manager" (a DN), read that user's mail. $null if unavailable.
function Get-ManagerAddress([string]$AgentAddress, [string]$AgentName) {
    try {
        $a = $AgentAddress -replace '[\\*()\x00]', ''
        $n = $AgentName    -replace '[\\*()\x00]', ''
        # the agent can be matched by mail, proxyAddresses, UPN or user name (the AD "mail" can differ from the address book)
        $ids = "(mail=$a)(proxyAddresses=smtp:$a)(userPrincipalName=$a)"
        if ($n) { $ids += "(sAMAccountName=$n)" }
        $filter = "(&(objectCategory=person)(objectClass=user)(|$ids))"

        # local domain first, then the whole forest (global catalog) for agents in another domain
        $roots = @($null)
        try { $roots += "GC://" + [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Name } catch { Write-Verbose "Forest lookup failed: $($_.Exception.Message)" }

        $r = $null
        foreach ($root in $roots) {
            $ds = if ($root) { New-Object System.DirectoryServices.DirectorySearcher((New-Object System.DirectoryServices.DirectoryEntry($root))) } else { New-Object System.DirectoryServices.DirectorySearcher }
            $ds.Filter = $filter
            [void]$ds.PropertiesToLoad.Add("manager")
            $r = $ds.FindOne()
            if ($r) { break }
        }
        if (-not $r) { Write-Verbose "Agent $AgentAddress / $AgentName not found in AD (searched mail, proxyAddresses, UPN, user name; domain and forest)"; return $null }
        if ($r.Properties["manager"].Count -eq 0) { Write-Verbose "Agent found in AD ($($r.Path)) but its manager attribute is empty"; return $null }

        $mgrDn = [string]$r.Properties["manager"][0]
        $m = [ADSI]("LDAP://" + ($mgrDn -replace '/', '\/'))
        $mail = [string]$m.Properties["mail"].Value
        if ($mail) { return $mail }
        Write-Verbose "Manager of $AgentAddress has no mail attribute ($mgrDn)"
    } catch { Write-Verbose "Manager lookup for $AgentAddress failed: $($_.Exception.Message)" }
    return $null
}

# EWS ResolveNames = what Outlook's "Check Names" (Ctrl+K) does. Returns an SMTP address or $null.
# With several matches it takes the one whose name/address contains the agent; otherwise only a single match is used.
function Resolve-ViaEws([string]$Agent) {
    try {
        $esc = [System.Security.SecurityElement]::Escape($Agent)
        $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>
    <m:ResolveNames ReturnFullContactData="false" SearchScope="ActiveDirectoryContacts">
      <m:UnresolvedEntry>$esc</m:UnresolvedEntry>
    </m:ResolveNames>
  </soap:Body>
</soap:Envelope>
"@
        $r = Invoke-WebRequest -Uri $EwsUrl -Method Post -ContentType "text/xml; charset=utf-8" -Body $xml -UseDefaultCredentials -UseBasicParsing
        $found = @()
        foreach ($m in [regex]::Matches($r.Content, '<t:Mailbox>(.*?)</t:Mailbox>', 'Singleline')) {
            $addr = [regex]::Match($m.Groups[1].Value, '<t:EmailAddress>([^<]+)</t:EmailAddress>').Groups[1].Value
            $name = [regex]::Match($m.Groups[1].Value, '<t:Name>([^<]*)</t:Name>').Groups[1].Value
            $rt   = [regex]::Match($m.Groups[1].Value, '<t:RoutingType>([^<]*)</t:RoutingType>').Groups[1].Value
            if ($addr -and ($rt -eq '' -or $rt -eq 'SMTP')) { $found += [pscustomobject]@{ Addr = $addr; Name = $name } }
        }
        if ($found.Count -eq 1) { return $found[0].Addr }
        $hit = @($found | Where-Object { $_.Addr -like "*$Agent*" -or $_.Name -like "*$Agent*" })
        if ($hit.Count -eq 1) { return $hit[0].Addr }
    }
    catch { Write-Verbose "EWS name resolution for '$Agent' failed: $($_.Exception.Message)" }
    return $null
}

function Resolve-AgentAddress([string]$Agent) {
    if ($SmtpServer -or $UseEws) {      # no Outlook: look the agent up in Active Directory
        try {
            $ds = New-Object System.DirectoryServices.DirectorySearcher
            $safe = $Agent -replace '[\\*()\x00]', ''
            $ds.Filter = "(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=$safe)(employeeID=$safe)(employeeNumber=$safe)))"
            [void]$ds.PropertiesToLoad.Add("mail")
            $r = $ds.FindOne()
            if ($r -and $r.Properties["mail"].Count -gt 0) { return [string]$r.Properties["mail"][0] }
        } catch { Write-Verbose "AD lookup for '$Agent' failed: $($_.Exception.Message)" }
        if ($UseEws) {      # same name resolution as Ctrl+K in Outlook
            $x = Resolve-ViaEws $Agent
            if ($x) { return $x }
        }
        return $null
    }
    try {
        $r = (Get-Outlook).Session.CreateRecipient($Agent)
        if ($r.Resolve()) {
            $addr = $r.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            if ($addr) { return $addr }
        }
    } catch {}
    return $null
}

function Send-EventMail([hashtable]$f) {
    $isIscc = ($f.Type -eq 'ISCC')
    $to = @($TestTo)
    $cc = @()            # CC only when the mail really goes to its real recipients
    $note = ""
    if ($isIscc) {
        if ($SendToAgents -and $IsccTo.Count -gt 0) { $to = $IsccTo; $cc = $IsccCc }
        elseif ($SendToAgents) { $note = "<p dir=`"rtl`" style=`"color:#b00`">[לא הוגדרו נמענים להתראות ISCC - נשלח לכתובת הבדיקה]</p>" }
        else { $note = "<p dir=`"rtl`" style=`"color:#b00`">[מייל בדיקה - התראת ISCC]</p>" }
    }
    elseif ($SendToAgents) {
        $a = Resolve-AgentAddress $f.Agent
        if ($a) {
            $to = @($a); $cc = @($script:AgentCc)
            if ($CcManager) {
                $mgr = Get-ManagerAddress $a $f.Agent
                if ($mgr -and ($cc -notcontains $mgr) -and ($mgr -ne $a)) { $cc += $mgr }
            }
        }
        else { $note = "<p dir=`"rtl`" style=`"color:#b00`">[לא נמצאה כתובת מייל לסוכן '$([System.Net.WebUtility]::HtmlEncode($f.Agent))' - נשלח לכתובת הבדיקה]</p>" }
    }
    elseif ($f.Agent) {
        $who = [System.Net.WebUtility]::HtmlEncode($f.Agent)
        if ($f.AgentId -and $f.AgentId -ne $f.Agent) { $who += " (מספר עובד $([System.Net.WebUtility]::HtmlEncode($f.AgentId)))" }
        $note = "<p dir=`"rtl`" style=`"color:#b00`">[מייל בדיקה - סוכן: $who]</p>"
    }

    $dn = if ($f.DN) { [System.Net.WebUtility]::HtmlEncode($f.DN) } else { "לא זוהתה" }
    $subject = "השלוחה שלך ($dn) דורשת הגדרה מחדש"
    $link = if ($FormLink) { " – <a href=`"$([System.Net.WebUtility]::HtmlEncode($FormLink))`">קישור לטופס</a>" } else { "" }
    $from_html = [System.Net.WebUtility]::HtmlEncode($From)
    $body = @"
<div dir="rtl" style="font-family:Arial,sans-serif;font-size:14px;text-align:right">
<p>שלום,</p>
<p>השלוחה שאליה אתה מחובר/ת בסרגל הטלפוניה (<b>$dn</b>) אינה מוגדרת כראוי ויש להגדיר אותה מחדש.</p>
<p><b>מה צריך לעשות:</b></p>
<ol>
<li>לסגור את סרגל הטלפוניה.</li>
<li>לסגור את הג'אבר במידה ועובדים מהבית.</li>
<li>למלא טופס בית-משרד עבור השלוחה <b>$dn</b>$link</li>
<li>להפעיל מחדש את הג'אבר (במידה ועובדים מהבית) וסרגל הטלפוניה.</li>
</ol>
<p><b>במידה והשלוחה הינה ג'אבר בלבד:</b></p>
<ol>
<li>לסגור את סרגל הטלפוניה.</li>
<li>לסגור את הג'אבר.</li>
<li>להפעיל מחדש את הג'אבר וסרגל הטלפוניה.</li>
</ol>
<p>במידה וקיימת תקלה בהפעלת הסרגל ניתן לפנות למשוב.</p>
<p>תודה,<br>צוות CTI</p>
$note
</div>
"@

    if ($isIscc) {
        $subject = "שלוחה תפוסה בהעברת שיחה בין-שרתים (ISCC) - $dn"
        $when = [System.Net.WebUtility]::HtmlEncode($f.DetectedAt)
        $conn = [System.Net.WebUtility]::HtmlEncode($f.ConnID)
        $body = @"
<div dir="rtl" style="font-family:Arial,sans-serif;font-size:14px;text-align:right">
<p>שלום,</p>
<p>זוהתה שגיאת "DN is busy" בהעברת שיחה בין-שרתים (ISCC, <b>transfer</b>) לשלוחה <b>$dn</b>.</p>
<ul>
<li>שלוחת יעד: <b>$dn</b></li>
<li>זמן זיהוי: $when</li>
<li>ConnID: $conn</li>
</ul>
<p>תודה,<br>תמיכת מוקד – $from_html</p>
$note
</div>
"@
    }

    if ($UseEws) {
        $e = { param($s) [System.Security.SecurityElement]::Escape($s) }
        $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems"/></m:SavedItemFolderId>
      <m:Items>
        <t:Message>
          <t:Subject>$(& $e $subject)</t:Subject>
          <t:Body BodyType="HTML">$(& $e $body)</t:Body>
          <t:ToRecipients>$(($to | ForEach-Object { "<t:Mailbox><t:EmailAddress>$(& $e $_)</t:EmailAddress></t:Mailbox>" }) -join '')</t:ToRecipients>$(if ($cc.Count -gt 0) { "<t:CcRecipients>" + (($cc | ForEach-Object { "<t:Mailbox><t:EmailAddress>$(& $e $_)</t:EmailAddress></t:Mailbox>" }) -join '') + "</t:CcRecipients>" })
          <t:From><t:Mailbox><t:EmailAddress>$(& $e $From)</t:EmailAddress></t:Mailbox></t:From>
        </t:Message>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>
"@
        $r = Invoke-WebRequest -Uri $EwsUrl -Method Post -ContentType "text/xml; charset=utf-8" -Body $xml -UseDefaultCredentials -UseBasicParsing
        if ($r.Content -notmatch 'ResponseClass="Success"') { throw "EWS did not report success: $($r.Content.Substring(0, [Math]::Min(300, $r.Content.Length)))" }
    }
    elseif ($SmtpServer) {
        $msg = New-Object System.Net.Mail.MailMessage
        $msg.From = New-Object System.Net.Mail.MailAddress($From)
        foreach ($t in $to) { $msg.To.Add($t) }
        $msg.Subject = $subject
        $msg.Body = $body
        $msg.IsBodyHtml = $true
        $msg.SubjectEncoding = [System.Text.Encoding]::UTF8
        $msg.BodyEncoding    = [System.Text.Encoding]::UTF8
        foreach ($c in $cc) { $msg.CC.Add($c) }
        $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
        $smtp.EnableSsl = [bool]$UseSsl
        $smtp.UseDefaultCredentials = [bool]$UseCurrentUser
        try { $smtp.Send($msg) } finally { $msg.Dispose(); $smtp.Dispose() }
    }
    else {
        $m = (Get-Outlook).CreateItem(0)
        $m.To = ($to -join ';')
        if ($cc.Count -gt 0) { $m.CC = ($cc -join ';') }
        $m.SentOnBehalfOfName = $From
        $m.Subject = $subject
        $m.HTMLBody = $body
        $m.Send()
    }
    $kind = if ($isIscc) { "ISCC dn $($f.DN)" } else { "agent $($f.Agent)" }
    Write-Host ("{0} mail -> {1} cc {2} ({3}, ConnID {4})" -f (Get-Date -Format "HH:mm:ss"), ($to -join ';'), ($cc -join ';'), $kind, $f.ConnID)
}

$script:Hits  = @{}       # counting key -> times the error was seen (in memory; lost on restart)
$script:Burst = @{}       # DN key -> datetime of first detection in current burst

# Records one occurrence; returns $true once the key has been seen -MinOccurrences times within -WithinMinutes.
function Test-Threshold([string]$Key, [datetime]$Now) {
    if ($MinOccurrences -le 1) { return $true }
    $cut  = $Now.AddMinutes(-$WithinMinutes)
    $list = @($script:Hits[$Key] | Where-Object { $_ -ge $cut }) + $Now
    $script:Hits[$Key] = $list
    return ($list.Count -ge $MinOccurrences)
}

function Handle-Line([string]$Line) {
    if ($Line -notmatch '^DNBUSY\|') { return }
    $f = @{}
    foreach ($part in ($Line -split '\|')) {
        $i = $part.IndexOf('=')
        if ($i -gt 0) { $f[$part.Substring(0, $i)] = $part.Substring($i + 1) }
    }
    $f.AgentId = $f.Agent
    if ($f.Type -eq 'ISCC') { $key = "iscc:$($f.DN)" }      # no agent: dedupe per busy destination DN
    else {
        if ($f.Agent) { $f.Agent = Resolve-UserName $f.Agent }
        $key = if ($f.Agent) { $f.Agent.ToLower() } else { "conn:$($f.ConnID)" }
    }
    $now = Get-Date
    if ($BurstSeconds -gt 0) {
        $burstKey = if ($f.Type -eq 'ISCC') { "iscc:$($f.DN)" } else { $f.DN }
        if ($script:Burst.ContainsKey($burstKey)) {
            if (($now - $script:Burst[$burstKey]).TotalSeconds -lt $BurstSeconds) { return }
            else { $script:Burst.Remove($burstKey) }
        }
        $script:Burst[$burstKey] = $now
    }
    if ($script:Sent.ContainsKey($key)) {
        if ($DedupeMinutes -le 0 -or ($now - $script:Sent[$key]).TotalMinutes -lt $DedupeMinutes) { return }
    }
    $countKey = if ($f.DN) { "$($f.Type):$($f.DN)" } else { $key }      # count per DN (agent and ISCC events counted separately)
    if (-not (Test-Threshold $countKey $now)) { return }
    $script:Hits.Remove($countKey)
    Send-EventMail $f
    $script:Sent[$key] = $now
    Save-State $key $now
}

Write-Host "Watching $($SharePath -join ', ') (${Prefix}_yyyyMMdd.log) ; mails go to $(if ($SendToAgents) { 'agents (fallback ' + $TestTo + ')' } else { $TestTo }). Ctrl+C to stop."

# One read position per share. All shares use the same "already mailed" state, so an agent/DN seen
# by both servers is mailed once.
$script:Positions = @{}
$script:Day       = $null

function Read-ShareFile([string]$Share, [string]$FileName) {
    $path = Join-Path $Share $FileName
    if (-not (Test-Path -LiteralPath $path)) { return }
    $pos = [long]$script:Positions[$Share]
    $fs = [System.IO.File]::Open($path, 'Open', 'Read', [System.IO.FileShare]'ReadWrite, Delete')
    try {
        if ($fs.Length -lt $pos) { $pos = 0 }
        if ($fs.Length -gt $pos) {
            $fs.Seek($pos, 'Begin') | Out-Null
            $buf = New-Object byte[] ($fs.Length - $pos)
            $n = $fs.Read($buf, 0, $buf.Length)
            $cut = [Array]::LastIndexOf($buf, [byte]10, $n - 1)      # only complete lines
            if ($cut -ge 0) {
                $pos += $cut + 1
                $script:Positions[$Share] = $pos
                $text = [System.Text.Encoding]::UTF8.GetString($buf, 0, $cut + 1)
                foreach ($l in ($text -split "`r?`n")) { if ($l) { Handle-Line $l } }
            }
        }
    } finally { $fs.Close() }
}

while ($true) {
    try {
        $dayStr   = (Get-Date).ToString("yyyyMMdd")
        $fileName = "{0}_{1}.log" -f $Prefix, $dayStr
        if ($script:Day -ne $dayStr) {
            $first = ($null -eq $script:Day)
            $script:Day   = $dayStr
            $script:Hits  = @{}
            $script:Burst = @{}
            $script:Sent = Load-State          # new day (or first start): who was already mailed today
            foreach ($sh in $SharePath) {
                $p = Join-Path $sh $fileName
                # first start: skip history; new day: read the new file from the top
                $script:Positions[$sh] = if ($first -and (Test-Path -LiteralPath $p)) { (Get-Item -LiteralPath $p).Length } else { 0 }
            }
        }
        foreach ($sh in $SharePath) {
            try { Read-ShareFile $sh $fileName }
            catch { Write-Warning "Share '$sh': $($_.Exception.Message) [line $($_.InvocationInfo.ScriptLineNumber)]" }     # one unreachable share must not block the other
        }
    }
    catch {
        Write-Warning "Poll error: $($_.Exception.Message) [line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())]"
        $script:Outlook = $null     # recreate the Outlook object next time
    }
    Start-Sleep -Seconds $PollSeconds
}
