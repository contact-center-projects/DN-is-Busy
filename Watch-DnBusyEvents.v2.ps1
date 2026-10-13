<#
.SYNOPSIS
    Reads Genesys URS rotating log files in order and emits one structured line
    per "AttributeErrorMessage 'DN is busy'" event, for SCOM Log File monitoring.

.DESCRIPTION
    v2 changes (URS rotates a ~51MB file every ~25 seconds):
      - Files are processed in NAME order (the name contains a sortable timestamp),
        never by LastWriteTime, so the 0-byte new file / full old file tie at
        rotation cannot make the script flip back and forth and replay files.
      - The old file is always read to its end before moving to the next file,
        so lines written between polls and rotation are not lost.
      - Reads in 4MB chunks cut at a newline (no partial lines, no split UTF-8),
        and finds events with .NET regex on the chunk instead of per-line
        PowerShell, so it keeps up with ~2MB/s of log data.
      - Fields (DN/Place/Agent/VQ/ConnID) are taken from the occurrence NEAREST
        the trigger line, not the first one in a 60-line window.
      - Output keeps the same prefix/fields as v1 and appends LogFile=<name>.

.PARAMETER LogFolder
    Folder containing the URS log files.
.PARAMETER FilePattern
    Filename filter. Default "UniversalRoutingServer*.log".
.PARAMETER OutputFolder
    Folder for the daily output files <OutputPrefix>_yyyyMMdd.log (one record per line).
    Point SCOM at "<OutputFolder>\dnbusy_*.log" with pattern "^DNBUSY\|".
.PARAMETER OutputPrefix
    Output file name prefix. Default "dnbusy".
.PARAMETER RetentionDays
    Daily files older than this many days are deleted (checked at start and at midnight). Default 7.
.PARAMETER PollIntervalSeconds
    Poll interval. Default 5.
.PARAMETER FromBeginning
    On first attach, read the newest file from its start instead of its end.

.EXAMPLE
    .\Watch-DnBusyEvents.v2.ps1 -LogFolder "D:\Genesys_Logs\UniversalRoutingServer_Backup" `
        -FilePattern "UniversalRoutingServer_Backup*.log" -OutputFolder "D:\Monitoring\DnBusy" -Verbose
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$LogFolder,

    [string]$FilePattern = "UniversalRoutingServer*.log",

    [Parameter(Mandatory = $true)]
    [string]$OutputFolder,

    [string]$OutputPrefix = "dnbusy",

    [int]$RetentionDays = 7,

    [int]$PollIntervalSeconds = 5,

    [switch]$FromBeginning
)

$ErrorActionPreference = "Stop"

$ChunkBytes   = 4MB
$BeforeChars  = 8000    # context kept before a trigger
$AfterChars   = 6000    # context needed after a trigger
$MaxIdlePolls = 2       # flush a trailing event after this many polls with no new data

$RxOpt     = [System.Text.RegularExpressions.RegexOptions]::Compiled
$TriggerRx = New-Object System.Text.RegularExpressions.Regex("AttributeErrorMessage\s+'DN is busy'", $RxOpt)
$DnRx      = New-Object System.Text.RegularExpressions.Regex("'DN'\s+'(\d+)'", $RxOpt)
$PlaceRx   = New-Object System.Text.RegularExpressions.Regex("'PLACE'\s+'(\d+)'", $RxOpt)
$AgentRx   = New-Object System.Text.RegularExpressions.Regex("'AGENT'\s+'([^']+)'", $RxOpt)
$VqRx      = New-Object System.Text.RegularExpressions.Regex("'VQ'\s+'([^']+)'", $RxOpt)
$ConnRx    = New-Object System.Text.RegularExpressions.Regex("AttributeConnID\s+(\S+)", $RxOpt)
$IsccRx    = New-Object System.Text.RegularExpressions.Regex("'iscc-destination-dn'\s+'(\d+)'", $RxOpt)
$LineTsRx  = New-Object System.Text.RegularExpressions.Regex("^\d{1,2}:\d{2}:\d{2}\.\d+", ($RxOpt -bor [System.Text.RegularExpressions.RegexOptions]::Multiline))

$script:CurrentName = $null
$script:Position    = 0
$script:Carry       = ""
$script:IdlePolls   = 0
$script:Buffer      = New-Object byte[] $ChunkBytes

function Get-DailyOutputPath {
    return Join-Path $OutputFolder ("{0}_{1}.log" -f $OutputPrefix, (Get-Date).ToString("yyyyMMdd"))
}

# Deletes <prefix>_yyyyMMdd.log files whose date (taken from the name) is older than $RetentionDays.
function Remove-OldOutputFiles {
    $cutoff = (Get-Date).Date.AddDays(-$RetentionDays)
    foreach ($f in Get-ChildItem -Path $OutputFolder -Filter "${OutputPrefix}_????????.log" -File -ErrorAction SilentlyContinue) {
        $m = [regex]::Match($f.Name, "^$([regex]::Escape($OutputPrefix))_(\d{8})\.log$")
        if (-not $m.Success) { continue }
        $d = [datetime]::MinValue
        if ([datetime]::TryParseExact($m.Groups[1].Value, "yyyyMMdd", $null, [System.Globalization.DateTimeStyles]::None, [ref]$d) -and $d -lt $cutoff) {
            try { Remove-Item -LiteralPath $f.FullName -Force; Write-Verbose "Deleted old output file $($f.Name)" }
            catch { Write-Warning "Could not delete $($f.FullName): $($_.Exception.Message)" }
        }
    }
}

function Ensure-OutputFile {
    if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    $path = Get-DailyOutputPath
    if (-not (Test-Path $path)) { New-Item -ItemType File -Path $path -Force | Out-Null }
    $script:CleanupDay = (Get-Date).Date
    Remove-OldOutputFiles
}

function Get-NearestValue {
    param([string]$Text, [System.Text.RegularExpressions.Regex]$Rx, [int]$Anchor)
    $best = ""
    $bestDist = [int]::MaxValue
    foreach ($m in $Rx.Matches($Text)) {
        $d = [Math]::Abs($m.Index - $Anchor)
        if ($d -lt $bestDist) { $bestDist = $d; $best = $m.Groups[1].Value }
    }
    return $best
}

function Emit-DnBusyRecord {
    param([string]$Window, [int]$Anchor, [string]$FileName)

    $connId   = Get-NearestValue $Window $ConnRx $Anchor
    $nowStamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    # The T-Server EventError block (from its timestamped header line up to the trigger) is either
    #   - an AGENT event: has the 'AGENT' extension (it also carries iscc-destination-dn, so that
    #     alone does not make it ISCC), or
    #   - an ISCC event: no 'AGENT' extension, only 'iscc-destination-dn' = the DN that was busy.
    $before     = $Window.Substring(0, $Anchor)
    $tsMatches  = $LineTsRx.Matches($before)
    $blockStart = if ($tsMatches.Count -gt 0) { $tsMatches[$tsMatches.Count - 1].Index } else { 0 }
    $block      = $before.Substring($blockStart)
    $isccMatch  = $IsccRx.Matches($block)
    if (-not $AgentRx.IsMatch($block) -and $isccMatch.Count -gt 0) {
        $destDn = $isccMatch[$isccMatch.Count - 1].Groups[1].Value
        $record = "DNBUSY|DetectedAt=$nowStamp|Type=ISCC|DN=$destDn|Reason=DN is busy found in iscc-destination-dn|ConnID=$connId|LogFile=$FileName"
    }
    else {
        $dn     = Get-NearestValue $Window $DnRx $Anchor
        $place  = Get-NearestValue $Window $PlaceRx $Anchor
        $agent  = Get-NearestValue $Window $AgentRx $Anchor
        $vq     = Get-NearestValue $Window $VqRx $Anchor
        $record = "DNBUSY|DetectedAt=$nowStamp|DN=$dn|Place=$place|Agent=$agent|VQ=$vq|ConnID=$connId|LogFile=$FileName"
    }

    Add-Content -Path (Get-DailyOutputPath) -Value $record
    Write-Verbose $record
}

# Emits every complete event in $Text. Returns the text to carry into the next scan:
# the last $BeforeChars (pre-context) or, if an event still lacks its after-context,
# everything from that event's pre-context on. Emitted events are never re-scanned.
function Scan-Text {
    param([string]$Text, [bool]$Final, [string]$FileName)

    $lastEnd   = 0
    $carryFrom = -1

    foreach ($m in $TriggerRx.Matches($Text)) {
        if (-not $Final -and ($m.Index + $AfterChars) -gt $Text.Length) {
            $carryFrom = [Math]::Max([Math]::Max(0, $m.Index - $BeforeChars), $lastEnd)
            break
        }
        $ws = [Math]::Max(0, $m.Index - $BeforeChars)
        $we = [Math]::Min($Text.Length, $m.Index + $AfterChars)
        Emit-DnBusyRecord -Window $Text.Substring($ws, $we - $ws) -Anchor ($m.Index - $ws) -FileName $FileName
        $lastEnd = $m.Index + $m.Length
    }

    if ($carryFrom -lt 0) {
        $carryFrom = [Math]::Max([Math]::Max(0, $Text.Length - $BeforeChars), $lastEnd)
    }
    return $Text.Substring($carryFrom)
}

# Reads everything currently available in the current file. Returns $true if data was read.
function Read-CurrentFile {
    $path = Join-Path $LogFolder $script:CurrentName
    if (-not (Test-Path -LiteralPath $path)) { return $false }

    $read = $false
    $share = [System.IO.FileShare]"ReadWrite, Delete"
    $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    try {
        if ($fs.Length -lt $script:Position) { $script:Position = 0 }   # truncated/replaced
        while ($fs.Length -gt $script:Position) {
            $fs.Seek($script:Position, [System.IO.SeekOrigin]::Begin) | Out-Null
            $n = $fs.Read($script:Buffer, 0, $script:Buffer.Length)
            if ($n -le 0) { break }

            $cut = [Array]::LastIndexOf($script:Buffer, [byte]10, $n - 1)
            if ($cut -lt 0) {
                if ($n -lt $script:Buffer.Length) { break }   # partial line, wait for the rest
                $cut = $n - 1                                  # one line longer than the buffer
            }
            $count = $cut + 1
            $text  = [System.Text.Encoding]::UTF8.GetString($script:Buffer, 0, $count)
            $script:Position += $count
            $script:Carry = Scan-Text -Text ($script:Carry + $text) -Final $false -FileName $script:CurrentName
            $read = $true
        }
    }
    finally {
        $fs.Close()
    }
    return $read
}

# ---- main -------------------------------------------------------------------
Ensure-OutputFile
Write-Host "Watching '$LogFolder\$FilePattern' -> '$OutputFolder\${OutputPrefix}_yyyyMMdd.log' (poll every ${PollIntervalSeconds}s, keep $RetentionDays days). Ctrl+C to stop."

while ($true) {
    try {
        if ((Get-Date).Date -ne $script:CleanupDay) {   # new day: create today's file, purge old ones
            Ensure-OutputFile
        }

        $files = @(Get-ChildItem -Path $LogFolder -Filter $FilePattern -File -ErrorAction SilentlyContinue | Sort-Object Name)

        if ($files.Count -gt 0) {
            if (-not $script:CurrentName) {
                $newest = $files[-1]
                $script:CurrentName = $newest.Name
                $script:Position    = if ($FromBeginning) { 0 } else { $newest.Length }
                Write-Verbose "Attached to $($newest.Name) at position $($script:Position)"
            }
            elseif (-not (Test-Path -LiteralPath (Join-Path $LogFolder $script:CurrentName))) {
                # current file was deleted (retention) - continue with the next one
                $nextFile = $files | Where-Object { [string]::CompareOrdinal($_.Name, $script:CurrentName) -gt 0 } | Select-Object -First 1
                if ($nextFile) { $script:CurrentName = $nextFile.Name; $script:Position = 0 }
            }

            $progress = $false
            while ($true) {
                # Decide "is there a newer file" BEFORE reading, so the read below sees
                # the complete old file once URS has moved on to the new one.
                $nextFile = $files | Where-Object { [string]::CompareOrdinal($_.Name, $script:CurrentName) -gt 0 } | Select-Object -First 1
                if (Read-CurrentFile) { $progress = $true }
                if ($null -eq $nextFile) { break }
                Write-Verbose "Rotated: $($script:CurrentName) -> $($nextFile.Name)"
                $script:CurrentName = $nextFile.Name
                $script:Position    = 0
            }

            if ($progress) {
                $script:IdlePolls = 0
            }
            else {
                $script:IdlePolls++
                if ($script:IdlePolls -ge $MaxIdlePolls -and $script:Carry.Length -gt 0) {
                    $script:Carry = Scan-Text -Text $script:Carry -Final $true -FileName $script:CurrentName
                }
            }
        }
    }
    catch {
        Write-Warning "Error during poll cycle: $($_.Exception.Message)"
    }
    Start-Sleep -Seconds $PollIntervalSeconds
}
