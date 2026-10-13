<#
.SYNOPSIS
    Tails the newest Genesys URS log file in a folder and emits one structured
    line per "EventError ... DN is busy" occurrence, for SCOM Log File
    monitoring to pick up.

.DESCRIPTION
    URS logs a multi-line block for each EventError. This script:
      1. Finds the most recently written log file matching -FilePattern in
         -LogFolder (handles log rotation by re-checking each poll cycle).
      2. Tails new content appended to that file.
      3. Buffers recent lines; when it sees AttributeErrorMessage 'DN is busy',
         it pulls context from the lines just before and after to extract the
         relevant attributes (DN, ThisDN, OtherDN, Agent, VQ, Switch, etc).
      4. Appends one single-line record per event to -OutputFile.

    Point a SCOM Log File monitor at -OutputFile with a pattern like
    "^DNBUSY\|" to alert on new lines.

.PARAMETER LogFolder
    Folder containing URS log files (rotating / dated files).

.PARAMETER FilePattern
    Filename filter used to find URS log files in -LogFolder. Default "urs*.log".

.PARAMETER OutputFile
    Single-line output file that SCOM will monitor. Created if missing.

.PARAMETER PollIntervalSeconds
    How often to check for new log content / a rotated file. Default 5.

.PARAMETER FromBeginning
    If set, when the script first attaches to a file it reads from the start
    instead of from the current end. Useful for testing against a static
    sample log; leave off in production so you don't replay old history.

.EXAMPLE
    .\Watch-DnBusyEvents.ps1 -LogFolder "D:\GCTI\logs\urs" -OutputFile "D:\Monitoring\dnbusy.log"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$LogFolder,

    [string]$FilePattern = "urs*.log",

    [Parameter(Mandatory = $true)]
    [string]$OutputFile,

    [int]$PollIntervalSeconds = 5,

    [switch]$FromBeginning
)

$ErrorActionPreference = "Stop"

# ---- state ----------------------------------------------------------------
$script:ActiveFile     = $null
$script:HasAttachedOnce = $false
$script:Position       = 0
$script:LineBuffer     = New-Object System.Collections.Generic.List[string]   # rolling context (pre-trigger)
$script:BufferMaxLines = 60

$script:Capturing      = $false
$script:CaptureLines   = New-Object System.Collections.Generic.List[string]
$script:CaptureBudget  = 0
$script:PreContext     = @()
$script:IdlePolls      = 0

$CaptureForwardLines = 25   # how many lines after the trigger to keep scanning for attributes
$MaxIdlePollsWhileCapturing = 2   # flush a stalled capture after this many polls with no new lines

function Ensure-OutputFile {
    $dir = Split-Path -Path $OutputFile -Parent
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    if (-not (Test-Path $OutputFile)) {
        New-Item -ItemType File -Path $OutputFile -Force | Out-Null
    }
}

function Get-LatestLogFile {
    Get-ChildItem -Path $LogFolder -Filter $FilePattern -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Extract-Field {
    param([string]$Text, [string]$Pattern)
    $m = [regex]::Match($Text, $Pattern)
    if ($m.Success) { return $m.Groups[1].Value }
    return ""
}

function Emit-DnBusyRecord {
    param([string[]]$Block)

    $text = ($Block -join "`n")

    $dn        = Extract-Field $text "'DN'\s+'(\d+)'"
    $place     = Extract-Field $text "'PLACE'\s+'(\d+)'"
    $agent     = Extract-Field $text "'AGENT'\s+'([^']+)'"
    $vq        = Extract-Field $text "'VQ'\s+'([^']+)'"
    $connId    = Extract-Field $text "AttributeConnID\s+(\S+)"

    $nowStamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    $record = "DNBUSY|DetectedAt=$nowStamp|DN=$dn|Place=$place|Agent=$agent|VQ=$vq|ConnID=$connId"

    Add-Content -Path $OutputFile -Value $record
    Write-Verbose $record
}

function Process-Line {
    param([string]$Line)

    if ($script:Capturing) {
        if ($Line -match "AttributeErrorMessage\s+'DN is busy'") {
            # A new event started before the previous one's forward-capture window
            # finished (e.g. back-to-back bursts). Flush what we have so far for
            # the in-progress event, then start capturing the new one fresh.
            $full = $script:PreContext + $script:CaptureLines.ToArray()
            Emit-DnBusyRecord -Block $full

            $script:PreContext    = $script:LineBuffer.ToArray()
            $script:CaptureLines  = New-Object System.Collections.Generic.List[string]
            $script:CaptureLines.Add($Line) | Out-Null
            $script:CaptureBudget = $CaptureForwardLines
            $script:IdlePolls = 0

            $script:LineBuffer.Add($Line) | Out-Null
            if ($script:LineBuffer.Count -gt $script:BufferMaxLines) {
                $script:LineBuffer.RemoveAt(0)
            }
            return
        }

        $script:CaptureLines.Add($Line) | Out-Null
        $script:CaptureBudget--
        if ($script:CaptureBudget -le 0) {
            $full = $script:PreContext + $script:CaptureLines.ToArray()
            Emit-DnBusyRecord -Block $full
            $script:Capturing = $false
            $script:CaptureLines.Clear()
            $script:PreContext = @()
        }
        return
    }

    if ($Line -match "AttributeErrorMessage\s+'DN is busy'") {
        $script:PreContext    = $script:LineBuffer.ToArray()
        $script:Capturing     = $true
        $script:CaptureLines  = New-Object System.Collections.Generic.List[string]
        $script:CaptureLines.Add($Line) | Out-Null
        $script:CaptureBudget = $CaptureForwardLines
        $script:IdlePolls = 0
    }

    $script:LineBuffer.Add($Line) | Out-Null
    if ($script:LineBuffer.Count -gt $script:BufferMaxLines) {
        $script:LineBuffer.RemoveAt(0)
    }
}

function Attach-ToFile {
    param([System.IO.FileInfo]$File)

    Write-Verbose "Attaching to log file: $($File.FullName)"
    $script:ActiveFile = $File.FullName
    $script:LineBuffer.Clear()
    $script:Capturing = $false
    $script:CaptureLines.Clear()
    $script:PreContext = @()

    if ($FromBeginning -or $script:HasAttachedOnce) {
        # Every file after the first is one we've never seen before (reached via
        # rotation) - read it from the start so we don't skip a burst of lines
        # URS wrote before our next poll noticed the new file existed.
        $script:Position = 0
    } else {
        # First attach at script startup: skip existing backlog, start at the
        # current end of whatever file happens to be newest right now.
        $script:Position = $File.Length
    }
    $script:HasAttachedOnce = $true
}

function Read-NewContent {
    if (-not (Test-Path $script:ActiveFile)) { return }

    $fs = [System.IO.File]::Open($script:ActiveFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $length = $fs.Length
        if ($length -lt $script:Position) {
            # file was truncated/replaced; restart from beginning
            $script:Position = 0
        }
        if ($length -gt $script:Position) {
            $fs.Seek($script:Position, [System.IO.SeekOrigin]::Begin) | Out-Null
            $bytesToRead = $length - $script:Position
            $buffer = New-Object byte[] $bytesToRead
            $fs.Read($buffer, 0, $bytesToRead) | Out-Null
            $script:Position = $length

            $text = [System.Text.Encoding]::UTF8.GetString($buffer)
            $lines = $text -split "`r?`n"
            foreach ($line in $lines) {
                if ($line -ne "") {
                    Process-Line -Line $line
                }
            }
            $script:IdlePolls = 0
        }
        elseif ($script:Capturing) {
            # No new data this poll while a block is mid-capture (e.g. file end reached
            # before the forward-capture budget was used up). Flush after a short grace
            # period so we don't lose events sitting at the tail of the file.
            $script:IdlePolls++
            if ($script:IdlePolls -ge $MaxIdlePollsWhileCapturing) {
                $full = $script:PreContext + $script:CaptureLines.ToArray()
                Emit-DnBusyRecord -Block $full
                $script:Capturing = $false
                $script:CaptureLines.Clear()
                $script:PreContext = @()
                $script:IdlePolls = 0
            }
        }
    }
    finally {
        $fs.Close()
    }
}

# ---- main -------------------------------------------------------------
Ensure-OutputFile
Write-Host "Watching '$LogFolder\$FilePattern' -> '$OutputFile' (poll every ${PollIntervalSeconds}s). Ctrl+C to stop."

while ($true) {
    try {
        $latest = Get-LatestLogFile
        if ($null -ne $latest) {
            if ($script:ActiveFile -ne $latest.FullName) {
                Attach-ToFile -File $latest
            }
            Read-NewContent
        }
    }
    catch {
        Write-Warning "Error during poll cycle: $($_.Exception.Message)"
    }
    Start-Sleep -Seconds $PollIntervalSeconds
}
