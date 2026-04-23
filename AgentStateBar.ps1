param(
    [switch]$SelfTest
)

$ErrorActionPreference = 'SilentlyContinue'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class AgentStateNative {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int maxCount);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

    public const int SW_RESTORE = 9;
    public const int SW_SHOW = 5;
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const uint WM_CLOSE = 0x0010;
    public const uint GW_OWNER = 4;
    public const uint GA_ROOT = 2;
    public const uint GA_ROOTOWNER = 3;
}
"@

$script:CodexHome = Join-Path $env:USERPROFILE '.codex'
$script:CodexSessionRoot = Join-Path $script:CodexHome 'sessions'
$script:CodexHistoryPath = Join-Path $script:CodexHome 'history.jsonl'
$script:CodexLogPath = Join-Path $script:CodexHome 'log\codex-tui.log'
$script:ClaudeHome = Join-Path $env:USERPROFILE '.claude'
$script:ClaudeSessionRoot = Join-Path $script:ClaudeHome 'sessions'
$script:ClaudeProjectRoot = Join-Path $script:ClaudeHome 'projects'
$script:ClaudeTelemetryRoot = Join-Path $script:ClaudeHome 'telemetry'
$script:ClaudeHistoryPath = Join-Path $script:ClaudeHome 'history.jsonl'
$script:LegacyStateHome = Join-Path $env:USERPROFILE '.codex-state'
$script:StateHome = Join-Path $env:USERPROFILE '.agent-state'
$script:LaunchRoot = Join-Path $script:StateHome 'launches'
$script:SettingsPath = Join-Path $script:StateHome 'config.json'
$script:CodexLauncherPath = Join-Path $PSScriptRoot 'StartCodexTracked.bat'
$script:ClaudeLauncherPath = Join-Path $PSScriptRoot 'StartClaudeTracked.bat'

$script:ExpandedWidth = 390
$script:ExpandedHeight = 520
$script:CollapsedThickness = 34
$script:ExpandedMinLength = 250
$script:ExpandedMaxLength = 620
$script:CollapsedMinLength = 64
$script:CollapsedMaxLength = 260
$script:IsExpanded = $false
$script:IsDragging = $false
$script:IsConfigDialogOpen = $false
$script:IsCollapsedPointerDown = $false
$script:IsCollapsedDragging = $false
$script:CollapsedPressPoint = $null
$script:AnchorSide = 'Right'
$script:AnchorOffset = $null
$script:ConfigOpen = $false
$script:Settings = [ordered]@{
    ExpandMode = 'Hover'
    AnchorSide = 'Right'
    AnchorOffset = $null
}
$script:HistoryNames = @{}
$script:ClaudeHistoryNames = @{}
$script:CodexSessions = @()
$script:ClaudeSessions = @()
$script:ClaudeSessionByProcessId = @{}
$script:ClaudeTelemetryFilesBySession = @{}
$script:CodexThreadStates = @{}
$script:CodexSessionOutcomeCache = @{}
$script:ClaudeSessionStateCache = @{}
$script:ClaudeTelemetryFileStateCache = @{}
$script:CodexLogPosition = $null
$script:InitialCodexLogReadBytes = 2097152
$script:LastHistoryLoad = [datetime]::MinValue
$script:LastSessionLoad = [datetime]::MinValue
$script:LastClaudeSessionLoad = [datetime]::MinValue
$script:LastClaudeTelemetryIndexLoad = [datetime]::MinValue
$script:LastInstances = @()
$script:CurrentCollapsedLength = $script:CollapsedMinLength
$script:CurrentExpandedLength = $script:ExpandedMinLength
$script:LaunchMatchWindowSeconds = 180
$script:LaunchStartupGraceSeconds = 45
$script:LaunchProcessFutureSkewSeconds = 12
$script:LaunchIdentitySlackSeconds = 15
$script:SessionStartMatchSeconds = 180
$script:SessionActivityLeadSeconds = 15
$script:SessionFutureStartGraceSeconds = 90
$script:SessionStartPenaltySeconds = 600
$script:ClaudeSessionLoadIntervalSeconds = 4
$script:ClaudeTelemetryLoadIntervalSeconds = 2

function Initialize-StateHome {
    if (-not (Test-Path $script:StateHome)) {
        New-Item -ItemType Directory -Path $script:StateHome -Force | Out-Null
    }
    if (-not (Test-Path $script:LaunchRoot)) {
        New-Item -ItemType Directory -Path $script:LaunchRoot -Force | Out-Null
    }

    $legacyLaunchRoot = Join-Path $script:LegacyStateHome 'launches'
    if (Test-Path $legacyLaunchRoot) {
        Get-ChildItem -File -Path $legacyLaunchRoot -Filter '*.json' | ForEach-Object {
            $target = Join-Path $script:LaunchRoot $_.Name
            if (-not (Test-Path $target)) {
                try { Copy-Item -LiteralPath $_.FullName -Destination $target -Force } catch { }
            }
        }
    }
}

function Load-Settings {
    if (Test-Path $script:SettingsPath) {
        try {
            $loaded = Get-Content -Raw -Path $script:SettingsPath -Encoding UTF8 | ConvertFrom-Json
            if ($loaded.ExpandMode -in @('Hover', 'Click')) {
                $script:Settings.ExpandMode = [string]$loaded.ExpandMode
            }
            if ($loaded.AnchorSide -in @('Left', 'Right', 'Top', 'Bottom')) {
                $script:Settings.AnchorSide = [string]$loaded.AnchorSide
            }
            if ($null -ne $loaded.AnchorOffset) {
                $script:Settings.AnchorOffset = [double]$loaded.AnchorOffset
            }
        } catch { }
    }
}

function Save-Settings {
    try {
        [pscustomobject]@{
            ExpandMode = $script:Settings.ExpandMode
            AnchorSide = $script:AnchorSide
            AnchorOffset = $script:AnchorOffset
        } | ConvertTo-Json -Depth 3 | Set-Content -Path $script:SettingsPath -Encoding UTF8
    } catch { }
}

Initialize-StateHome
Load-Settings

function Shorten-Text {
    param(
        [string]$Text,
        [int]$Max = 64
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $clean = ($Text -replace '\s+', ' ').Trim()
    if ($clean.Length -le $Max) { return $clean }
    return $clean.Substring(0, [Math]::Max(0, $Max - 3)) + '...'
}

function Get-WorkingDetail {
    param(
        $State,
        [string]$Fallback = 'active'
    )
    if ($State) {
        if ($State.HasRecentDisconnect) { return 'network retry' }
        if ($State.LastEvent -and $State.LastEvent -notin @('waiting for input', 'ready')) {
            return $State.LastEvent
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($Fallback)) { return $Fallback }
    return 'active'
}

function Convert-WmiTime {
    param($Value)
    if ($Value -is [datetime]) { return $Value }
    if ($Value -is [datetimeoffset]) { return $Value.LocalDateTime }
    try { return [Management.ManagementDateTimeConverter]::ToDateTime([string]$Value) } catch { return Get-Date }
}

function Convert-SerializedDateTime {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value }
    if ($Value -is [datetimeoffset]) { return $Value.LocalDateTime }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    if ($text -match '^/Date\((\d+)([+-]\d{4})?\)/$') {
        try { return [datetimeoffset]::FromUnixTimeMilliseconds([int64]$Matches[1]).LocalDateTime } catch { }
    }

    try { return ([datetimeoffset]::Parse($text)).LocalDateTime } catch { }
    try { return [datetime]$text } catch { }
    return $null
}

function Convert-UnixMillisecondsToDateTime {
    param($Value)
    if ($null -eq $Value) { return $null }
    try { return [datetimeoffset]::FromUnixTimeMilliseconds([int64]$Value).LocalDateTime } catch { return $null }
}

function Clamp-Double {
    param(
        [double]$Value,
        [double]$Min,
        [double]$Max
    )
    if ($Max -lt $Min) { return $Min }
    return [Math]::Min($Max, [Math]::Max($Min, $Value))
}

function Is-HoverExpandMode {
    return $script:Settings.ExpandMode -ne 'Click'
}

function Get-CollapsedLength {
    $count = [Math]::Max(1, $script:LastInstances.Count)
    $length = 28 + ($count * 18)
    return [int](Clamp-Double $length $script:CollapsedMinLength $script:CollapsedMaxLength)
}

function Get-ExpandedLength {
    $count = [Math]::Max(1, $script:LastInstances.Count)
    $length = 108 + ($count * 88)
    return [int](Clamp-Double $length $script:ExpandedMinLength $script:ExpandedMaxLength)
}

function Parse-CodexSessionFileName {
    param([System.IO.FileInfo]$File)
    if ($File.Name -notmatch '^rollout-(\d{4})-(\d{2})-(\d{2})T(\d{2})-(\d{2})-(\d{2})-([0-9a-f-]+)\.jsonl$') {
        return $null
    }

    $start = Get-Date -Year ([int]$Matches[1]) -Month ([int]$Matches[2]) -Day ([int]$Matches[3]) -Hour ([int]$Matches[4]) -Minute ([int]$Matches[5]) -Second ([int]$Matches[6])
    return [pscustomobject]@{
        SessionId = $Matches[7]
        StartTime = $start
        Path = $File.FullName
        LastWriteTime = $File.LastWriteTime
        Cwd = ''
    }
}

function Load-CodexSessions {
    if ((Get-Date) -lt $script:LastSessionLoad.AddSeconds(8)) { return }
    $script:LastSessionLoad = Get-Date
    if (-not (Test-Path $script:CodexSessionRoot)) {
        $script:CodexSessions = @()
        return
    }

    $items = @()
    Get-ChildItem -Recurse -File -Path $script:CodexSessionRoot -Filter '*.jsonl' | ForEach-Object {
        $session = Parse-CodexSessionFileName $_
        if ($null -eq $session) { return }
        if ($_.Length -gt 0) {
            try {
                $first = Get-Content -Path $_.FullName -TotalCount 1 -Encoding UTF8
                if ($first) {
                    $meta = $first | ConvertFrom-Json
                    if ($meta.type -eq 'session_meta' -and $meta.payload.cwd) {
                        $session.Cwd = [string]$meta.payload.cwd
                    }
                }
            } catch { }
        }
        $items += $session
    }

    $script:CodexSessions = @($items | Sort-Object StartTime -Descending)
}

function Convert-EventTimestamp {
    param($Record)
    if ($null -eq $Record) { return $null }

    if ($Record.timestamp) {
        try { return ([datetimeoffset]::Parse([string]$Record.timestamp)).LocalDateTime } catch { }
    }

    if ($Record.payload) {
        foreach ($name in @('started_at', 'completed_at')) {
            $value = $Record.payload.$name
            if ($null -ne $value -and "$value".Length -gt 0) {
                try { return [datetimeoffset]::FromUnixTimeSeconds([int64]$value).LocalDateTime } catch { }
            }
        }

        foreach ($name in @('timestamp', 'created_at')) {
            $value = $Record.payload.$name
            if ($null -ne $value -and "$value".Length -gt 0) {
                try { return ([datetimeoffset]::Parse([string]$value)).LocalDateTime } catch { }
            }
        }
    }

    return $null
}

function Convert-StructuredText {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }

    $props = $null
    try { $props = $Value.PSObject.Properties } catch { }
    if ($props) {
        foreach ($name in @('text', 'message', 'content')) {
            if ($props.Match($name).Count -gt 0) {
                $text = Convert-StructuredText $Value.$name
                if (-not [string]::IsNullOrWhiteSpace($text)) { return $text }
            }
        }
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $parts = foreach ($item in $Value) {
            $text = Convert-StructuredText $item
            if (-not [string]::IsNullOrWhiteSpace($text)) { $text }
        }
        return ($parts -join ' ').Trim()
    }

    try { return [string]$Value } catch { return '' }
}

function Read-FileTailLines {
    param(
        [string]$Path,
        [int]$MaxBytes = 1048576,
        [int]$MaxLines = 220
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path $Path)) { return @() }

    try {
        $file = Get-Item -LiteralPath $Path
        $start = [Math]::Max(0, $file.Length - $MaxBytes)
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $stream.Seek($start, [System.IO.SeekOrigin]::Begin) | Out-Null
            $bytesToRead = [int]($file.Length - $start)
            $buffer = New-Object byte[] $bytesToRead
            $read = $stream.Read($buffer, 0, $bytesToRead)
        } finally {
            $stream.Dispose()
        }

        if ($read -le 0) { return @() }
        $text = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $read)
        $lines = @($text -split "`r?`n")
        if ($start -gt 0 -and $lines.Count -gt 0) {
            $lines = @($lines | Select-Object -Skip 1)
        }
        if ($lines.Count -le $MaxLines) { return $lines }
        return @($lines | Select-Object -Last $MaxLines)
    } catch {
        return @()
    }
}

function Get-CodexSessionOutcome {
    param($Session)
    if ($null -eq $Session -or [string]::IsNullOrWhiteSpace([string]$Session.Path) -or -not (Test-Path $Session.Path)) {
        return $null
    }

    $cacheKey = [string]$Session.SessionId
    $file = $null
    try { $file = Get-Item -LiteralPath $Session.Path } catch { return $null }

    if ($script:CodexSessionOutcomeCache.ContainsKey($cacheKey)) {
        $cached = $script:CodexSessionOutcomeCache[$cacheKey]
        if (
            $cached.Path -eq $Session.Path -and
            $cached.FileLength -eq $file.Length -and
            $cached.FileLastWriteTimeUtc -eq $file.LastWriteTimeUtc
        ) {
            return $cached
        }
    }

    $currentTurnId = ''
    $lastTurnId = ''
    $lastTurnStartedTime = $null
    $lastTurnCompletedTime = $null
    $lastTurnCompletedNormally = $null
    $lastTurnSummary = ''
    $lastSessionErrorTime = $null
    $lastSessionErrorText = ''
    $lastOutcomeCode = 'Unknown'
    $lastOutcomeTime = $null
    $lastOutcomeDetail = ''

    try {
        $lines = Read-FileTailLines -Path $Session.Path -MaxBytes 1048576 -MaxLines 220
        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $lineHeader = if ($line.Length -gt 320) { $line.Substring(0, 320) } else { $line }
            if (
                $lineHeader.IndexOf('"type":"event_msg"') -lt 0 -or
                (
                    $lineHeader.IndexOf('"type":"task_started"') -lt 0 -and
                    $lineHeader.IndexOf('"type":"task_complete"') -lt 0 -and
                    $lineHeader.IndexOf('"type":"error"') -lt 0
                )
            ) {
                continue
            }

            $record = $null
            try { $record = $line | ConvertFrom-Json } catch { }
            if ($null -eq $record -or [string]$record.type -ne 'event_msg' -or $null -eq $record.payload) { continue }

            $payloadType = [string]$record.payload.type
            $eventTime = Convert-EventTimestamp $record
            if ($null -eq $eventTime) { $eventTime = $file.LastWriteTime }

            switch ($payloadType) {
                'task_started' {
                    $currentTurnId = if ($record.payload.turn_id) { [string]$record.payload.turn_id } else { $currentTurnId }
                    if ($currentTurnId) { $lastTurnId = $currentTurnId }
                    $lastTurnStartedTime = $eventTime
                    $lastTurnCompletedTime = $null
                    $lastTurnCompletedNormally = $null
                    $lastTurnSummary = ''
                    $lastSessionErrorTime = $null
                    $lastSessionErrorText = ''
                    $lastOutcomeCode = 'Started'
                    $lastOutcomeTime = $eventTime
                    $lastOutcomeDetail = 'active'
                }
                'error' {
                    $lastSessionErrorTime = $eventTime
                    $lastSessionErrorText = Shorten-Text (Convert-StructuredText $record.payload.message) 140
                    if ([string]::IsNullOrWhiteSpace($lastSessionErrorText) -and $record.payload.codex_error_info) {
                        $lastSessionErrorText = Shorten-Text ([string]$record.payload.codex_error_info) 140
                    }
                    $lastOutcomeCode = 'Error'
                    $lastOutcomeTime = $eventTime
                    $lastOutcomeDetail = if ($lastSessionErrorText) { $lastSessionErrorText } else { 'session error' }
                }
                'task_complete' {
                    if ($record.payload.turn_id) {
                        $currentTurnId = [string]$record.payload.turn_id
                        $lastTurnId = $currentTurnId
                    }

                    $lastTurnCompletedTime = $eventTime
                    $lastTurnSummary = Shorten-Text (Convert-StructuredText $record.payload.last_agent_message) 140
                    $lastTurnCompletedNormally = -not [string]::IsNullOrWhiteSpace($lastTurnSummary)
                    $lastOutcomeTime = $eventTime

                    if ($lastTurnCompletedNormally) {
                        $lastOutcomeCode = 'Completed'
                        $lastOutcomeDetail = $lastTurnSummary
                    } else {
                        $lastOutcomeCode = if ($lastSessionErrorTime) { 'Error' } else { 'Incomplete' }
                        $lastOutcomeDetail = if ($lastSessionErrorText) { $lastSessionErrorText } else { 'turn ended without final response' }
                    }
                }
            }
        }
    } catch { }

    $result = [pscustomobject]@{
        SessionId = $cacheKey
        Path = $Session.Path
        FileLength = $file.Length
        FileLastWriteTimeUtc = $file.LastWriteTimeUtc
        LastTurnId = $lastTurnId
        LastTurnStartedTime = $lastTurnStartedTime
        LastTurnCompletedTime = $lastTurnCompletedTime
        LastTurnCompletedNormally = $lastTurnCompletedNormally
        LastTurnSummary = $lastTurnSummary
        LastSessionErrorTime = $lastSessionErrorTime
        LastSessionErrorText = $lastSessionErrorText
        LastOutcomeCode = $lastOutcomeCode
        LastOutcomeTime = $lastOutcomeTime
        LastOutcomeDetail = $lastOutcomeDetail
    }

    $script:CodexSessionOutcomeCache[$cacheKey] = $result
    return $result
}

function Load-HistoryNames {
    if ((Get-Date) -lt $script:LastHistoryLoad.AddSeconds(12)) { return }
    $script:LastHistoryLoad = Get-Date

    $codexNames = @{}
    if (Test-Path $script:CodexHistoryPath) {
        try {
            Get-Content -Path $script:CodexHistoryPath -Encoding UTF8 | ForEach-Object {
                if ([string]::IsNullOrWhiteSpace($_)) { return }
                try {
                    $entry = $_ | ConvertFrom-Json
                    if ($entry.session_id -and $entry.text -and -not $codexNames.ContainsKey([string]$entry.session_id)) {
                        $codexNames[[string]$entry.session_id] = Shorten-Text ([string]$entry.text) 72
                    }
                } catch { }
            }
        } catch { }
    }
    $script:HistoryNames = $codexNames

    $claudeNames = @{}
    if (Test-Path $script:ClaudeHistoryPath) {
        try {
            Get-Content -Path $script:ClaudeHistoryPath -Encoding UTF8 | ForEach-Object {
                if ([string]::IsNullOrWhiteSpace($_)) { return }
                try {
                    $entry = $_ | ConvertFrom-Json
                    if ($entry.sessionId -and $entry.display -and -not $claudeNames.ContainsKey([string]$entry.sessionId)) {
                        $claudeNames[[string]$entry.sessionId] = Shorten-Text ([string]$entry.display) 72
                    }
                } catch { }
            }
        } catch { }
    }
    $script:ClaudeHistoryNames = $claudeNames
}

function Get-ClaudeMessageBlocks {
    param($Message)
    if ($null -eq $Message) { return @() }
    if ($Message.content -is [System.Collections.IEnumerable] -and -not ($Message.content -is [string])) {
        return @($Message.content)
    }
    if ($null -ne $Message.content) {
        return @($Message.content)
    }
    return @()
}

function Load-ClaudeSessions {
    if ((Get-Date) -lt $script:LastClaudeSessionLoad.AddSeconds($script:ClaudeSessionLoadIntervalSeconds)) { return }
    $script:LastClaudeSessionLoad = Get-Date

    $script:ClaudeSessions = @()
    $script:ClaudeSessionByProcessId = @{}
    if (-not (Test-Path $script:ClaudeSessionRoot)) { return }

    $projectFilesBySession = @{}
    if (Test-Path $script:ClaudeProjectRoot) {
        Get-ChildItem -Recurse -File -Path $script:ClaudeProjectRoot -Filter '*.jsonl' | ForEach-Object {
            $sessionId = [string]$_.BaseName
            if (-not $projectFilesBySession.ContainsKey($sessionId) -or $_.LastWriteTime -gt $projectFilesBySession[$sessionId].LastWriteTime) {
                $projectFilesBySession[$sessionId] = $_
            }
        }
    }

    $items = @()
    Get-ChildItem -File -Path $script:ClaudeSessionRoot -Filter '*.json' | ForEach-Object {
        try {
            $obj = Get-Content -Raw -LiteralPath $_.FullName -Encoding UTF8 | ConvertFrom-Json
            if (-not $obj.sessionId -or -not $obj.pid) { return }

            $sessionId = [string]$obj.sessionId
            $projectFile = if ($projectFilesBySession.ContainsKey($sessionId)) { $projectFilesBySession[$sessionId] } else { $null }
            $entry = [pscustomobject]@{
                ProcessId = [int]$obj.pid
                SessionId = $sessionId
                Cwd = [string]$obj.cwd
                StartTime = Convert-UnixMillisecondsToDateTime $obj.startedAt
                Path = if ($projectFile) { $projectFile.FullName } else { '' }
                LastWriteTime = if ($projectFile) { $projectFile.LastWriteTime } else { $_.LastWriteTime }
                StateFilePath = $_.FullName
            }
            $items += $entry
            $script:ClaudeSessionByProcessId[[int]$obj.pid] = $entry
        } catch { }
    }

    $script:ClaudeSessions = @($items | Sort-Object StartTime -Descending)
}

function Find-ClaudeSession {
    param($Process)
    if (-not $Process) { return $null }
    Load-ClaudeSessions
    $processId = [int]$Process.ProcessId
    if ($script:ClaudeSessionByProcessId.ContainsKey($processId)) {
        return $script:ClaudeSessionByProcessId[$processId]
    }
    return $null
}

function Load-ClaudeTelemetryIndex {
    if ((Get-Date) -lt $script:LastClaudeTelemetryIndexLoad.AddSeconds($script:ClaudeTelemetryLoadIntervalSeconds)) { return }
    $script:LastClaudeTelemetryIndexLoad = Get-Date
    $script:ClaudeTelemetryFilesBySession = @{}
    if (-not (Test-Path $script:ClaudeTelemetryRoot)) { return }

    Get-ChildItem -File -Path $script:ClaudeTelemetryRoot -Filter '*.json' | ForEach-Object {
        if ($_.Name -notmatch '^[^.]+\.([0-9a-fA-F-]{36})\.') { return }
        $sessionId = ([string]$Matches[1]).ToLowerInvariant()
        if (-not $script:ClaudeTelemetryFilesBySession.ContainsKey($sessionId)) {
            $script:ClaudeTelemetryFilesBySession[$sessionId] = New-Object System.Collections.Generic.List[object]
        }
        [void]$script:ClaudeTelemetryFilesBySession[$sessionId].Add($_)
    }

    foreach ($sessionId in @($script:ClaudeTelemetryFilesBySession.Keys)) {
        $script:ClaudeTelemetryFilesBySession[$sessionId] = @(
            $script:ClaudeTelemetryFilesBySession[$sessionId] |
                Sort-Object LastWriteTime -Descending
        )
    }
}

function Get-ClaudeTelemetryFileState {
    param(
        [System.IO.FileInfo]$File,
        [string]$SessionId
    )
    if ($null -eq $File -or [string]::IsNullOrWhiteSpace($SessionId)) { return $null }

    $cacheKey = [string]$File.FullName
    $fileInfo = $null
    try { $fileInfo = Get-Item -LiteralPath $File.FullName } catch { return $null }

    if ($script:ClaudeTelemetryFileStateCache.ContainsKey($cacheKey)) {
        $cached = $script:ClaudeTelemetryFileStateCache[$cacheKey]
        if (
            $cached -and
            $cached.LastWriteTime -eq $fileInfo.LastWriteTime -and
            $cached.Length -eq $fileInfo.Length -and
            $cached.SessionId -eq $SessionId
        ) {
            return $cached.State
        }
    }

    $state = $null
    $lines = @(Read-FileTailLines -Path $fileInfo.FullName -MaxBytes 1572864 -MaxLines 1200)
    [array]::Reverse($lines)

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $entry = $null
        try { $entry = $line | ConvertFrom-Json } catch { continue }
        if (-not $entry -or -not $entry.event_data) { continue }

        $eventData = $entry.event_data
        $eventSessionId = [string]$eventData.session_id
        if ([string]::IsNullOrWhiteSpace($eventSessionId)) { continue }
        if ($eventSessionId.ToLowerInvariant() -ne $SessionId) { continue }

        $timestamp = Convert-SerializedDateTime $eventData.client_timestamp
        if ($null -eq $timestamp) { continue }

        $eventName = [string]$eventData.event_name
        switch ($eventName) {
            'tengu_api_query' {
                $state = [pscustomobject]@{
                    SessionId = $SessionId
                    Source = 'Telemetry'
                    LastKind = 'api_query'
                    LastTimestamp = $timestamp
                    LastAssistantText = ''
                    LastErrorText = ''
                    LastErrorTime = $null
                }
            }
            'tengu_api_success' {
                $state = [pscustomobject]@{
                    SessionId = $SessionId
                    Source = 'Telemetry'
                    LastKind = 'api_success'
                    LastTimestamp = $timestamp
                    LastAssistantText = ''
                    LastErrorText = ''
                    LastErrorTime = $null
                }
            }
            'tengu_api_error' {
                $state = [pscustomobject]@{
                    SessionId = $SessionId
                    Source = 'Telemetry'
                    LastKind = 'api_error'
                    LastTimestamp = $timestamp
                    LastAssistantText = ''
                    LastErrorText = 'Claude request failed'
                    LastErrorTime = $timestamp
                }
            }
        }

        if ($state) { break }
    }

    $script:ClaudeTelemetryFileStateCache[$cacheKey] = [pscustomobject]@{
        SessionId = $SessionId
        LastWriteTime = $fileInfo.LastWriteTime
        Length = $fileInfo.Length
        State = $state
    }

    return $state
}

function Get-ClaudeTelemetryState {
    param($Session)
    if ($null -eq $Session -or [string]::IsNullOrWhiteSpace([string]$Session.SessionId)) { return $null }

    Load-ClaudeTelemetryIndex
    $sessionId = ([string]$Session.SessionId).ToLowerInvariant()
    if (-not $script:ClaudeTelemetryFilesBySession.ContainsKey($sessionId)) { return $null }

    foreach ($file in @($script:ClaudeTelemetryFilesBySession[$sessionId])) {
        $state = Get-ClaudeTelemetryFileState -File $file -SessionId $sessionId
        if ($state) { return $state }
    }

    return $null
}

function Get-ClaudeSessionState {
    param($Session)
    if ($null -eq $Session -or [string]::IsNullOrWhiteSpace([string]$Session.Path) -or -not (Test-Path $Session.Path)) {
        return $null
    }

    $cacheKey = [string]$Session.SessionId
    $file = $null
    try { $file = Get-Item -LiteralPath $Session.Path } catch { return $null }

    if ($script:ClaudeSessionStateCache.ContainsKey($cacheKey)) {
        $cached = $script:ClaudeSessionStateCache[$cacheKey]
        if (
            $cached -and
            $cached.LastWriteTime -eq $file.LastWriteTime -and
            $cached.Length -eq $file.Length
        ) {
            return $cached.State
        }
    }

    $state = [pscustomobject]@{
        SessionId = [string]$Session.SessionId
        Source = 'Transcript'
        LastKind = ''
        LastTimestamp = $null
        LastAssistantText = ''
        LastErrorText = ''
        LastErrorTime = $null
        LastUserPromptTime = $null
        LastAssistantTextTime = $null
        LastWorkingTime = $null
        LastTurnCompleteTime = $null
    }

    foreach ($line in (Read-FileTailLines -Path $Session.Path -MaxBytes 1572864 -MaxLines 260)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $entry = $null
        try { $entry = $line | ConvertFrom-Json } catch { continue }
        if (-not $entry) { continue }

        $timestamp = Convert-SerializedDateTime $entry.timestamp
        if ($null -eq $timestamp -and $entry.snapshot -and $entry.snapshot.timestamp) {
            $timestamp = Convert-SerializedDateTime $entry.snapshot.timestamp
        }
        if ($null -eq $timestamp) { continue }

        switch ([string]$entry.type) {
            'user' {
                $isToolResult = $false
                $isToolError = $false
                $isInterrupted = $false
                $userText = Shorten-Text (Convert-StructuredText $entry.message.content) 120
                if ($userText -eq '[Request interrupted by user]') {
                    $isInterrupted = $true
                }
                if ($entry.toolUseResult) {
                    $isToolResult = $true
                    if ($entry.toolUseResult.interrupted) {
                        $isInterrupted = $true
                    }
                    if ([string]$entry.toolUseResult -match '^(Error:|.*\bis_error\b.*)$') {
                        $isToolError = $true
                    }
                }

                $blocks = @(Get-ClaudeMessageBlocks $entry.message)
                foreach ($block in $blocks) {
                    if ($block.type -eq 'tool_result') {
                        $isToolResult = $true
                        if ($block.is_error -or ([string]$block.content -match '^(Error:|.*exceeds maximum allowed tokens.*)$')) {
                            $isToolError = $true
                        }
                    }
                }

                if ($isToolResult) {
                    $state.LastTimestamp = $timestamp
                    if ($isInterrupted) {
                        $state.LastKind = 'interrupted'
                    } elseif ($isToolError) {
                        $state.LastWorkingTime = $timestamp
                        $state.LastKind = 'tool_error'
                        $state.LastErrorTime = $timestamp
                        $state.LastErrorText = Shorten-Text (Convert-StructuredText $entry.message.content) 120
                    } else {
                        $state.LastWorkingTime = $timestamp
                        $state.LastKind = 'tool_result'
                    }
                } else {
                    $state.LastTimestamp = $timestamp
                    if ($isInterrupted) {
                        $state.LastKind = 'interrupted'
                    } else {
                        $state.LastUserPromptTime = $timestamp
                        $state.LastWorkingTime = $timestamp
                        $state.LastKind = 'user_prompt'
                    }
                }
            }
            'assistant' {
                $blocks = @(Get-ClaudeMessageBlocks $entry.message)
                $hasToolUse = $false
                $hasThinking = $false
                $assistantText = ''
                $stopReason = ''
                if ($entry.message -and $entry.message.stop_reason) {
                    $stopReason = [string]$entry.message.stop_reason
                }

                foreach ($block in $blocks) {
                    if ($block.type -eq 'tool_use') {
                        $hasToolUse = $true
                    } elseif ($block.type -like '*thinking*') {
                        $hasThinking = $true
                    } elseif ($block.type -eq 'text' -and -not [string]::IsNullOrWhiteSpace([string]$block.text)) {
                        $assistantText = [string]$block.text
                    }
                }

                $state.LastTimestamp = $timestamp
                if ($hasToolUse) {
                    $state.LastKind = 'assistant_tool'
                    $state.LastWorkingTime = $timestamp
                } elseif ($hasThinking) {
                    $state.LastKind = 'assistant_thinking'
                    $state.LastWorkingTime = $timestamp
                } elseif (-not [string]::IsNullOrWhiteSpace($assistantText)) {
                    $state.LastAssistantText = Shorten-Text $assistantText 120
                    if ([string]::IsNullOrWhiteSpace($stopReason) -or $stopReason -eq 'tool_use') {
                        $state.LastKind = 'assistant_text_pending'
                        $state.LastWorkingTime = $timestamp
                    } else {
                        $state.LastKind = 'assistant_text'
                        $state.LastAssistantTextTime = $timestamp
                    }
                }
            }
            'progress' {
                $state.LastTimestamp = $timestamp
                $state.LastWorkingTime = $timestamp
                $state.LastKind = 'progress'
            }
            'system' {
                if ($entry.subtype -eq 'turn_duration') {
                    $state.LastTimestamp = $timestamp
                    $state.LastTurnCompleteTime = $timestamp
                    $state.LastKind = 'turn_complete'
                }
            }
        }
    }

    $script:ClaudeSessionStateCache[$cacheKey] = [pscustomobject]@{
        LastWriteTime = $file.LastWriteTime
        Length = $file.Length
        State = $state
    }

    return $state
}

function Ensure-CodexThreadState {
    param([string]$SessionId)
    if (-not $script:CodexThreadStates.ContainsKey($SessionId)) {
        $script:CodexThreadStates[$SessionId] = [pscustomobject]@{
            SessionId = $SessionId
            IsWorking = $false
            LastLogTime = $null
            LastActivityTime = $null
            LastStartTime = $null
            LastCloseTime = $null
            LastRecoverTime = $null
            LastFatalTime = $null
            LastFatalText = ''
            LastLevel = 'INFO'
            HasRecentDisconnect = $false
            LastEvent = ''
            LastLine = ''
        }
    }
    return $script:CodexThreadStates[$SessionId]
}

function Parse-CodexLogLine {
    param([string]$Line)
    if ($Line -notmatch 'session_loop\{thread_id=([0-9a-f-]+)\}') { return }
    $sessionId = $Matches[1]
    $state = Ensure-CodexThreadState $sessionId

    $timestamp = Get-Date
    if ($Line -match '^(\d{4}-\d{2}-\d{2}T[0-9:\.]+Z)') {
        try { $timestamp = ([datetimeoffset]::Parse($Matches[1])).LocalDateTime } catch { }
    }
    $level = 'INFO'
    if ($Line -match '^\d{4}-\d{2}-\d{2}T[0-9:\.]+Z\s+([A-Z]+)\s+') {
        $level = $Matches[1]
    }

    $state.LastLogTime = $timestamp
    $state.LastLevel = $level
    $state.LastLine = Shorten-Text $Line 140

    if ($Line -match 'codex_core::tasks: new') {
        $state.IsWorking = $true
        $state.LastStartTime = $timestamp
        $state.LastActivityTime = $timestamp
        $state.LastRecoverTime = $timestamp
        $state.HasRecentDisconnect = $false
        $state.LastEvent = 'task started'
        return
    }

    if ($Line -match 'codex_core::tasks: close') {
        $state.IsWorking = $false
        $state.LastCloseTime = $timestamp
        $state.LastActivityTime = $timestamp
        $state.LastRecoverTime = $timestamp
        $state.HasRecentDisconnect = $false
        $state.LastEvent = 'waiting for input'
        return
    }

    if ($Line -match 'stream disconnected') {
        $state.LastActivityTime = $timestamp
        $state.HasRecentDisconnect = $true
        $state.LastEvent = 'network retry'
        return
    }

    if ($Line -match 'ToolCall:|model_client\.stream_responses_api|submission_dispatch\{.*codex\.op="user_input"|codex_core::client: new') {
        $state.LastActivityTime = $timestamp
        $state.LastRecoverTime = $timestamp
        $state.HasRecentDisconnect = $false
        if ($null -eq $state.LastCloseTime -or $timestamp -gt $state.LastCloseTime) {
            $state.IsWorking = $true
            if ([string]::IsNullOrWhiteSpace($state.LastEvent) -or $state.LastEvent -eq 'network retry') {
                $state.LastEvent = 'active'
            }
        }
        return
    }

    $looksLikeToolPayload = $Line -match 'ToolCall:'
    $isFatal = $false
    if (-not $looksLikeToolPayload) {
        if ($level -eq 'ERROR' -and $Line -match '\b(panic|fatal)\b|unrecoverable|session.+crash|exception') {
            $isFatal = $true
        } elseif ($Line -match '\bpanic\b|\bfatal\b') {
            $isFatal = $true
        }
    }

    if ($isFatal) {
        $state.LastFatalTime = $timestamp
        $state.LastFatalText = Shorten-Text $Line 120
        $state.LastEvent = 'fatal error'
    }
}

function Read-CodexLogUpdates {
    if (-not (Test-Path $script:CodexLogPath)) { return }
    try {
        $file = Get-Item $script:CodexLogPath
        $length = $file.Length
        if ($null -eq $script:CodexLogPosition) {
            $script:CodexLogPosition = [Math]::Max(0, $length - $script:InitialCodexLogReadBytes)
        }
        if ($length -lt $script:CodexLogPosition) {
            $script:CodexLogPosition = [Math]::Max(0, $length - $script:InitialCodexLogReadBytes)
        }
        if ($length -eq $script:CodexLogPosition) { return }

        $stream = [System.IO.File]::Open($script:CodexLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $stream.Seek($script:CodexLogPosition, [System.IO.SeekOrigin]::Begin) | Out-Null
            $bytesToRead = [int]($length - $script:CodexLogPosition)
            $buffer = New-Object byte[] $bytesToRead
            $read = $stream.Read($buffer, 0, $bytesToRead)
            $script:CodexLogPosition = $length
        } finally {
            $stream.Dispose()
        }

        if ($read -le 0) { return }
        $text = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $read)
        foreach ($line in ($text -split "`r?`n")) {
            if ($line) { Parse-CodexLogLine $line }
        }
    } catch { }
}

function Load-LaunchRecords {
    $records = @()
    if (-not (Test-Path $script:LaunchRoot)) { return $records }

    Get-ChildItem -File -Path $script:LaunchRoot -Filter '*.json' | ForEach-Object {
        try {
            $obj = Get-Content -Raw -Path $_.FullName -Encoding UTF8 | ConvertFrom-Json
            $startedAt = Convert-SerializedDateTime $obj.StartedAt
            if ($startedAt) {
                $records += [pscustomobject]@{
                    AgentType = if ($obj.AgentType) { [string]$obj.AgentType } else { 'Codex' }
                    InstanceId = [string]$obj.InstanceId
                    Title = [string]$obj.Title
                    StartedAt = $startedAt
                    WorkingDirectory = [string]$obj.WorkingDirectory
                    LauncherProcessId = if ($obj.LauncherProcessId) { [int]$obj.LauncherProcessId } else { 0 }
                    ParentProcessId = if ($obj.ParentProcessId) { [int]$obj.ParentProcessId } else { 0 }
                    RootShellProcessId = if ($obj.RootShellProcessId) { [int]$obj.RootShellProcessId } else { 0 }
                    LauncherProcessName = if ($obj.LauncherProcessName) { [string]$obj.LauncherProcessName } else { '' }
                    ParentProcessName = if ($obj.ParentProcessName) { [string]$obj.ParentProcessName } else { '' }
                    RootShellProcessName = if ($obj.RootShellProcessName) { [string]$obj.RootShellProcessName } else { '' }
                    LauncherProcessStartedAt = Convert-SerializedDateTime $obj.LauncherProcessStartedAt
                    ParentProcessStartedAt = Convert-SerializedDateTime $obj.ParentProcessStartedAt
                    RootShellProcessStartedAt = Convert-SerializedDateTime $obj.RootShellProcessStartedAt
                    FilePath = $_.FullName
                }
            }
        } catch { }
    }

    return @($records | Sort-Object StartedAt)
}

function Get-AgentProcesses {
    $items = @()
    $all = Get-ProcessSnapshot
    $shellNames = @('cmd.exe', 'powershell.exe', 'pwsh.exe', 'bash.exe', 'wsl.exe')
    Load-ClaudeSessions

    foreach ($proc in $all.Values) {
        $commandLine = [string]$proc.CommandLine
        $agentType = $null
        if ($proc.Name -eq 'node.exe' -and $commandLine -match '@openai[\\/]+codex[\\/]+bin[\\/]+codex\.js') {
            $agentType = 'Codex'
        } elseif ($proc.Name -eq 'claude.exe') {
            $agentType = 'Claude'
        }

        if (-not $agentType) { continue }

        $parent = $null
        if ($all.ContainsKey([int]$proc.ParentProcessId)) { $parent = $all[[int]$proc.ParentProcessId] }
        $lineage = Get-AgentProcessLineageInfo -Process $proc -ProcessMap $all -AgentType $agentType
        $groupKey = if ($lineage.RootShellProcessId -gt 0) {
            '{0}:shell:{1}' -f $agentType, $lineage.RootShellProcessId
        } elseif ([int]$proc.ParentProcessId -gt 0) {
            '{0}:parent:{1}' -f $agentType, ([int]$proc.ParentProcessId)
        } else {
            '{0}:pid:{1}' -f $agentType, ([int]$proc.ProcessId)
        }
        $items += [pscustomobject]@{
            AgentType = $agentType
            ProcessId = [int]$proc.ProcessId
            ParentProcessId = [int]$proc.ParentProcessId
            StartTime = Convert-WmiTime $proc.CreationDate
            ParentStartTime = if ($parent) { Convert-WmiTime $parent.CreationDate } else { Convert-WmiTime $proc.CreationDate }
            CommandLine = $commandLine
            ParentCommandLine = if ($parent) { [string]$parent.CommandLine } else { '' }
            ParentName = if ($parent) { [string]$parent.Name } else { '' }
            RootShellProcessId = [int]$lineage.RootShellProcessId
            GroupKey = $groupKey
            HasSameAgentAncestor = [bool]$lineage.HasSameAgentAncestor
            DepthToShell = [int]$lineage.DepthToShell
            IsDirectShellChild = [bool]$lineage.IsDirectShellChild
            HasLiveClaudeSession = if ($agentType -eq 'Claude') { $script:ClaudeSessionByProcessId.ContainsKey([int]$proc.ProcessId) } else { $false }
        }
    }

    $result = @()
    $result += @($items | Where-Object { $_.AgentType -ne 'Claude' })

    $claudeCandidates = @(
        $items |
        Where-Object { $_.AgentType -eq 'Claude' -and -not $_.HasSameAgentAncestor }
    )

    foreach ($group in ($claudeCandidates | Group-Object GroupKey)) {
        $selected = $group.Group |
            Sort-Object `
                @{ Expression = { if ($_.HasLiveClaudeSession) { 0 } else { 1 } } }, `
                @{ Expression = { if ($_.IsDirectShellChild) { 0 } else { 1 } } }, `
                @{ Expression = { $_.DepthToShell } }, `
                @{ Expression = { if ($shellNames -contains $_.ParentName) { 0 } else { 1 } } }, `
                StartTime, `
                ProcessId |
            Select-Object -First 1
        if ($selected) {
            $result += $selected
        }
    }

    return @($result | Sort-Object AgentType, StartTime, ProcessId)
}

function Get-ProcessSnapshot {
    $all = @{}
    Get-CimInstance Win32_Process | ForEach-Object {
        $all[[int]$_.ProcessId] = [pscustomobject]@{
            Name = [string]$_.Name
            ProcessId = [int]$_.ProcessId
            ParentProcessId = [int]$_.ParentProcessId
            CreationDate = Convert-WmiTime $_.CreationDate
            CommandLine = [string]$_.CommandLine
        }
    }
    return $all
}

function Get-ProcessAncestors {
    param(
        [int]$ProcessId,
        [hashtable]$ProcessMap
    )
    $items = @()
    $seen = @{}
    $currentId = $ProcessId
    while ($ProcessMap.ContainsKey($currentId) -and -not $seen.ContainsKey($currentId)) {
        $seen[$currentId] = $true
        $proc = $ProcessMap[$currentId]
        $items += $proc
        if (-not $proc.ParentProcessId -or $proc.ParentProcessId -eq $currentId) { break }
        $currentId = [int]$proc.ParentProcessId
    }
    return @($items)
}

function Get-AgentProcessLineageInfo {
    param(
        $Process,
        [hashtable]$ProcessMap,
        [string]$AgentType
    )

    $shellNames = @('cmd.exe', 'powershell.exe', 'pwsh.exe', 'bash.exe', 'wsl.exe')
    $stopNames = @('explorer.exe', 'Code.exe', 'devenv.exe')
    $rootShellPid = 0
    $depthToShell = 999
    $hasSameAgentAncestor = $false
    $isDirectShellChild = $false
    $depth = 0

    foreach ($ancestor in (Get-ProcessAncestors -ProcessId ([int]$Process.ProcessId) -ProcessMap $ProcessMap)) {
        if ($depth -gt 0 -and $stopNames -contains $ancestor.Name) { break }

        if ($depth -gt 0) {
            if ($AgentType -eq 'Claude' -and $ancestor.Name -eq 'claude.exe') {
                $hasSameAgentAncestor = $true
            }
            if (
                $AgentType -eq 'Codex' -and
                $ancestor.Name -eq 'node.exe' -and
                [string]$ancestor.CommandLine -match '@openai[\\/]+codex[\\/]+bin[\\/]+codex\.js'
            ) {
                $hasSameAgentAncestor = $true
            }
        }

        if ($shellNames -contains $ancestor.Name) {
            $rootShellPid = [int]$ancestor.ProcessId
            if ($depthToShell -eq 999) {
                $depthToShell = $depth
                $isDirectShellChild = $depth -eq 1
            }
        }

        $depth += 1
    }

    return [pscustomobject]@{
        RootShellProcessId = $rootShellPid
        DepthToShell = $depthToShell
        HasSameAgentAncestor = $hasSameAgentAncestor
        IsDirectShellChild = $isDirectShellChild
    }
}

function Get-DescendantProcessIds {
    param(
        [int]$RootProcessId,
        [hashtable]$ProcessMap
    )
    $queue = New-Object System.Collections.Generic.Queue[int]
    $seen = New-Object System.Collections.Generic.HashSet[int]
    $result = New-Object System.Collections.Generic.List[int]
    $queue.Enqueue($RootProcessId)
    [void]$seen.Add($RootProcessId)

    while ($queue.Count -gt 0) {
        $currentId = $queue.Dequeue()
        foreach ($proc in $ProcessMap.Values) {
            if ([int]$proc.ParentProcessId -eq $currentId -and -not $seen.Contains([int]$proc.ProcessId)) {
                [void]$seen.Add([int]$proc.ProcessId)
                $queue.Enqueue([int]$proc.ProcessId)
                $result.Add([int]$proc.ProcessId) | Out-Null
            }
        }
    }

    return @($result)
}

function Add-UniqueProcessId {
    param(
        [System.Collections.Generic.List[int]]$List,
        [int]$ProcessId
    )
    if ($ProcessId -gt 0 -and -not $List.Contains($ProcessId)) {
        $List.Add($ProcessId) | Out-Null
    }
}

function Get-InstanceControlRootProcessId {
    param(
        $Instance,
        [hashtable]$ProcessMap
    )
    $agentPid = [int]$Instance.Process.ProcessId
    $rootPid = $agentPid
    $shellNames = @('cmd.exe', 'powershell.exe', 'pwsh.exe', 'bash.exe', 'wsl.exe')
    $stopNames = @('explorer.exe', 'Code.exe', 'devenv.exe')

    foreach ($proc in (Get-ProcessAncestors -ProcessId $agentPid -ProcessMap $ProcessMap)) {
        if ($stopNames -contains $proc.Name) { break }
        if ($shellNames -contains $proc.Name) {
            $rootPid = [int]$proc.ProcessId
        }
    }

    if ($Instance.Launch -and $Instance.Launch.RootShellProcessId -gt 0 -and $ProcessMap.ContainsKey([int]$Instance.Launch.RootShellProcessId)) {
        $rootPid = [int]$Instance.Launch.RootShellProcessId
    }

    if ($Instance.Launch -and $Instance.Launch.ParentProcessId -gt 0 -and $ProcessMap.ContainsKey([int]$Instance.Launch.ParentProcessId)) {
        $launchParent = $ProcessMap[[int]$Instance.Launch.ParentProcessId]
        if ($shellNames -contains $launchParent.Name) {
            $rootPid = [int]$launchParent.ProcessId
        }
    }

    return $rootPid
}

function Get-InstanceWindowProcessIds {
    param(
        $Instance,
        [hashtable]$ProcessMap
    )
    $ordered = New-Object 'System.Collections.Generic.List[int]'
    $agentPid = [int]$Instance.Process.ProcessId
    $rootPid = Get-InstanceControlRootProcessId -Instance $Instance -ProcessMap $ProcessMap
    $stopNames = @('explorer.exe', 'Code.exe', 'devenv.exe')

    Add-UniqueProcessId -List $ordered -ProcessId $rootPid
    Add-UniqueProcessId -List $ordered -ProcessId $agentPid

    if ($Instance.Launch -and $Instance.Launch.LauncherProcessId -gt 0) {
        Add-UniqueProcessId -List $ordered -ProcessId ([int]$Instance.Launch.LauncherProcessId)
    }
    if ($Instance.Launch -and $Instance.Launch.ParentProcessId -gt 0) {
        Add-UniqueProcessId -List $ordered -ProcessId ([int]$Instance.Launch.ParentProcessId)
    }

    foreach ($proc in (Get-ProcessAncestors -ProcessId $agentPid -ProcessMap $ProcessMap)) {
        if ($stopNames -contains $proc.Name) { break }
        Add-UniqueProcessId -List $ordered -ProcessId ([int]$proc.ProcessId)
    }
    foreach ($descendantPid in (Get-DescendantProcessIds -RootProcessId $rootPid -ProcessMap $ProcessMap)) {
        Add-UniqueProcessId -List $ordered -ProcessId $descendantPid
    }

    return @($ordered)
}

function Find-BestCodexSession {
    param(
        [datetime]$ReferenceTime,
        [datetime]$AlternateTime = [datetime]::MinValue,
        [hashtable]$AssignedSessions = $null
    )
    if ($ReferenceTime -eq [datetime]::MinValue -and $AlternateTime -eq [datetime]::MinValue) {
        return $null
    }

    $bestSession = $null
    $bestScore = [double]::MaxValue
    $anchorTimes = @($ReferenceTime, $AlternateTime | Where-Object { $_ -ne [datetime]::MinValue })
    $now = Get-Date
    foreach ($candidate in $script:CodexSessions) {
        if ($AssignedSessions -and $AssignedSessions.ContainsKey($candidate.SessionId)) { continue }

        $threadState = $null
        if ($script:CodexThreadStates.ContainsKey($candidate.SessionId)) {
            $threadState = $script:CodexThreadStates[$candidate.SessionId]
        }

        $hasStrongSignal = $false
        $startsTooLate = $true
        $bestStartDelta = [double]::MaxValue
        $latestActivityTime = $null

        if ($candidate.LastWriteTime) {
            $latestActivityTime = [datetime]$candidate.LastWriteTime
        }

        if ($threadState) {
            foreach ($activityTime in @($threadState.LastStartTime, $threadState.LastActivityTime, $threadState.LastCloseTime, $threadState.LastFatalTime)) {
                if ($activityTime -and ($null -eq $latestActivityTime -or $activityTime -gt $latestActivityTime)) {
                    $latestActivityTime = [datetime]$activityTime
                }
            }
        }

        foreach ($anchorTime in $anchorTimes) {
            if ($candidate.StartTime -le $anchorTime.AddSeconds($script:SessionFutureStartGraceSeconds)) {
                $startsTooLate = $false
            }

            $startDelta = [Math]::Abs(($candidate.StartTime - $anchorTime).TotalSeconds)
            if ($startDelta -lt $bestStartDelta) {
                $bestStartDelta = $startDelta
            }

            if ($startDelta -le $script:SessionStartMatchSeconds) {
                $hasStrongSignal = $true
            }

            if ($latestActivityTime -and $latestActivityTime -ge $anchorTime.AddSeconds(-$script:SessionActivityLeadSeconds)) {
                $hasStrongSignal = $true
            }
        }

        if ($startsTooLate) { continue }
        if (-not $hasStrongSignal) { continue }

        $recencyPenalty = 86400.0
        if ($latestActivityTime) {
            $recencyPenalty = [Math]::Max(0, ($now - $latestActivityTime).TotalSeconds)
        }
        $startPenalty = if ($bestStartDelta -le $script:SessionStartMatchSeconds) {
            $bestStartDelta / 1000.0
        } else {
            $script:SessionStartPenaltySeconds + ($bestStartDelta / 1000.0)
        }
        $score = $recencyPenalty + $startPenalty

        if ($score -lt $bestScore) {
            $bestSession = $candidate
            $bestScore = $score
        }
    }

    return $bestSession
}

function Find-BestLaunchRecord {
    param(
        [string]$AgentType,
        [datetime]$ReferenceTime,
        [object[]]$Launches,
        [hashtable]$AssignedLaunches = $null,
        $Process = $null,
        [hashtable]$ProcessMap = $null
    )
    if ($ReferenceTime -eq [datetime]::MinValue -and -not $Process) { return $null }

    $lineageIds = $null
    $lineageMap = $null
    if ($Process -and $ProcessMap) {
        $lineageIds = New-Object System.Collections.Generic.HashSet[int]
        $lineageMap = @{}
        [void]$lineageIds.Add([int]$Process.ProcessId)
        $lineageMap[[int]$Process.ProcessId] = $ProcessMap[[int]$Process.ProcessId]
        foreach ($ancestorProc in (Get-ProcessAncestors -ProcessId ([int]$Process.ProcessId) -ProcessMap $ProcessMap)) {
            [void]$lineageIds.Add([int]$ancestorProc.ProcessId)
            $lineageMap[[int]$ancestorProc.ProcessId] = $ancestorProc
        }
    }

    $bestLaunch = $null
    $bestScore = [double]::MaxValue
    foreach ($record in $Launches) {
        if ($AssignedLaunches -and $AssignedLaunches.ContainsKey($record.InstanceId)) { continue }
        if ($record.AgentType -and $record.AgentType -ne $AgentType) { continue }

        $score = [double]::MaxValue
        if ($lineageIds) {
            $pidScore = 1000
            if ($record.LauncherProcessId -gt 0 -and $lineageIds.Contains([int]$record.LauncherProcessId)) {
                $candidate = $lineageMap[[int]$record.LauncherProcessId]
                if (Test-LaunchProcessIdentity -Launch $record -Process $candidate -Role 'Launcher') {
                    $pidScore = [Math]::Min($pidScore, 0)
                }
            }
            if ($record.ParentProcessId -gt 0 -and $lineageIds.Contains([int]$record.ParentProcessId)) {
                $candidate = $lineageMap[[int]$record.ParentProcessId]
                if (Test-LaunchProcessIdentity -Launch $record -Process $candidate -Role 'Parent') {
                    $pidScore = [Math]::Min($pidScore, 1)
                }
            }
            if ($record.RootShellProcessId -gt 0 -and $lineageIds.Contains([int]$record.RootShellProcessId)) {
                $candidate = $lineageMap[[int]$record.RootShellProcessId]
                if (Test-LaunchProcessIdentity -Launch $record -Process $candidate -Role 'RootShell') {
                    $pidScore = [Math]::Min($pidScore, 2)
                }
            }
            if ($pidScore -lt 1000) {
                $score = $pidScore
            }
        }

        if ($ReferenceTime -ne [datetime]::MinValue) {
            $score = [Math]::Min($score, [Math]::Abs(($record.StartedAt - $ReferenceTime).TotalSeconds))
        }
        if ($score -lt $bestScore -and $score -lt $script:LaunchMatchWindowSeconds) {
            $bestLaunch = $record
            $bestScore = $score
        }
    }

    return $bestLaunch
}

function Get-LaunchExpectedProcessNames {
    param([string]$Role)
    switch ($Role) {
        'Launcher' { return @('powershell.exe', 'pwsh.exe') }
        default { return @('cmd.exe', 'powershell.exe', 'pwsh.exe', 'bash.exe', 'wsl.exe', 'conhost.exe', 'openconsole.exe', 'WindowsTerminal.exe', 'windowsterminal.exe') }
    }
}

function Get-LaunchRecordedProcessName {
    param(
        $Launch,
        [string]$Role
    )
    switch ($Role) {
        'Launcher' { return [string]$Launch.LauncherProcessName }
        'Parent' { return [string]$Launch.ParentProcessName }
        'RootShell' { return [string]$Launch.RootShellProcessName }
    }
    return ''
}

function Get-LaunchRecordedProcessStartTime {
    param(
        $Launch,
        [string]$Role
    )
    switch ($Role) {
        'Launcher' { return $Launch.LauncherProcessStartedAt }
        'Parent' { return $Launch.ParentProcessStartedAt }
        'RootShell' { return $Launch.RootShellProcessStartedAt }
    }
    return $null
}

function Test-LaunchProcessIdentity {
    param(
        $Launch,
        $Process,
        [string]$Role
    )
    if (-not $Launch -or -not $Process) { return $false }

    $processName = [string]$Process.Name
    if ([string]::IsNullOrWhiteSpace($processName)) { return $false }

    $expectedNames = @(Get-LaunchExpectedProcessNames -Role $Role)
    if ($expectedNames.Count -gt 0 -and -not ($expectedNames -contains $processName)) {
        return $false
    }

    $recordedName = Get-LaunchRecordedProcessName -Launch $Launch -Role $Role
    if (-not [string]::IsNullOrWhiteSpace($recordedName) -and $recordedName -ne $processName) {
        return $false
    }

    $processStart = Convert-WmiTime $Process.CreationDate
    if ($processStart -gt $Launch.StartedAt.AddSeconds($script:LaunchProcessFutureSkewSeconds)) {
        return $false
    }

    $recordedStart = Get-LaunchRecordedProcessStartTime -Launch $Launch -Role $Role
    if ($recordedStart) {
        if ([Math]::Abs(($processStart - $recordedStart).TotalSeconds) -gt $script:LaunchIdentitySlackSeconds) {
            return $false
        }
    }

    return $true
}

function Get-LiveLaunchProcess {
    param(
        $Launch,
        [hashtable]$ProcessMap
    )
    if (-not $Launch -or -not $ProcessMap) { return $null }

    $candidates = @(
        [pscustomobject]@{ ProcessId = [int]$Launch.LauncherProcessId; Role = 'Launcher' },
        [pscustomobject]@{ ProcessId = [int]$Launch.ParentProcessId; Role = 'Parent' },
        [pscustomobject]@{ ProcessId = [int]$Launch.RootShellProcessId; Role = 'RootShell' }
    )

    foreach ($candidate in $candidates) {
        if ($candidate.ProcessId -le 0 -or -not $ProcessMap.ContainsKey($candidate.ProcessId)) { continue }

        $proc = $ProcessMap[$candidate.ProcessId]
        if (-not (Test-LaunchProcessIdentity -Launch $Launch -Process $proc -Role $candidate.Role)) { continue }
        $parent = $null
        if ($proc.ParentProcessId -gt 0 -and $ProcessMap.ContainsKey([int]$proc.ParentProcessId)) {
            $parent = $ProcessMap[[int]$proc.ParentProcessId]
        }

        return [pscustomobject]@{
            AgentType = $Launch.AgentType
            ProcessId = [int]$proc.ProcessId
            ParentProcessId = [int]$proc.ParentProcessId
            StartTime = Convert-WmiTime $proc.CreationDate
            ParentStartTime = if ($parent) { Convert-WmiTime $parent.CreationDate } else { Convert-WmiTime $proc.CreationDate }
            CommandLine = [string]$proc.CommandLine
            ParentCommandLine = if ($parent) { [string]$parent.CommandLine } else { '' }
            IsSynthetic = $true
            Name = [string]$proc.Name
        }
    }

    return $null
}

function Match-Instances {
    param([object[]]$Processes)
    Load-CodexSessions
    Load-ClaudeSessions
    $launches = Load-LaunchRecords
    $processMap = Get-ProcessSnapshot
    $assignedSessions = @{}
    $assignedLaunches = @{}
    $matched = @()

    foreach ($proc in $Processes) {
        $session = $null
        if ($proc.AgentType -eq 'Codex') {
            $session = Find-BestCodexSession -ReferenceTime $proc.StartTime -AlternateTime $proc.ParentStartTime -AssignedSessions $assignedSessions
        } elseif ($proc.AgentType -eq 'Claude') {
            $session = Find-ClaudeSession -Process $proc
        }

        $launch = Find-BestLaunchRecord -AgentType $proc.AgentType -ReferenceTime $proc.StartTime -Launches $launches -AssignedLaunches $assignedLaunches -Process $proc -ProcessMap $processMap

        if ($session) { $assignedSessions[$session.SessionId] = $true }
        if ($launch) { $assignedLaunches[$launch.InstanceId] = $true }
        $matched += [pscustomobject]@{
            AgentType = $proc.AgentType
            Process = $proc
            Session = $session
            Launch = $launch
            LaunchOnly = $false
        }
    }

    foreach ($launch in ($launches | Sort-Object StartedAt -Descending)) {
        if ($assignedLaunches.ContainsKey($launch.InstanceId)) { continue }
        if ((Get-Date) -gt $launch.StartedAt.AddSeconds($script:LaunchStartupGraceSeconds)) { continue }

        $liveProcess = Get-LiveLaunchProcess -Launch $launch -ProcessMap $processMap
        if (-not $liveProcess) { continue }
        $assignedLaunches[$launch.InstanceId] = $true
        $matched += [pscustomobject]@{
            AgentType = $launch.AgentType
            Process = $liveProcess
            Session = $null
            Launch = $launch
            LaunchOnly = $true
        }
    }

    return $matched
}

function Get-ClaudeStatusDetail {
    param($State)
    if (-not $State) { return 'Claude session mapped; waiting for activity' }
    switch ([string]$State.LastKind) {
        'api_query' { return 'Claude is thinking' }
        'api_success' { return 'Claude is ready' }
        'api_error' {
            if (-not [string]::IsNullOrWhiteSpace($State.LastErrorText)) { return $State.LastErrorText }
            return 'Claude request failed'
        }
        'interrupted' { return 'Claude request interrupted' }
        'user_prompt' { return 'waiting for Claude response' }
        'assistant_thinking' { return 'Claude is thinking' }
        'assistant_tool' { return 'Claude is using tools' }
        'assistant_text_pending' {
            if (-not [string]::IsNullOrWhiteSpace($State.LastAssistantText)) { return $State.LastAssistantText }
            return 'Claude is responding'
        }
        'tool_result' { return 'Claude is processing tool output' }
        'progress' { return 'Claude is running a task' }
        'tool_error' {
            if (-not [string]::IsNullOrWhiteSpace($State.LastErrorText)) { return $State.LastErrorText }
            return 'Claude tool error'
        }
        'assistant_text' {
            if (-not [string]::IsNullOrWhiteSpace($State.LastAssistantText)) { return $State.LastAssistantText }
            return 'Claude replied'
        }
        'turn_complete' {
            if (-not [string]::IsNullOrWhiteSpace($State.LastAssistantText)) { return $State.LastAssistantText }
            return 'Claude is ready'
        }
    }
    return 'Claude session active'
}

function Get-DisplayName {
    param($Instance)
    if ($Instance.AgentType -eq 'Codex') {
        if ($Instance.Session -and $script:HistoryNames.ContainsKey($Instance.Session.SessionId)) {
            return $script:HistoryNames[$Instance.Session.SessionId]
        }
        if ($Instance.Session -and $Instance.Session.Cwd) {
            return "Codex - $(Split-Path -Leaf $Instance.Session.Cwd)"
        }
        return "Codex $($Instance.Process.ProcessId)"
    }

    if ($Instance.AgentType -eq 'Claude') {
        if ($Instance.Session -and $script:ClaudeHistoryNames.ContainsKey($Instance.Session.SessionId)) {
            return $script:ClaudeHistoryNames[$Instance.Session.SessionId]
        }
        if ($Instance.Session -and $Instance.Session.Cwd) {
            return "Claude - $(Split-Path -Leaf $Instance.Session.Cwd)"
        }
        if ($Instance.Launch -and $Instance.Launch.WorkingDirectory) {
            return "Claude - $(Split-Path -Leaf $Instance.Launch.WorkingDirectory)"
        }
        return "Claude $($Instance.Process.ProcessId)"
    }

    return "$($Instance.AgentType) $($Instance.Process.ProcessId)"
}

function Get-InstanceStatus {
    param($Instance)
    if ($Instance.AgentType -eq 'Claude') {
        if ($null -eq $Instance.Session) {
            if ($Instance.LaunchOnly -and $Instance.Launch) {
                return [pscustomobject]@{
                    Code = 'Starting'
                    Label = 'starting'
                    Brush = '#58a6ff'
                    Detail = 'terminal launched; waiting for Claude session'
                }
            }
            return [pscustomobject]@{
                Code = 'Ready'
                Label = 'active'
                Brush = '#35d07f'
                Detail = 'Claude process detected; session not mapped yet'
            }
        }

        $telemetryState = Get-ClaudeTelemetryState $Instance.Session
        $transcriptState = Get-ClaudeSessionState $Instance.Session
        $state = $null
        if ($telemetryState -and $transcriptState) {
            if ($telemetryState.LastTimestamp -and $transcriptState.LastTimestamp) {
                if ($telemetryState.LastTimestamp -ge $transcriptState.LastTimestamp) {
                    $state = $telemetryState
                } else {
                    $state = $transcriptState
                }
            } elseif ($telemetryState.LastTimestamp) {
                $state = $telemetryState
            } else {
                $state = $transcriptState
            }
        } elseif ($telemetryState) {
            $state = $telemetryState
        } else {
            $state = $transcriptState
        }

        if ($state) {
            $detail = Get-ClaudeStatusDetail $state
            switch ([string]$state.LastKind) {
                'api_query' { return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail } }
                'api_success' { return [pscustomobject]@{ Code = 'Ready'; Label = 'ready'; Brush = '#35d07f'; Detail = $detail } }
                'api_error' { return [pscustomobject]@{ Code = 'Error'; Label = 'error'; Brush = '#ff4d5e'; Detail = $detail } }
                'interrupted' { return [pscustomobject]@{ Code = 'Ready'; Label = 'ready'; Brush = '#35d07f'; Detail = $detail } }
                'user_prompt' { return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail } }
                'assistant_thinking' { return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail } }
                'assistant_tool' { return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail } }
                'assistant_text_pending' { return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail } }
                'tool_result' { return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail } }
                'progress' { return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail } }
                'tool_error' { return [pscustomobject]@{ Code = 'Error'; Label = 'error'; Brush = '#ff4d5e'; Detail = $detail } }
                'assistant_text' { return [pscustomobject]@{ Code = 'Ready'; Label = 'ready'; Brush = '#35d07f'; Detail = $detail } }
                'turn_complete' { return [pscustomobject]@{ Code = 'Ready'; Label = 'ready'; Brush = '#35d07f'; Detail = $detail } }
            }
        }

        return [pscustomobject]@{
            Code = 'Ready'
            Label = 'active'
            Brush = '#35d07f'
            Detail = 'Claude session detected; waiting for activity'
        }
    }

    if ($null -eq $Instance.Session) {
        if ($Instance.LaunchOnly -and $Instance.Launch) {
            return [pscustomobject]@{
                Code = 'Starting'
                Label = 'starting'
                Brush = '#58a6ff'
                Detail = 'terminal launched; waiting for Codex session'
            }
        }
        return [pscustomobject]@{ Code = 'Unknown'; Label = 'unknown'; Brush = '#8b949e'; Detail = 'Codex session not mapped yet' }
    }

    $sid = $Instance.Session.SessionId
    $state = $null
    if ($script:CodexThreadStates.ContainsKey($sid)) {
        $state = $script:CodexThreadStates[$sid]
    }

    $sessionOutcome = Get-CodexSessionOutcome $Instance.Session
    $logSaysWorking = $false
    $latestStateActivityTime = $null
    if ($state) {
        $logSaysWorking = $state.IsWorking
        if ($state.LastCloseTime -and ($null -eq $state.LastStartTime -or $state.LastCloseTime -ge $state.LastStartTime)) {
            $logSaysWorking = $false
        }
        $latestStateActivityTime = if ($state.LastActivityTime) { $state.LastActivityTime } else { $state.LastStartTime }
    }

    $sessionOutcomeTime = if ($sessionOutcome) { $sessionOutcome.LastOutcomeTime } else { $null }

    if ($state -and $state.LastFatalTime -and ($null -eq $state.LastRecoverTime -or $state.LastFatalTime -gt $state.LastRecoverTime)) {
        if ($null -eq $sessionOutcomeTime -or $state.LastFatalTime -gt $sessionOutcomeTime) {
            return [pscustomobject]@{ Code = 'Error'; Label = 'error'; Brush = '#ff4d5e'; Detail = $state.LastFatalText }
        }
    }

    if (
        $logSaysWorking -and
        (
            $null -eq $sessionOutcomeTime -or
            $null -eq $latestStateActivityTime -or
            $latestStateActivityTime.AddSeconds(5) -ge $sessionOutcomeTime
        )
    ) {
        $detail = Get-WorkingDetail -State $state
        return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail }
    }

    if ($sessionOutcome) {
        switch ($sessionOutcome.LastOutcomeCode) {
            'Error' {
                $detail = if ($sessionOutcome.LastOutcomeDetail) { $sessionOutcome.LastOutcomeDetail } else { 'session error' }
                return [pscustomobject]@{ Code = 'Error'; Label = 'error'; Brush = '#ff4d5e'; Detail = $detail }
            }
            'Incomplete' {
                $detail = if ($sessionOutcome.LastOutcomeDetail) { $sessionOutcome.LastOutcomeDetail } else { 'turn ended without final response' }
                return [pscustomobject]@{ Code = 'Error'; Label = 'error'; Brush = '#ff4d5e'; Detail = $detail }
            }
            'Started' {
                $detail = Get-WorkingDetail -State $state -Fallback $sessionOutcome.LastOutcomeDetail
                return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail }
            }
            'Completed' {
                $detail = if ($sessionOutcome.LastTurnSummary) { $sessionOutcome.LastTurnSummary } elseif ($state -and $state.LastEvent) { $state.LastEvent } else { 'waiting for input' }
                return [pscustomobject]@{ Code = 'Ready'; Label = 'ready'; Brush = '#35d07f'; Detail = $detail }
            }
        }
    }

    if ($logSaysWorking) {
        $detail = Get-WorkingDetail -State $state
        return [pscustomobject]@{ Code = 'Working'; Label = 'working'; Brush = '#ffc247'; Detail = $detail }
    }

    $detail = if ($state -and $state.LastEvent) { $state.LastEvent } else { 'waiting or no recent activity' }
    return [pscustomobject]@{ Code = 'Ready'; Label = 'ready'; Brush = '#35d07f'; Detail = $detail }
}

function Get-TopWindows {
    $windows = @()
    $callback = [AgentStateNative+EnumWindowsProc]{
        param([IntPtr]$hWnd, [IntPtr]$lParam)
        $visible = [AgentStateNative]::IsWindowVisible($hWnd)
        $iconic = [AgentStateNative]::IsIconic($hWnd)
        if ($visible -or $iconic) {
            $titleSb = New-Object System.Text.StringBuilder 512
            $classSb = New-Object System.Text.StringBuilder 256
            [void][AgentStateNative]::GetWindowText($hWnd, $titleSb, $titleSb.Capacity)
            [void][AgentStateNative]::GetClassName($hWnd, $classSb, $classSb.Capacity)
            [uint32]$procId = 0
            [void][AgentStateNative]::GetWindowThreadProcessId($hWnd, [ref]$procId)
            $title = $titleSb.ToString().Trim()
            $ownerHandle = [AgentStateNative]::GetWindow($hWnd, [AgentStateNative]::GW_OWNER)
            $rootHandle = [AgentStateNative]::GetAncestor($hWnd, [AgentStateNative]::GA_ROOT)
            $rootOwnerHandle = [AgentStateNative]::GetAncestor($hWnd, [AgentStateNative]::GA_ROOTOWNER)
            $script:WindowScratch += [pscustomobject]@{
                Handle = $hWnd
                HandleValue = $hWnd.ToInt64()
                ProcessId = [int]$procId
                Title = $title
                ClassName = $classSb.ToString()
                Visible = $visible
                Iconic = $iconic
                OwnerHandle = $ownerHandle
                OwnerHandleValue = $ownerHandle.ToInt64()
                RootHandle = $rootHandle
                RootHandleValue = $rootHandle.ToInt64()
                RootOwnerHandle = $rootOwnerHandle
                RootOwnerHandleValue = $rootOwnerHandle.ToInt64()
            }
        }
        return $true
    }
    $script:WindowScratch = @()
    [AgentStateNative]::EnumWindows($callback, [IntPtr]::Zero) | Out-Null
    $windows = $script:WindowScratch
    $script:WindowScratch = $null
    return @($windows)
}

function Find-WindowByHandleValue {
    param(
        [object[]]$Windows,
        [long]$HandleValue
    )
    if ($HandleValue -eq 0) { return $null }
    return $Windows | Where-Object { $_.HandleValue -eq $HandleValue } | Select-Object -First 1
}

function Resolve-AgentWindowTarget {
    param(
        $Window,
        [object[]]$Windows
    )
    if (-not $Window) { return $null }

    $candidateMap = @{}
    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($handleValue in @(
        [long]$Window.HandleValue,
        [long]$Window.OwnerHandleValue,
        [long]$Window.RootOwnerHandleValue,
        [long]$Window.RootHandleValue
    )) {
        if ($handleValue -eq 0 -or $candidateMap.ContainsKey($handleValue)) { continue }
        $candidate = Find-WindowByHandleValue -Windows $Windows -HandleValue $handleValue
        if ($candidate) {
            $candidateMap[$handleValue] = $true
            $ordered.Add($candidate) | Out-Null
        }
    }

    if ($ordered.Count -eq 0) { return $Window }

    $resolved = $ordered | Sort-Object `
        @{ Expression = { if ($_.ClassName -eq 'CASCADIA_HOSTING_WINDOW_CLASS') { 0 } else { 1 } } }, `
        @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.Title)) { 1 } else { 0 } } }, `
        @{ Expression = { if ($_.Iconic) { 0 } else { 1 } } } |
        Select-Object -First 1

    if ($resolved) { return $resolved }
    return $Window
}

function Find-AgentWindow {
    param($Instance)
    $windows = Get-TopWindows
    $target = $null
    $terminalWindows = @(
        $windows | Where-Object {
            $_.ClassName -in @('CASCADIA_HOSTING_WINDOW_CLASS', 'PseudoConsoleWindow') -or
            $_.Title -match 'CODEX|Codex|CLAUDE|Claude|Terminal|PowerShell|命令提示符|Command Prompt'
        }
    )

    if ($Instance.Launch -and $Instance.Launch.Title) {
        $escaped = [regex]::Escape($Instance.Launch.Title)
        $target = $windows | Where-Object { $_.Title -match $escaped } | Select-Object -First 1
    }

    if (-not $target) {
        $processMap = Get-ProcessSnapshot
        $candidatePids = Get-InstanceWindowProcessIds -Instance $Instance -ProcessMap $processMap
        foreach ($candidatePid in $candidatePids) {
            $candidate = $windows |
                Where-Object { $_.ProcessId -eq $candidatePid } |
                Sort-Object `
                    @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.Title)) { 1 } else { 0 } } }, `
                    @{ Expression = { if ($_.Iconic) { 0 } else { 1 } } } |
                Select-Object -First 1
            if ($candidate) {
                $target = $candidate
                break
            }
        }
    }

    if (-not $target -and $Instance.Session -and $Instance.Session.Cwd) {
        $leaf = [regex]::Escape((Split-Path -Leaf $Instance.Session.Cwd))
        $target = $terminalWindows | Where-Object { $_.Title -match $leaf } | Select-Object -First 1
    }

    if (-not $target -and $Instance.AgentType -eq 'Claude') {
        $target = $terminalWindows | Where-Object { $_.Title -match 'CLAUDE|Claude|claudeproject' } | Select-Object -First 1
    }

    if (-not $target -and $Instance.AgentType -eq 'Codex') {
        $target = $terminalWindows | Where-Object { $_.Title -match 'CODEX|Codex' } | Select-Object -First 1
    }

    if (-not $target) {
        $target = $terminalWindows | Select-Object -First 1
    }

    return Resolve-AgentWindowTarget -Window $target -Windows $windows
}

function Restore-AgentWindow {
    param($Instance)
    $target = Find-AgentWindow $Instance
    if (-not $target) { return $false }
    [void][AgentStateNative]::ShowWindowAsync($target.Handle, [AgentStateNative]::SW_RESTORE)
    Start-Sleep -Milliseconds 80
    [void][AgentStateNative]::ShowWindow($target.Handle, [AgentStateNative]::SW_SHOW)
    [void][AgentStateNative]::BringWindowToTop($target.Handle)
    [void][AgentStateNative]::SetWindowPos(
        $target.Handle,
        [IntPtr]::Zero,
        0,
        0,
        0,
        0,
        [AgentStateNative]::SWP_NOMOVE -bor [AgentStateNative]::SWP_NOSIZE -bor [AgentStateNative]::SWP_SHOWWINDOW
    )
    [void][AgentStateNative]::SetForegroundWindow($target.Handle)
    return $true
}

function Request-AgentWindowClose {
    param($Instance)
    $target = Find-AgentWindow $Instance
    if (-not $target) { return $false }
    [void][AgentStateNative]::PostMessage($target.Handle, [AgentStateNative]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
    return $true
}

function Invoke-TaskKill {
    param([int]$ProcessId)
    if ($ProcessId -le 0) { return }
    try {
        $null = & taskkill.exe /PID $ProcessId /T /F 2>$null
    } catch { }
}

function Stop-ProcessTree {
    param(
        [int]$RootProcessId,
        [hashtable]$ProcessMap = $null
    )
    $all = if ($ProcessMap) { $ProcessMap } else { Get-ProcessSnapshot }

    $toStop = New-Object System.Collections.Generic.HashSet[int]
    [void]$toStop.Add($RootProcessId)
    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($proc in $all.Values) {
            if ($toStop.Contains([int]$proc.ParentProcessId) -and -not $toStop.Contains([int]$proc.ProcessId)) {
                [void]$toStop.Add([int]$proc.ProcessId)
                $changed = $true
            }
        }
    }

    $ids = @($toStop) | Sort-Object -Descending
    foreach ($processIdToStop in $ids) {
        try { Stop-Process -Id $processIdToStop -Force } catch { }
    }
}

function Close-AgentInstance {
    param($Instance)
    if (-not $Instance) { return }
    $processMap = Get-ProcessSnapshot
    $rootPid = Get-InstanceControlRootProcessId -Instance $Instance -ProcessMap $processMap
    $windowPids = Get-InstanceWindowProcessIds -Instance $Instance -ProcessMap $processMap
    [void](Request-AgentWindowClose $Instance)
    Start-Sleep -Milliseconds 140
    foreach ($windowProcessId in $windowPids) {
        if ($windowProcessId -ne $rootPid -and $processMap.ContainsKey($windowProcessId)) {
            $name = [string]$processMap[$windowProcessId].Name
            if ($name -in @('cmd.exe', 'powershell.exe', 'pwsh.exe', 'conhost.exe', 'OpenConsole.exe')) {
                Invoke-TaskKill -ProcessId $windowProcessId
            }
        }
    }
    Invoke-TaskKill -ProcessId $rootPid
    Stop-ProcessTree -RootProcessId $rootPid -ProcessMap $processMap
    if ($rootPid -ne [int]$Instance.Process.ProcessId) {
        Invoke-TaskKill -ProcessId ([int]$Instance.Process.ProcessId)
        Stop-ProcessTree -RootProcessId ([int]$Instance.Process.ProcessId)
    }
    $rootStillAlive = Get-Process -Id $rootPid -ErrorAction SilentlyContinue
    if ($rootStillAlive) {
        Start-Sleep -Milliseconds 120
        Invoke-TaskKill -ProcessId $rootPid
        Stop-ProcessTree -RootProcessId $rootPid
    }
    if ($Instance.Launch -and $Instance.Launch.FilePath -and (Test-Path $Instance.Launch.FilePath)) {
        try { Remove-Item -LiteralPath $Instance.Launch.FilePath -Force } catch { }
    }
    Start-Sleep -Milliseconds 250
    Refresh-Instances
}

function Get-DefaultStartDirectory {
    if ($env:AGENTSTATE_DEFAULT_CWD -and (Test-Path $env:AGENTSTATE_DEFAULT_CWD)) {
        return $env:AGENTSTATE_DEFAULT_CWD
    }
    $recent = @($script:LastInstances | Where-Object { $_.AgentType -eq 'Codex' -and $_.Session -and $_.Session.Cwd })
    if ($recent.Count -gt 0) {
        $cwd = [string]$recent[0].Session.Cwd
        if (Test-Path $cwd) { return $cwd }
    }
    if (Test-Path $env:USERPROFILE) { return $env:USERPROFILE }
    $current = (Get-Location).Path
    if ($current -and (Test-Path $current)) {
        return $current
    }
    return (Get-Location).Path
}

function Start-NewCodex {
    if (-not (Test-Path $script:CodexLauncherPath)) { return }
    $cwd = Get-DefaultStartDirectory
    Start-Process -FilePath $script:CodexLauncherPath -WorkingDirectory $cwd | Out-Null
}

function Start-NewClaude {
    if (-not (Test-Path $script:ClaudeLauncherPath)) { return }
    $cwd = Get-DefaultStartDirectory
    Start-Process -FilePath $script:ClaudeLauncherPath -WorkingDirectory $cwd | Out-Null
}

function New-TextBlock {
    param(
        [string]$Text,
        [int]$Size = 12,
        [string]$Color = '#e6edf3',
        [string]$Weight = 'Normal'
    )
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $Text
    $tb.FontFamily = 'Segoe UI'
    $tb.FontSize = $Size
    $tb.Foreground = $Color
    $tb.FontWeight = $Weight
    $tb.TextTrimming = 'CharacterEllipsis'
    $tb.VerticalAlignment = 'Center'
    return $tb
}

function Is-ClickInsideButton {
    param($OriginalSource)
    $current = $OriginalSource
    while ($current) {
        if ($current -is [System.Windows.Controls.Button]) { return $true }
        try {
            $current = [System.Windows.Media.VisualTreeHelper]::GetParent($current)
        } catch {
            $current = $null
        }
    }
    return $false
}

function New-StatusDot {
    param([string]$Brush)
    $ellipse = New-Object System.Windows.Shapes.Ellipse
    $ellipse.Width = 12
    $ellipse.Height = 12
    $ellipse.Fill = $Brush
    $ellipse.Stroke = '#101318'
    $ellipse.StrokeThickness = 1
    $ellipse.Margin = '7,0,10,0'
    $ellipse.VerticalAlignment = 'Center'
    return $ellipse
}

Load-HistoryNames
Load-CodexSessions
Read-CodexLogUpdates

if ($SelfTest) {
    $processes = Get-AgentProcesses
    $instances = Match-Instances $processes
    Write-Output ("Processes: {0}" -f [int](@($processes).Count))
    Write-Output ("Instances: {0}" -f [int](@($instances).Count))
    $rows = foreach ($instance in $instances) {
        $status = Get-InstanceStatus $instance
        $target = Find-AgentWindow $instance
        [pscustomobject]@{
            Type = $instance.AgentType
            Name = Get-DisplayName $instance
            Status = $status.Label
            Detail = $status.Detail
            Pid = $instance.Process.ProcessId
            SessionId = if ($instance.Session) { $instance.Session.SessionId } else { '' }
            ProcessStart = $instance.Process.StartTime
            Window = if ($target) { Shorten-Text ("$($target.ClassName) | $($target.Title)") 64 } else { 'not found' }
        }
    }
    if ($rows.Count -eq 0) {
        Write-Output 'No supported agent instances detected.'
    } else {
        ($rows | Format-Table -AutoSize | Out-String).TrimEnd() | Write-Output
    }
    exit
}

$work = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
if ($work.Height -lt 700) {
    $script:ExpandedMaxLength = [Math]::Max(300, $work.Height - 80)
}
if ($script:Settings.AnchorSide -in @('Left', 'Right', 'Top', 'Bottom')) {
    $script:AnchorSide = $script:Settings.AnchorSide
}
if ($null -ne $script:Settings.AnchorOffset) {
    $script:AnchorOffset = [double]$script:Settings.AnchorOffset
} elseif ($script:AnchorSide -in @('Left', 'Right')) {
    $script:AnchorOffset = [Math]::Max($work.Top + 20, $work.Top + (($work.Height - $script:ExpandedMinLength) / 2))
} else {
    $script:AnchorOffset = [Math]::Max($work.Left + 20, $work.Left + (($work.Width - $script:ExpandedWidth) / 2))
}

$window = New-Object System.Windows.Window
$window.WindowStyle = 'None'
$window.ResizeMode = 'NoResize'
$window.AllowsTransparency = $true
$window.Background = 'Transparent'
$window.Topmost = $true
$window.ShowInTaskbar = $false

$root = New-Object System.Windows.Controls.Border
$root.Background = '#E6191D24'
$root.BorderBrush = '#404956'
$root.BorderThickness = '1'
$root.CornerRadius = '16'
$root.Padding = '0'

$dock = New-Object System.Windows.Controls.DockPanel
$root.Child = $dock

$header = New-Object System.Windows.Controls.Grid
$header.Height = 44
$header.Margin = '0,6,0,0'
$header.Cursor = 'SizeAll'
$header.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '28' }))
$header.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
$header.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '0' }))
[System.Windows.Controls.DockPanel]::SetDock($header, 'Top')
$dock.Children.Add($header) | Out-Null

$headerDot = New-StatusDot '#58a6ff'
[System.Windows.Controls.Grid]::SetColumn($headerDot, 0)
$header.Children.Add($headerDot) | Out-Null

$title = New-TextBlock 'AgentState' 14 '#f0f6fc' 'SemiBold'
[System.Windows.Controls.Grid]::SetColumn($title, 1)
$header.Children.Add($title) | Out-Null

$footer = New-Object System.Windows.Controls.Grid
$footer.Height = 42
$footer.Margin = '8,0,8,8'
$footer.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '66' }))
$footer.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '66' }))
$footer.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
$footer.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '42' }))
$footer.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '42' }))
[System.Windows.Controls.DockPanel]::SetDock($footer, 'Bottom')
$dock.Children.Add($footer) | Out-Null

$configButton = New-Object System.Windows.Controls.Button
$configButton.Content = '配置'
$configButton.Width = 54
$configButton.Height = 26
$configButton.Background = '#2c3441'
$configButton.Foreground = '#c7d0db'
$configButton.BorderBrush = '#4a5564'
$configButton.ToolTip = '打开配置窗口'
[System.Windows.Controls.Grid]::SetColumn($configButton, 0)
$footer.Children.Add($configButton) | Out-Null

$exitButton = New-Object System.Windows.Controls.Button
$exitButton.Content = '退出'
$exitButton.Width = 54
$exitButton.Height = 26
$exitButton.Background = '#2c3441'
$exitButton.Foreground = '#c7d0db'
$exitButton.BorderBrush = '#4a5564'
$exitButton.ToolTip = '退出 AgentState'
$exitButton.Add_Click({
    param($sender, $eventArgs)
    $eventArgs.Handled = $true
    $window.Close()
})
[System.Windows.Controls.Grid]::SetColumn($exitButton, 1)
$footer.Children.Add($exitButton) | Out-Null

$claudePlusButton = New-Object System.Windows.Controls.Button
$claudePlusButton.Content = '+'
$claudePlusButton.Width = 32
$claudePlusButton.Height = 32
$claudePlusButton.FontSize = 20
$claudePlusButton.FontWeight = 'Bold'
$claudePlusButton.Background = '#f0883e'
$claudePlusButton.Foreground = '#ffffff'
$claudePlusButton.BorderBrush = '#ffb86b'
$claudePlusButton.ToolTip = 'Start new Claude terminal'
$claudePlusButton.Add_Click({
    param($sender, $eventArgs)
    $eventArgs.Handled = $true
    Start-NewClaude
})
[System.Windows.Controls.Grid]::SetColumn($claudePlusButton, 3)
$footer.Children.Add($claudePlusButton) | Out-Null

$plusButton = New-Object System.Windows.Controls.Button
$plusButton.Content = '+'
$plusButton.Width = 32
$plusButton.Height = 32
$plusButton.FontSize = 20
$plusButton.FontWeight = 'Bold'
$plusButton.Background = '#1f6feb'
$plusButton.Foreground = '#ffffff'
$plusButton.BorderBrush = '#58a6ff'
$plusButton.ToolTip = 'Start new Codex terminal'
$plusButton.Add_Click({
    param($sender, $eventArgs)
    $eventArgs.Handled = $true
    Start-NewCodex
})
[System.Windows.Controls.Grid]::SetColumn($plusButton, 4)
$footer.Children.Add($plusButton) | Out-Null

$itemsPanel = New-Object System.Windows.Controls.StackPanel
$itemsPanel.Margin = '0,4,8,8'
$expandedScroll = New-Object System.Windows.Controls.ScrollViewer
$expandedScroll.VerticalScrollBarVisibility = 'Auto'
$expandedScroll.HorizontalScrollBarVisibility = 'Disabled'
$expandedScroll.Content = $itemsPanel
$dock.Children.Add($expandedScroll) | Out-Null

$collapsedPanel = New-Object System.Windows.Controls.StackPanel
$collapsedPanel.HorizontalAlignment = 'Center'
$collapsedPanel.VerticalAlignment = 'Center'

$collapsedScroll = New-Object System.Windows.Controls.ScrollViewer
$collapsedScroll.Content = $collapsedPanel
$collapsedScroll.HorizontalScrollBarVisibility = 'Hidden'
$collapsedScroll.VerticalScrollBarVisibility = 'Hidden'
$collapsedScroll.CanContentScroll = $false
$collapsedScroll.Focusable = $false
$collapsedScroll.Background = 'Transparent'
$collapsedScroll.BorderThickness = '0'
$collapsedScroll.HorizontalContentAlignment = 'Center'
$collapsedScroll.VerticalContentAlignment = 'Center'
$collapsedScroll.Visibility = 'Collapsed'
$dock.Children.Add($collapsedScroll) | Out-Null

$window.Content = $root

function New-CollapsedMarker {
    param(
        [string]$Brush,
        [string]$Tooltip
    )
    $marker = New-Object System.Windows.Controls.Border
    if ($script:AnchorSide -in @('Top', 'Bottom')) {
        $marker.Width = 12
        $marker.Height = 12
        $marker.Margin = '4,0,4,0'
    } else {
        $marker.Width = 12
        $marker.Height = 12
        $marker.Margin = '0,4,0,4'
    }
    $marker.CornerRadius = '6'
    $marker.Background = $Brush
    $marker.BorderBrush = '#0b0e12'
    $marker.BorderThickness = '1.2'
    $marker.ToolTip = $Tooltip
    return $marker
}

function Update-CollapsedLayout {
    if ($script:AnchorSide -in @('Top', 'Bottom')) {
        $collapsedPanel.Orientation = 'Horizontal'
        $collapsedPanel.Margin = '10,0,10,0'
        $collapsedScroll.HorizontalScrollBarVisibility = 'Hidden'
        $collapsedScroll.VerticalScrollBarVisibility = 'Hidden'
    } else {
        $collapsedPanel.Orientation = 'Vertical'
        $collapsedPanel.Margin = '0,10,0,10'
        $collapsedScroll.HorizontalScrollBarVisibility = 'Hidden'
        $collapsedScroll.VerticalScrollBarVisibility = 'Hidden'
    }
}

function Show-ConfigDialog {
    if ($script:IsConfigDialogOpen) { return }

    $script:IsConfigDialogOpen = $true
    try {
        $dialog = New-Object System.Windows.Window
        $dialog.Title = 'AgentState 配置'
        $dialog.Owner = $window
        $dialog.WindowStartupLocation = 'CenterScreen'
        $dialog.ResizeMode = 'NoResize'
        $dialog.SizeToContent = 'WidthAndHeight'
        $dialog.WindowStyle = 'ToolWindow'
        $dialog.Background = '#141a21'
        $dialog.Topmost = $true
        $dialog.ShowInTaskbar = $false

        $dialogBorder = New-Object System.Windows.Controls.Border
        $dialogBorder.Background = '#1c232d'
        $dialogBorder.BorderBrush = '#3a4655'
        $dialogBorder.BorderThickness = '1'
        $dialogBorder.CornerRadius = '10'
        $dialogBorder.Padding = '14'

        $dialogStack = New-Object System.Windows.Controls.StackPanel
        $dialogBorder.Child = $dialogStack

        $dialogTitle = New-TextBlock '配置' 14 '#f0f6fc' 'SemiBold'
        $dialogStack.Children.Add($dialogTitle) | Out-Null

        $dialogHint = New-TextBlock '展开方式' 11 '#9da7b3' 'SemiBold'
        $dialogHint.Margin = '0,8,0,4'
        $dialogStack.Children.Add($dialogHint) | Out-Null

        $radioHover = New-Object System.Windows.Controls.RadioButton
        $radioHover.Content = '鼠标移动到侧边时展开'
        $radioHover.Foreground = '#e6edf3'
        $radioHover.GroupName = 'expand_mode'
        $radioHover.Margin = '0,2,0,2'
        $radioHover.IsChecked = $script:Settings.ExpandMode -eq 'Hover'
        $dialogStack.Children.Add($radioHover) | Out-Null

        $radioClick = New-Object System.Windows.Controls.RadioButton
        $radioClick.Content = '点击折叠栏时展开'
        $radioClick.Foreground = '#e6edf3'
        $radioClick.GroupName = 'expand_mode'
        $radioClick.Margin = '0,2,0,2'
        $radioClick.IsChecked = $script:Settings.ExpandMode -eq 'Click'
        $dialogStack.Children.Add($radioClick) | Out-Null

        $dialogFooter = New-Object System.Windows.Controls.Grid
        $dialogFooter.Margin = '0,12,0,0'
        $dialogFooter.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
        $dialogFooter.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '62' }))
        $dialogStack.Children.Add($dialogFooter) | Out-Null

        $closeDialogButton = New-Object System.Windows.Controls.Button
        $closeDialogButton.Content = '关闭'
        $closeDialogButton.Width = 56
        $closeDialogButton.Height = 26
        $closeDialogButton.Background = '#2c3441'
        $closeDialogButton.Foreground = '#c7d0db'
        $closeDialogButton.BorderBrush = '#4a5564'
        $closeDialogButton.Add_Click({
            param($sender, $eventArgs)
            $eventArgs.Handled = $true
            $dialog.Close()
        })
        [System.Windows.Controls.Grid]::SetColumn($closeDialogButton, 1)
        $dialogFooter.Children.Add($closeDialogButton) | Out-Null

        $radioHover.Add_Checked({ Set-ExpandMode 'Hover' })
        $radioClick.Add_Checked({ Set-ExpandMode 'Click' })

        $dialog.Content = $dialogBorder
        $dialog.ShowDialog() | Out-Null
    } finally {
        $script:IsConfigDialogOpen = $false
    }
}

function Get-VisibleSize {
    param([bool]$Expanded)
    $script:CurrentCollapsedLength = Get-CollapsedLength
    $script:CurrentExpandedLength = Get-ExpandedLength

    if ($script:AnchorSide -in @('Left', 'Right')) {
        return [pscustomobject]@{
            Width = if ($Expanded) { $script:ExpandedWidth } else { $script:CollapsedThickness }
            Height = if ($Expanded) { $script:CurrentExpandedLength } else { $script:CurrentCollapsedLength }
        }
    }

    return [pscustomobject]@{
        Width = if ($Expanded) { $script:ExpandedWidth } else { $script:CurrentCollapsedLength }
        Height = if ($Expanded) { $script:CurrentExpandedLength } else { $script:CollapsedThickness }
    }
}

function Set-WindowBounds {
    $work = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $size = Get-VisibleSize $script:IsExpanded
    $window.Width = $size.Width
    $window.Height = $size.Height
    $window.Topmost = $true

    switch ($script:AnchorSide) {
        'Left' {
            $window.Left = $work.Left
            $window.Top = Clamp-Double $script:AnchorOffset $work.Top ($work.Bottom - $size.Height)
        }
        'Right' {
            $window.Left = $work.Right - $size.Width
            $window.Top = Clamp-Double $script:AnchorOffset $work.Top ($work.Bottom - $size.Height)
        }
        'Top' {
            $window.Left = Clamp-Double $script:AnchorOffset $work.Left ($work.Right - $size.Width)
            $window.Top = $work.Top
        }
        'Bottom' {
            $window.Left = Clamp-Double $script:AnchorOffset $work.Left ($work.Right - $size.Width)
            $window.Top = $work.Bottom - $size.Height
        }
    }
}

function Set-ExpandMode {
    param([string]$Mode)
    if ($Mode -notin @('Hover', 'Click')) { return }
    $script:Settings.ExpandMode = $Mode
    Save-Settings
    if (Is-HoverExpandMode) {
        if ($window.IsMouseOver) {
            Set-Expanded $true
        } else {
            Set-Expanded $false
        }
    } else {
        Set-Expanded $false
    }
}

function Set-Expanded {
    param([bool]$Expanded)
    if ($script:IsDragging) { return }
    $script:IsExpanded = $Expanded
    Update-CollapsedLayout
    $header.Visibility = if ($Expanded) { 'Visible' } else { 'Collapsed' }
    $footer.Visibility = if ($Expanded) { 'Visible' } else { 'Collapsed' }
    $expandedScroll.Visibility = if ($Expanded) { 'Visible' } else { 'Collapsed' }
    $collapsedScroll.Visibility = if ($Expanded) { 'Collapsed' } else { 'Visible' }
    Set-WindowBounds
}

function Snap-WindowToNearestEdge {
    $work = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $distances = @{
        Left = [Math]::Abs($window.Left - $work.Left)
        Right = [Math]::Abs($work.Right - ($window.Left + $window.Width))
        Top = [Math]::Abs($window.Top - $work.Top)
        Bottom = [Math]::Abs($work.Bottom - ($window.Top + $window.Height))
    }
    $script:AnchorSide = ($distances.GetEnumerator() | Sort-Object Value | Select-Object -First 1).Key
    $expandedSize = Get-VisibleSize $true
    if ($script:AnchorSide -in @('Left', 'Right')) {
        $script:AnchorOffset = Clamp-Double $window.Top $work.Top ($work.Bottom - $expandedSize.Height)
    } else {
        $script:AnchorOffset = Clamp-Double $window.Left $work.Left ($work.Right - $expandedSize.Width)
    }
    Save-Settings
    Set-WindowBounds
}

function Refresh-CollapsedIndicators {
    Update-CollapsedLayout
    $collapsedPanel.Children.Clear()

    if ($script:LastInstances.Count -eq 0) {
        $collapsedPanel.Children.Add((New-CollapsedMarker '#586270' 'no active agents')) | Out-Null
        return
    }

    foreach ($instance in $script:LastInstances) {
        $status = Get-InstanceStatus $instance
        $tooltip = "$($status.Label) | $(Get-DisplayName $instance)"
        $collapsedPanel.Children.Add((New-CollapsedMarker $status.Brush $tooltip)) | Out-Null
    }

    if ($script:AnchorSide -in @('Top', 'Bottom')) {
        $collapsedScroll.ScrollToLeftEnd()
    } else {
        $collapsedScroll.ScrollToTop()
    }
}

function New-InstanceCard {
    param($Instance)
    $status = Get-InstanceStatus $Instance
    $name = Get-DisplayName $Instance
    $sessionId = if ($Instance.Session) { $Instance.Session.SessionId } else { '' }
    $cwd = if ($Instance.Session -and $Instance.Session.Cwd) { $Instance.Session.Cwd } elseif ($Instance.Launch) { $Instance.Launch.WorkingDirectory } else { '' }
    $started = $Instance.Process.StartTime.ToString('HH:mm:ss')

    $card = New-Object System.Windows.Controls.Border
    $card.CornerRadius = '10'
    $card.Margin = '0,0,0,8'
    $card.Padding = '0'
    $card.MinHeight = 56
    $card.Background = '#232a33'
    $card.BorderBrush = '#303946'
    $card.BorderThickness = '1'
    $card.Cursor = 'Hand'
    $card.Tag = $Instance

    $grid = New-Object System.Windows.Controls.Grid
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '28' }))
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
    $grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '38' }))
    $card.Child = $grid

    $dot = New-StatusDot $status.Brush
    [System.Windows.Controls.Grid]::SetColumn($dot, 0)
    $grid.Children.Add($dot) | Out-Null

    $textPanel = New-Object System.Windows.Controls.StackPanel
    $textPanel.Margin = '0,7,4,7'
    [System.Windows.Controls.Grid]::SetColumn($textPanel, 1)
    $grid.Children.Add($textPanel) | Out-Null

    $nameText = New-TextBlock $name 13 '#f0f6fc' 'SemiBold'
    $textPanel.Children.Add($nameText) | Out-Null

    $detail = "$($Instance.AgentType) | $($status.Label) | pid $($Instance.Process.ProcessId) | $started"
    if ($cwd) { $detail = "$detail | $(Split-Path -Leaf $cwd)" }
    $detailText = New-TextBlock $detail 11 '#9da7b3'
    $textPanel.Children.Add($detailText) | Out-Null

    if ($status.Detail) {
        $eventText = New-TextBlock (Shorten-Text $status.Detail 84) 10 '#77818f'
        $textPanel.Children.Add($eventText) | Out-Null
    }

    $closeItem = New-Object System.Windows.Controls.Button
    $closeItem.Content = 'x'
    $closeItem.Width = 24
    $closeItem.Height = 24
    $closeItem.Margin = '0,0,7,0'
    $closeItem.Background = '#402630'
    $closeItem.Foreground = '#ffb3bd'
    $closeItem.BorderBrush = '#7d3442'
    $closeItem.ToolTip = 'Close this agent terminal'
    $closeItem.Tag = $Instance
    $closeItem.ClickMode = 'Press'
    $closeItem.Add_Click({
        param($sender, $eventArgs)
        $eventArgs.Handled = $true
        Close-AgentInstance $sender.Tag
    })
    [System.Windows.Controls.Grid]::SetColumn($closeItem, 2)
    $grid.Children.Add($closeItem) | Out-Null

    $tip = "Type: $($Instance.AgentType)`nStatus: $($status.Label)`nPID: $($Instance.Process.ProcessId)"
    if ($sessionId) { $tip += "`nSession: $sessionId" }
    if ($cwd) { $tip += "`nCWD: $cwd" }
    if ($Instance.Launch -and $Instance.Launch.Title) { $tip += "`nWindow: $($Instance.Launch.Title)" }
    $card.ToolTip = $tip

    $card.Add_MouseLeftButtonDown({
        param($sender, $eventArgs)
        if ($eventArgs.Handled -or (Is-ClickInsideButton $eventArgs.OriginalSource)) { return }
        $eventArgs.Handled = $true
        [void](Restore-AgentWindow $sender.Tag)
    })

    return $card
}

function Refresh-Instances {
    Load-HistoryNames
    Read-CodexLogUpdates
    $processes = Get-AgentProcesses
    $script:LastInstances = Match-Instances $processes

    $itemsPanel.Children.Clear()
    if ($script:LastInstances.Count -eq 0) {
        $empty = New-Object System.Windows.Controls.Border
        $empty.Margin = '0,0,0,8'
        $empty.Padding = '12,10,12,10'
        $empty.MinHeight = 52
        $empty.Background = '#232a33'
        $empty.BorderBrush = '#303946'
        $empty.BorderThickness = '1'
        $empty.CornerRadius = '10'
        $empty.Child = New-TextBlock 'No active agents' 12 '#9da7b3' 'SemiBold'
        $itemsPanel.Children.Add($empty) | Out-Null
    } else {
        foreach ($instance in $script:LastInstances) {
            $itemsPanel.Children.Add((New-InstanceCard $instance)) | Out-Null
        }
    }

    Refresh-CollapsedIndicators
    Set-WindowBounds
}

$configButton.Add_Click({
    param($sender, $eventArgs)
    $eventArgs.Handled = $true
    Show-ConfigDialog
})

$header.Add_MouseLeftButtonDown({
    param($sender, $eventArgs)
    if ($eventArgs.ClickCount -ne 1) { return }
    if (Is-ClickInsideButton $eventArgs.OriginalSource) { return }
    $script:IsDragging = $true
    Set-Expanded $true
    try { $window.DragMove() } catch { }
    Snap-WindowToNearestEdge
    $script:IsDragging = $false
    Set-Expanded $true
})

$collapsedScroll.Add_MouseLeftButtonDown({
    param($sender, $eventArgs)
    if ($script:IsExpanded -or (Is-HoverExpandMode)) { return }
    $eventArgs.Handled = $true
    $script:IsCollapsedPointerDown = $true
    $script:IsCollapsedDragging = $false
    $script:CollapsedPressPoint = $eventArgs.GetPosition($window)
})

$collapsedScroll.Add_MouseMove({
    param($sender, $eventArgs)
    if ($script:IsExpanded -or (Is-HoverExpandMode)) { return }
    if (-not $script:IsCollapsedPointerDown -or $script:IsCollapsedDragging) { return }
    if ($eventArgs.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { return }

    $currentPoint = $eventArgs.GetPosition($window)
    $dx = [Math]::Abs($currentPoint.X - $script:CollapsedPressPoint.X)
    $dy = [Math]::Abs($currentPoint.Y - $script:CollapsedPressPoint.Y)
    if (($dx + $dy) -lt 6) { return }

    $script:IsCollapsedDragging = $true
    $script:IsDragging = $true
    try { $window.DragMove() } catch { }
    Snap-WindowToNearestEdge
    $script:IsDragging = $false
    Set-Expanded $false
})

$collapsedScroll.Add_MouseLeftButtonUp({
    param($sender, $eventArgs)
    if (-not $script:IsCollapsedPointerDown) { return }
    $eventArgs.Handled = $true
    $shouldExpand = -not $script:IsCollapsedDragging
    $script:IsCollapsedPointerDown = $false
    $script:IsCollapsedDragging = $false
    $script:CollapsedPressPoint = $null

    if ($shouldExpand) {
        Set-Expanded $true
        try { $window.Activate() | Out-Null } catch { }
    }
})

$collapsedScroll.Add_PreviewMouseWheel({
    param($sender, $eventArgs)
    if ($script:AnchorSide -in @('Top', 'Bottom')) {
        $collapsedScroll.ScrollToHorizontalOffset([Math]::Max(0, $collapsedScroll.HorizontalOffset - ($eventArgs.Delta / 3)))
    } else {
        $collapsedScroll.ScrollToVerticalOffset([Math]::Max(0, $collapsedScroll.VerticalOffset - ($eventArgs.Delta / 3)))
    }
    $eventArgs.Handled = $true
})

$window.Add_MouseEnter({
    if (Is-HoverExpandMode) {
        Set-Expanded $true
    }
})

$window.Add_MouseLeave({
    if (Is-HoverExpandMode -and -not $script:IsConfigDialogOpen) {
        Set-Expanded $false
    }
})

$window.Add_Deactivated({
    if (-not (Is-HoverExpandMode) -and $script:IsExpanded -and -not $script:IsConfigDialogOpen) {
        Set-Expanded $false
    }
})

$window.Add_Activated({ $window.Topmost = $true })
$window.Add_ContentRendered({ $window.Topmost = $true })

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(3)
$timer.Add_Tick({ Refresh-Instances })
$timer.Start()

Set-Expanded $false
Refresh-Instances
$window.ShowDialog() | Out-Null
