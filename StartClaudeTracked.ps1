$ErrorActionPreference = 'SilentlyContinue'

$stateHome = Join-Path $env:USERPROFILE '.agent-state'
$launchRoot = Join-Path $stateHome 'launches'
if (-not (Test-Path $launchRoot)) {
    New-Item -ItemType Directory -Path $launchRoot -Force | Out-Null
}

function Convert-WmiDateToIsoString {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return ([datetimeoffset]$Value).ToString('o') }
    if ($Value -is [datetimeoffset]) { return $Value.ToString('o') }
    try {
        return ([datetimeoffset]([Management.ManagementDateTimeConverter]::ToDateTime([string]$Value))).ToString('o')
    } catch {
        return $null
    }
}

$claude = Get-Command claude -ErrorAction SilentlyContinue
if (-not $claude) {
    $fallback = Join-Path $env:USERPROFILE '.local\bin\claude.exe'
    if (Test-Path $fallback) {
        $claude = [pscustomobject]@{ Source = $fallback }
    }
}
if (-not $claude) {
    Write-Host 'Claude executable not found.'
    exit 1
}

$startedAt = Get-Date
$instanceId = '{0}-{1}' -f $startedAt.ToString('yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8))
$title = "CLAUDE $instanceId"
$launchPath = Join-Path $launchRoot "$instanceId.json"
$parentPid = 0
$launcherProc = $null
$parentProc = $null
try {
    $launcherProc = Get-CimInstance Win32_Process -Filter "ProcessId=$PID"
    $parentPid = [int]$launcherProc.ParentProcessId
    if ($parentPid -gt 0) {
        $parentProc = Get-CimInstance Win32_Process -Filter "ProcessId=$parentPid"
    }
} catch { }

try { $Host.UI.RawUI.WindowTitle = $title } catch { }

[pscustomobject]@{
    AgentType = 'Claude'
    InstanceId = $instanceId
    Title = $title
    StartedAt = ([datetimeoffset]$startedAt).ToString('o')
    WorkingDirectory = (Get-Location).Path
    LauncherProcessId = $PID
    ParentProcessId = $parentPid
    RootShellProcessId = $parentPid
    LauncherProcessName = if ($launcherProc) { [string]$launcherProc.Name } else { 'powershell.exe' }
    ParentProcessName = if ($parentProc) { [string]$parentProc.Name } else { '' }
    RootShellProcessName = if ($parentProc) { [string]$parentProc.Name } else { '' }
    LauncherProcessStartedAt = if ($launcherProc) { Convert-WmiDateToIsoString $launcherProc.CreationDate } else { ([datetimeoffset]$startedAt).ToString('o') }
    ParentProcessStartedAt = if ($parentProc) { Convert-WmiDateToIsoString $parentProc.CreationDate } else { $null }
    RootShellProcessStartedAt = if ($parentProc) { Convert-WmiDateToIsoString $parentProc.CreationDate } else { $null }
} | ConvertTo-Json -Depth 3 | Set-Content -Path $launchPath -Encoding UTF8

& $claude.Source --dangerously-skip-permissions @args
$exitCode = $LASTEXITCODE

try {
    if (Test-Path $launchPath) {
        Remove-Item -LiteralPath $launchPath -Force
    }
} catch { }

exit $exitCode
