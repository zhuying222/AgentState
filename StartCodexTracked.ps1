$ErrorActionPreference = 'SilentlyContinue'

$stateHome = Join-Path $env:USERPROFILE '.agent-state'
$launchRoot = Join-Path $stateHome 'launches'
if (-not (Test-Path $launchRoot)) {
    New-Item -ItemType Directory -Path $launchRoot -Force | Out-Null
}

$startedAt = Get-Date
$instanceId = '{0}-{1}' -f $startedAt.ToString('yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0, 8))
$title = "CODEX $instanceId"
$launchPath = Join-Path $launchRoot "$instanceId.json"
$parentPid = 0
try {
    $parentPid = [int](Get-CimInstance Win32_Process -Filter "ProcessId=$PID").ParentProcessId
} catch { }

try { $Host.UI.RawUI.WindowTitle = $title } catch { }

[pscustomobject]@{
    AgentType = 'Codex'
    InstanceId = $instanceId
    Title = $title
    StartedAt = ([datetimeoffset]$startedAt).ToString('o')
    WorkingDirectory = (Get-Location).Path
    LauncherProcessId = $PID
    ParentProcessId = $parentPid
    RootShellProcessId = $parentPid
} | ConvertTo-Json -Depth 3 | Set-Content -Path $launchPath -Encoding UTF8

& codex
$exitCode = $LASTEXITCODE

try {
    if (Test-Path $launchPath) {
        Remove-Item -LiteralPath $launchPath -Force
    }
} catch { }

exit $exitCode
