$ErrorActionPreference = 'SilentlyContinue'

$stateHome = Join-Path $env:USERPROFILE '.agent-state'
$launchRoot = Join-Path $stateHome 'launches'
if (-not (Test-Path $launchRoot)) {
    New-Item -ItemType Directory -Path $launchRoot -Force | Out-Null
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
try {
    $parentPid = [int](Get-CimInstance Win32_Process -Filter "ProcessId=$PID").ParentProcessId
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
} | ConvertTo-Json -Depth 3 | Set-Content -Path $launchPath -Encoding UTF8

& $claude.Source --dangerously-skip-permissions @args
$exitCode = $LASTEXITCODE

try {
    if (Test-Path $launchPath) {
        Remove-Item -LiteralPath $launchPath -Force
    }
} catch { }

exit $exitCode
