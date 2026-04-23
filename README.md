# AgentState

AgentState is a lightweight Windows status bar for tracked Codex and Claude terminals.

## Files

- `AgentStateBar.ps1`: main UI and process tracking logic
- `RunAgentState.bat` / `RunAgentState.vbs`: launchers for the status bar
- `StartCodexTracked.bat` / `StartCodexTracked.ps1`: tracked Codex launcher
- `StartClaudeTracked.bat` / `StartClaudeTracked.ps1`: tracked Claude launcher
- `assets/AgentState.ico`: tray/window icon

## Requirements

- Windows PowerShell 5.1 or newer
- Windows Terminal
- `codex` available in `PATH` for Codex tracking
- `claude` available in `PATH` for Claude tracking

## Run

Start the bar:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\AgentStateBar.ps1
```

Or use:

```powershell
.\RunAgentState.bat
```

Start a tracked Codex terminal:

```powershell
.\StartCodexTracked.bat
```

Start a tracked Claude terminal:

```powershell
.\StartClaudeTracked.bat
```

## Notes

- Clicking an agent card restores the corresponding terminal window.
- The bar tracks active work, waiting state, startup state, and error state.
- Launch metadata is stored under `%USERPROFILE%\.agent-state`.
