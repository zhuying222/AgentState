@echo off
if not defined AGENTSTATE_DEFAULT_CWD set "AGENTSTATE_DEFAULT_CWD=%CD%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0AgentStateBar.ps1"
