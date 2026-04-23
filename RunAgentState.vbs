Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(WScript.ScriptFullName)
cwd = shell.CurrentDirectory
If Len(cwd) > 0 Then
  shell.Environment("PROCESS")("AGENTSTATE_DEFAULT_CWD") = cwd
End If
shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & base & "\AgentStateBar.ps1""", 0, False
