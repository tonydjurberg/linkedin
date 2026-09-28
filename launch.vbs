Set shell = CreateObject("WScript.Shell")
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
exePath = scriptDir & "\ProspectHunter.exe"
shell.Run Chr(34) & exePath & Chr(34), 0, False
