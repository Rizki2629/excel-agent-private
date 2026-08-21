Option Explicit

Dim shell, fso, projectDir, command
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

projectDir = fso.GetParentFolderName(WScript.ScriptFullName)
command = "cmd /c cd /d """ & projectDir & """ && node server.js >> server.out.log 2>> server.err.log"

' WindowStyle 0 = hidden; False = do not wait.
shell.Run command, 0, False
