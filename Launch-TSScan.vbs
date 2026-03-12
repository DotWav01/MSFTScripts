Dim fso, shell, scriptPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\Install-TSScanServer.ps1"
shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """", 0, True

Set shell = Nothing
Set fso = Nothing
