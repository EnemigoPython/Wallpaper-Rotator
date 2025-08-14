' Direct Wallpaper Rotator - Simplest Silent Runner
Dim objShell, scriptDir, pythonScript

Set objShell = CreateObject("WScript.Shell")
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
pythonScript = scriptDir & "\wallpaper_rotator.py"

' Run Python script completely hidden
objShell.Run "pythonw.exe """ & pythonScript & """ --quiet", 0, True

Set objShell = Nothing