Option Explicit
Dim oShell, oFSO, sRoot, sDesktop, oShortcut

Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

sRoot    = oFSO.GetParentFolderName(WScript.ScriptFullName)
sDesktop = oShell.SpecialFolders("Desktop")

Set oShortcut = oShell.CreateShortcut(sDesktop & "\Data Extraction Bot.lnk")
oShortcut.TargetPath       = "wscript.exe"
oShortcut.Arguments        = """" & sRoot & "\LAUNCH.vbs"""
oShortcut.WorkingDirectory = sRoot
oShortcut.IconLocation     = "shell32.dll,14"
oShortcut.Description      = "Data Extraction Bot"
oShortcut.Save

MsgBox "Shortcut created on Desktop!", vbInformation, "Data Extraction Bot"
