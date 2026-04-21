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
oShortcut.IconLocation     = "shell32.dll,14"   ' globe/browser icon
oShortcut.Description      = "Data Extraction Bot"
oShortcut.Save

MsgBox "Иконката е създадена на работния плот!", vbInformation, "Data Extraction Bot"
