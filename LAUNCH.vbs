Option Explicit
Dim oShell, oFSO, sRoot, sPython, sCmd

Set oShell  = CreateObject("WScript.Shell")
Set oFSO    = CreateObject("Scripting.FileSystemObject")

' Directory where this .vbs file lives
sRoot = oFSO.GetParentFolderName(WScript.ScriptFullName)

' ── Find python.exe inside WinPython subfolders ──────────────────────────────
sPython = ""
Dim oWinPy, oPyDir

If oFSO.FolderExists(sRoot & "\WinPython") Then
    For Each oWinPy In oFSO.GetFolder(sRoot & "\WinPython").SubFolders
        ' Level 1: WinPython\python-X.Y.Z\
        If Left(oWinPy.Name, 7) = "python-" And oFSO.FileExists(oWinPy.Path & "\python.exe") Then
            sPython = oWinPy.Path & "\python.exe"
        End If
        ' Level 2: WinPython\SomeFolder\python-X.Y.Z\
        For Each oPyDir In oWinPy.SubFolders
            If Left(oPyDir.Name, 7) = "python-" And oFSO.FileExists(oPyDir.Path & "\python.exe") Then
                sPython = oPyDir.Path & "\python.exe"
            End If
        Next
    Next
End If

If sPython = "" Then
    MsgBox "Python not found in WinPython folder!" & vbCrLf & _
           "Check if WinPython is installed in: " & sRoot, _
           vbCritical, "Data Extraction Bot"
    WScript.Quit 1
End If

' ── Start bot — run_bot.py finds a free port automatically ──────────────────
sCmd = """" & sPython & """ """ & sRoot & "\run_bot.py"""
oShell.CurrentDirectory = sRoot
oShell.Run sCmd, 0, False   ' 0 = hidden window, browser opens automatically
