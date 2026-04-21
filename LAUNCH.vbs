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
    MsgBox "Python не е намерен в WinPython папката!" & vbCrLf & _
           "Провери дали WinPython е инсталиран в: " & sRoot, _
           vbCritical, "Data Extraction Bot"
    WScript.Quit 1
End If

' ── Start uvicorn hidden (window style 0 = no CMD window) ────────────────────
sCmd = """" & sPython & """ -m uvicorn app.main:app " & _
       "--host 127.0.0.1 --port 5000 " & _
       "--app-dir """ & sRoot & """"

oShell.CurrentDirectory = sRoot
oShell.Run sCmd, 0, False   ' 0 = hidden, False = don't wait

' ── Wait for server to start, then open browser ──────────────────────────────
WScript.Sleep 3000
oShell.Run "http://localhost:5000"
