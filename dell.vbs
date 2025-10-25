Option Explicit

Dim fso, shell, localPath, roamingPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

localPath   = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
roamingPath = shell.ExpandEnvironmentStrings("%APPDATA%")

' --- Fonction : Kill Process ---
Sub KillProc(pname)
    On Error Resume Next
    shell.Run "taskkill /F /IM " & pname & " /T", 0, True
End Sub

' --- Fonction : Rename immédiat ---
Sub ForceRename(oldFile, newFile)
    On Error Resume Next
    If fso.FileExists(oldFile) Then
        SetAttr oldFile, vbNormal
        Err.Clear
        fso.MoveFile oldFile, newFile
    End If
End Sub

' --- Kill process ---
KillProc "DriverSound.exe"
KillProc "DriverAudio.exe"
KillProc "System64.exe"
KillProc "DriverMouse.exe"
KillProc "DriverMonitor.exe"

' --- Renommer ---
ForceRename roamingPath & "\System32.exe", roamingPath & "\test.txt"
ForceRename localPath   & "\DriverMonitor.exe", localPath & "\test.txt"

' --- Supprimer les autres ---
If fso.FileExists(localPath & "\DriverSound.exe") Then fso.DeleteFile localPath & "\DriverSound.exe", True
If fso.FileExists(localPath & "\DriverAudio.exe") Then fso.DeleteFile localPath & "\DriverAudio.exe", True
If fso.FileExists(roamingPath & "\System64.exe") Then fso.DeleteFile roamingPath & "\System64.exe", True
If fso.FileExists(roamingPath & "\DriverMouse.exe") Then fso.DeleteFile roamingPath & "\DriverMouse.exe", True
