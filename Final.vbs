Option Explicit

Dim shell, fso
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")

' --- Cibles à nettoyer ---
Dim pdAudioHelper, appAudioHelper, appBackup, taskName
pdAudioHelper  = shell.ExpandEnvironmentStrings("%ProgramData%") & "\AudioHelper"
appAudioHelper = shell.ExpandEnvironmentStrings("%APPDATA%")   & "\AudioHelper"
appBackup      = shell.ExpandEnvironmentStrings("%APPDATA%")   & "\AudioBackup"
taskName       = "\Microsoft\Windows\Update\WinUpdateSvc"

' --- UAC: s'auto-élève si nécessaire ---
If Not IsElevated() Then
  CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ //nologo", "", "runas", 0
  WScript.Quit
End If

' --- Tuer les processus suspects ---
KillIfRunning "RtkAudUService.exe"
KillIfRunning "RtkAudUService64.exe"

' --- Supprimer la tâche planifiée si elle existe ---
If TaskExists(taskName) Then
  RunCmd "schtasks /delete /tn """ & taskName & """ /f", True
End If

' --- Réinitialiser ACL + attributs puis supprimer les dossiers ---
CleanFolder pdAudioHelper
CleanFolder appAudioHelper
CleanFolder appBackup

WScript.Quit

' ======================= Fonctions utilitaires =======================

Function RunCmd(cmd, wait)
  RunCmd = shell.Run("cmd /c " & cmd & " >nul 2>&1", 0, wait)
End Function

Function IsElevated()
  Dim rc : rc = shell.Run("cmd /c ""net session >nul 2>&1""", 0, True)
  IsElevated = (rc = 0)
End Function

Sub KillIfRunning(procName)
  RunCmd "taskkill /f /im " & procName, True
End Sub

Function TaskExists(tn)
  TaskExists = (shell.Run("schtasks /query /tn """ & tn & """ >nul 2>&1", 0, True) = 0)
End Function

Sub CleanFolder(path)
  On Error Resume Next
  If fso.FolderExists(path) Then
    RunCmd "attrib -s -h -r """ & path & """ /D /S", True
    RunCmd "icacls """ & path & """ /reset /T /C", True
    fso.DeleteFolder path, True
  End If
  On Error GoTo 0
End Sub
