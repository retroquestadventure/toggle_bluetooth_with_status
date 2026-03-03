Set WshShell = WScript.CreateObject("WScript.Shell")

' --- ADMIN-RECHTE ERZWINGEN ---
If Not WScript.Arguments.Named.Exists("admin") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & _
    WScript.ScriptFullName & """ /admin", "", "runas", 1
    WScript.Quit
End If

Dim regPath
regPath = "HKLM\SYSTEM\CurrentControlSet\Services\bthserv\Parameters\RadioSupport"

' --- FUNKTION: STATUS LESEN ODER INITIAL ERSTELLEN ---
Function GetBTStatus()
    On Error Resume Next
    Dim val
    val = WshShell.RegRead(regPath)
    If Err.Number <> 0 Then
        Err.Clear
        WshShell.RegWrite regPath, 1, "REG_DWORD"
        GetBTStatus = "AN"
    Else
        If val = 1 Then GetBTStatus = "AN" Else GetBTStatus = "AUS"
    End If
End Function

' --- HAUPTABLAUF ---
currentStatus = GetBTStatus()

If currentStatus = "AN" Then
    msg = "Bluetooth ist AKTIVIERT. Ausschalten?"
    newStatusValue = 0 
Else
    msg = "Bluetooth ist DEAKTIVIERT. Einschalten?"
    newStatusValue = 1 
End If

ans = MsgBox(msg, 36, "Bluetooth Manager")

If ans = 6 Then
    ' --- FENSTER-RESET ---
    ' Beendet alle Instanzen der Einstellungen, damit wir bei "Null" starten
    WshShell.Run "taskkill /F /IM SystemSettings.exe", 0, True
    WScript.Sleep 500 ' Kurze Pause zum Schließen
    
    ' 1. Einstellungen frisch öffnen
    WshShell.Run "ms-settings:bluetooth"
    WScript.Sleep 2500 
    
    ' Deine 4 TABs (jetzt auf sicherem Boden)
    WshShell.SendKeys "{TAB}"
    WScript.Sleep 200
    WshShell.SendKeys "{TAB}"
    WScript.Sleep 200
    WshShell.SendKeys "{TAB}"
    WScript.Sleep 200
    WshShell.SendKeys "{TAB}"
    WScript.Sleep 200
    
    WshShell.SendKeys " "
    WScript.Sleep 1000
    WshShell.SendKeys "%{F4}"
    
    ' 2. Registry synchronisieren
    WshShell.RegWrite regPath, newStatusValue, "REG_DWORD"
End If