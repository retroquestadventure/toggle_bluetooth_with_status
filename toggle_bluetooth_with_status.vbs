Set WshShell = WScript.CreateObject("WScript.Shell")

' --- ADMIN-RECHTE ERZWINGEN ---
If Not WScript.Arguments.Named.Exists("admin") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & _
    WScript.ScriptFullName & """ /admin", "", "runas", 1
    WScript.Quit
End If

Dim regPath
regPath = "HKLM\SYSTEM\CurrentControlSet\Services\bthserv\Parameters\RadioSupport"

' --- FUNKTION: STATUS LESEN ---
Function GetBTStatus()
    On Error Resume Next
    statusValue = WshShell.RegRead(regPath)
    If Err.Number <> 0 Then
        WshShell.RegWrite regPath, 1, "REG_DWORD"
        GetBTStatus = "AN"
        Err.Clear
    Else
        If statusValue = 1 Then GetBTStatus = "AN" Else GetBTStatus = "AUS"
    End If
End Function

' --- HAUPTABLAUF ---
currentStatus = GetBTStatus()

If currentStatus = "AN" Then
    msg = "Bluetooth ist AKTIVIERT. Ausschalten?"
    newStatusValue = 0 ' Zielwert für Registry nach dem Umschalten
Else
    msg = "Bluetooth ist DEAKTIVIERT. Einschalten?"
    newStatusValue = 1 ' Zielwert für Registry nach dem Umschalten
End If

ans = MsgBox(msg, 36, "Bluetooth Manager (Sync-Mode)")

If ans = 6 Then
    ' 1. Einstellungen öffnen und umschalten
    WshShell.Run "ms-settings:bluetooth"
    WScript.Sleep 2500 
    
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
    
    ' 2. Registry manuell aktualisieren, damit der Status sofort stimmt
    WshShell.RegWrite regPath, newStatusValue, "REG_DWORD"
    
    ' Optional: Kurze Bestätigung
    ' MsgBox "Status in Registry aktualisiert!", 64, "Info"
End If