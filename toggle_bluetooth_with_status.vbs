Set WshShell = WScript.CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' 1. Status via PowerShell abfragen und in Datei schreiben
' Wir nutzen einen Befehl, der robuster ist als GetResults
cmd = "powershell -command ""(Get-NetAdapter | Where-Object {$_.InterfaceDescription -match 'Bluetooth'}).Status"" > bt_status.txt"
WshShell.Run cmd, 0, True

' 2. Status aus der Datei lesen
Set file = FSO.OpenTextFile("bt_status.txt", 1)
status = Trim(file.ReadAll)
file.Close
FSO.DeleteFile("bt_status.txt")

' 3. Status anzeigen und Entscheidung treffen
if status = "Up" Then
    msg = "Bluetooth ist AKTIVIERT. Ausschalten?"
Else
    msg = "Bluetooth ist DEAKTIVIERT. Einschalten?"
End If

choice = MsgBox(msg, vbYesNo + vbQuestion, "Bluetooth Status")

' 4. Schalten, wenn auf "Ja" geklickt wurde
If choice = vbYes Then
    WshShell.Run "ms-settings:bluetooth"
    WScript.Sleep 1500 ' Warten, bis App offen ist
    WshShell.SendKeys "{TAB}" ' Zum Schalter navigieren
    WScript.Sleep 200
	WshShell.SendKeys "{TAB}" ' Zum Schalter navigieren
    WScript.Sleep 200
	WshShell.SendKeys "{TAB}" ' Zum Schalter navigieren
    WScript.Sleep 200
	WshShell.SendKeys "{TAB}" ' Zum Schalter navigieren
    WScript.Sleep 200
    WshShell.SendKeys " " ' Umschalten
    WScript.Sleep 500
    WshShell.SendKeys "%{F4}" ' Fenster schließen
End If