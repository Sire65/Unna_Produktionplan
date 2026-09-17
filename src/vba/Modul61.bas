Attribute VB_Name = "Modul61"
'Option Explicit
'
'Sub DruckProduktionSeiten1und2()
'
'Dim wsZiel As Worksheet
'Dim bestaetigt As VbMsgBoxResult
'Dim wsStart As Worksheet
'
'' Zielblatt festlegen
'Set wsZiel = ThisWorkbook.Sheets("Produktionsplan")
'' Startseite festlegen
'Set wsStart = ThisWorkbook.Sheets("Startseite")
'
'' Zu Zielblatt wechseln
'wsZiel.Activate
'
'' Bestätigung abfragen
'bestaetigt = MsgBox("Wirklich den Produktionsplan Seite 1 und 2 drucken?", vbQuestion + vbYesNo, "Druck bestätigen")
'
'If bestaetigt = vbYes Then
'    ' Druckbereich festlegen (Seite 1 und 2 drucken)
'    wsZiel.PrintOut From:=1, To:=2, Preview:=False
'    ' Nach dem Drucken zurück zur Startseite
'    wsStart.Activate
'Else
'    ' Abbruch, zurück zur Startseite
'    wsStart.Activate
'    MsgBox "Druck abgebrochen.", vbInformation, "Abbruch"
'End If
'End Sub
Sub DruckProduktionSeiten1und2()
    Dim wsZiel As Worksheet
    Dim bestaetigt As VbMsgBoxResult
    Dim wsStart As Worksheet
    ' Zielblatt festlegen
    Set wsZiel = ThisWorkbook.Sheets("Produktionsplan")
    ' Startseite festlegen
    Set wsStart = ThisWorkbook.Sheets("Startseite")
    ' Zu Zielblatt wechseln
    wsZiel.Activate
    ' Bestätigung abfragen
    bestaetigt = MsgBox( _
        "Wirklich den Produktionsplan Seite 1 bis 3 drucken?", _
        vbQuestion + vbYesNo, _
        "Druck bestätigen")
    If bestaetigt = vbYes Then
        ' Produktionsplan Seite 1 bis 3 drucken
        wsZiel.PrintOut From:=1, To:=3, Preview:=False
        ' Nach dem Drucken zurück zur Startseite
        wsStart.Activate
    Else
        ' Abbruch, zurück zur Startseite
        wsStart.Activate
        MsgBox "Druck abgebrochen.", _
               vbInformation, _
               "Abbruch"
    End If
End Sub
'Sub DruckReistabelle()
'    ' Reistabelle drucken: Seiten 5 und 6 im Produktionsplan
'    Dim wsZiel As Worksheet
'    Dim bestaetigt As VbMsgBoxResult
'
'    ' Zielblatt festlegen (immer in dieser Arbeitsmappe)
'    Set wsZiel = ThisWorkbook.Sheets("Produktionsplan")
'
'    ' Sicherheitsabfrage
'    bestaetigt = MsgBox("Wirklich die Reistabelle Seite 5 und 6 drucken?", _
'                        vbQuestion + vbYesNo, "Druck bestätigen")
'
'    If bestaetigt = vbYes Then
'        ' Druckbereich festlegen (Seiten 5 und 6 drucken)
'        ' Excel verwendet die Seitenumbrüche des Blatts
'        wsZiel.PrintOut From:=5, To:=6, Copies:=1, Preview:=False
'
'        ' Optional: Nach dem Druck zur Startseite wechseln
'        On Error Resume Next
'        ThisWorkbook.Sheets("Startseite").Activate
'        On Error GoTo 0
'    Else
'        ' Abbruch, nichts tun
'        MsgBox "Druck abgebrochen.", vbInformation, "Abbruch"
'        ' Optional: Zur Startseite wechseln
'        On Error Resume Next
'        ThisWorkbook.Sheets("Startseite").Activate
'        On Error GoTo 0
'    End If
'End Sub
Sub DruckReistabelle()
    Dim wsZiel As Worksheet
    Dim bestaetigt As VbMsgBoxResult
    Dim alterDruckbereich As String
    ' Zielblatt festlegen
    Set wsZiel = ThisWorkbook.Sheets("Produktionsplan")
    ' Sicherheitsabfrage
    bestaetigt = MsgBox( _
        "Wirklich die Reistabelle drucken?", _
        vbQuestion + vbYesNo, _
        "Druck bestätigen")
    If bestaetigt = vbYes Then
        On Error GoTo Fehler
        ' Bisherigen Druckbereich merken
        alterDruckbereich = wsZiel.PageSetup.PrintArea
        ' Reistabelle als Druckbereich festlegen
        wsZiel.PageSetup.PrintArea = "$AG$1:$AP$83"
        ' Reistabelle drucken
        wsZiel.PrintOut Copies:=1, Preview:=False
        ' Ursprünglichen Druckbereich wiederherstellen
        wsZiel.PageSetup.PrintArea = alterDruckbereich
        ' Zurück zur Startseite
        ThisWorkbook.Sheets("Startseite").Activate
        Exit Sub
    Else
        MsgBox "Druck abgebrochen.", vbInformation, "Abbruch"
        On Error Resume Next
        ThisWorkbook.Sheets("Startseite").Activate
        On Error GoTo 0
        Exit Sub
    End If
Fehler:
    ' Bei einem Fehler den ursprünglichen Druckbereich
    ' nach Möglichkeit ebenfalls wiederherstellen
    On Error Resume Next
    wsZiel.PageSetup.PrintArea = alterDruckbereich
    ThisWorkbook.Sheets("Startseite").Activate
    MsgBox "Fehler beim Drucken der Reistabelle: " & _
           Err.Description, vbExclamation, "Druckfehler"
    On Error GoTo 0
End Sub
'=========================================================
' Modul: DruckeBereichMengenabfrage
'=========================================================
' Beschreibung: Dieses Makro druckt einen definierten Bereich
'               ("M5:AB32") des Blattes "Transport" nach
'               Bestätigung durch den Benutzer.
'
' Zweck:       Automatisierter Druck des Menganabfrage-Zettels
'
' Autor:       [Dein Name]
' Erstellt:    28.10.2025
' Letzte Änderung: 28.10.2025
'
' Parameter:   Keine
' Rückgabe:    Keine
' Bemerkungen: - Prüft, ob das Blatt "Transport" existiert
'               - Sicherheitsabfrage über MsgBox
'               - Druckbereich wird auf M5:AB32 gesetzt
'               - Druck erfolgt nur bei Bestätigung "Ja"
'               - Wechselt optional zurück zum Blatt "Startseite"
'=========================================================
Sub DruckeBereichMegenabfrage()
    If MsgBox("Mengenbewertung drucken?", vbYesNo + vbQuestion, "Druckauftrag bestätigen") <> vbYes Then Exit Sub
    On Error GoTo Failed
    KC_TransportPrint "MENGENBEWERTUNG"
    ThisWorkbook.Worksheets("Startseite").Activate
    Exit Sub
Failed:
    MsgBox "Mengenbewertung konnte nicht gedruckt werden: " & Err.Description, vbExclamation, "Druckfehler"
End Sub
Private Sub SonderkostformenDrucken()
Dim wsZiel As Worksheet
Dim confirmPrint As VbMsgBoxResult
' Zielblatt setzen
Set wsZiel = ThisWorkbook.Sheets("Sonderkostformen")
' Wechsel zum Zielblatt
wsZiel.Activate
' Abfrage, ob gedruckt werden soll
confirmPrint = MsgBox("Wirklich die Sonderkostformen drucken?", _
                      vbQuestion + vbYesNo, "Druckbestätigung")
If confirmPrint = vbYes Then
    ' Seite 1 drucken (Anpassung: Bereich oder komplette Seite je nach Definition)
    ' Beispiel 1: Drucken der ersten Seite des aktuellen Blatts
    wsZiel.PrintOut From:=1, To:=1, Preview:=False
    ' Alternative falls Druck eines bestimmten Bereichs/Blatts gewünscht:
    ' If you want to print the whole sheet, einfach:
    ' wsZiel.PrintOut
    ' Falls nur Seite 1 eines mehrseitigen Blatts gemeint ist, nutze From/To wie oben.
End If
End Sub
Sub DruckMakro() 'Druckt alle benötigten Formulare und Pläne für den gewählten Tag
Dim bestätigung As VbMsgBoxResult
Dim ws As Worksheet
Dim blattNamen As Variant
Dim i As Long
'Sicherheitsabfrage
bestätigung = MsgBox("Möchten Sie den Druckvorgang starten?", vbQuestion + vbYesNo, "Sicherheitsabfrage")
If bestätigung <> vbYes Then Exit Sub
On Error GoTo Ende
Application.ScreenUpdating = False
Application.EnableEvents = False
' Produktionsplan – Seite 1 und 2
Set ws = ThisWorkbook.Sheets("Produktionsplan")
ws.PrintOut From:=1, To:=1, Copies:=1
ws.PrintOut From:=2, To:=2, Copies:=1
' Blätterliste zum Drucken
blattNamen = Array("Kita1", "Kita2_4", "Liedbach", "Strolche", "Lünern", "Lünern2", _
                   "Cluster1", "Cluster2", "Cluster3", "Cluster4")
For i = LBound(blattNamen) To UBound(blattNamen)
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(blattNamen(i))
    If Not ws Is Nothing Then
        ws.PrintOut Copies:=1
    End If
    On Error GoTo 0
    Set ws = Nothing
Next i
' Sonderkostformen – Seite 1
Set ws = ThisWorkbook.Sheets("Sonderkostformen")
ws.PrintOut From:=1, To:=1, Copies:=1
' Transport – Seite 1
Set ws = ThisWorkbook.Sheets("Transport")
KC_TransportPrint "TOURENPLAN"
Ende: Application.ScreenUpdating = True
Application.EnableEvents = True
If Err.Number <> 0 Then
    MsgBox "Fehler beim Drucken: " & Err.Description, vbExclamation
    ' Optional: Err.Clear oder weitere Fehlerbehandlung
End If
End Sub
'Sub DruckeTransportSeite1() 'BeispielMitBestätigung() Dim resp As VbMsgBoxResult Dim zielSheet As Worksheet
'
''Zum Blatt "Transport" wechseln
'On Error Resume Next
'Set ws = ThisWorkbook.Worksheets("Transport")
'Dim ws As Worksheet
'
'On Error GoTo 0
'
'If ws Is Nothing Then
'    MsgBox "Das Blatt 'Transport' wurde nicht gefunden.", vbExclamation
'    Exit Sub
'End If
'
'ws.Activate
'
''Sicherheitsabfrage (Bestätigen mit "Ja")
'bestätigung = MsgBox("Möchten Sie Seite 1 des Blattes ""Transport"" drucken?", _
'                      vbQuestion + vbYesNo, "Druckauftrag bestätigen")
'
'If bestätigung = vbYes Then
'    'Drucke Seite 1 (erste zu druckende Seite, 1 Exemplar)
'    ActiveSheet.PrintOut From:=1, To:=1, Copies:=1, Collate:=True
'End If
'
''Unabhängig von Ja/Nein/Abbruch: zurück zum Blatt "Startseite", falls vorhanden
'On Error Resume Next
'Set zielBlatt = ThisWorkbook.Worksheets("Startseite")
'On Error GoTo 0
'
'If Not zielBlatt Is Nothing Then
'    zielBlatt.Activate
'Else
'    MsgBox "Das Blatt 'Startseite' wurde nicht gefunden. Druck-Abschluss, aber Zielblatt nicht gesetzt.", vbInformation
'End If
'End Sub
'=========================================================
' Modul: DruckeTransportSeite1
'=========================================================
' Beschreibung: Dieses Makro druckt die erste Seite des Blattes "Transport"
'               nach Bestätigung durch den Benutzer und wechselt danach
'               zurück zum Blatt "Startseite" (falls vorhanden).
'
' Zweck:       Automatisierter Druck mit Sicherheitsabfrage
'
' Autor:       [Dein Name]
' Erstellt:    28.10.2025
' Letzte Änderung: 28.10.2025
'
' Parameter:   Keine
' Rückgabe:    Keine
' Bemerkungen: - Prüft, ob die Blätter "Transport" und "Startseite" existieren
'               - Nutzt MsgBox für Benutzerbestätigung
'               - Druck nur, wenn Benutzer mit "Ja" bestätigt
'=========================================================
Sub DruckeTransportSeite1()
    If MsgBox("Tourenplan (Seite 1) drucken?", vbYesNo + vbQuestion, "Druckauftrag bestätigen") <> vbYes Then Exit Sub
    On Error GoTo Failed
    KC_TransportPrint "TOURENPLAN"
    ThisWorkbook.Worksheets("Startseite").Activate
    Exit Sub
Failed:
    MsgBox "Tourenplan (Seite 1) konnte nicht gedruckt werden: " & Err.Description, vbExclamation, "Druckfehler"
End Sub
