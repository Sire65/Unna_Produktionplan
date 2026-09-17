Attribute VB_Name = "Modul66"
Option Explicit
' ---- KONFIG ----
'' Liste der auswählbaren "Targets":
''   "PP" = Produktionsplan (druckt Seite 1 und 2)
''   "SK" = Sonderkostformen (nur Seite 1)
''   "TR" = Transport (nur Seite 1)
''   Alle anderen Einträge werden als Blattnamen behandelt (ganze Seite)
'Private Function Kandidaten() As Variant
'    Kandidaten = Array( _
'        "PP", _
'        "Kita1", "Kita2_4", "Liedbach", "Strolche", "Lünern", "Lünern2", _
'        "Cluster1", "Cluster2", "Cluster3", "Cluster4", "Kontakt", _
'        "SK", "TR" _
'    )
'End Function
'' ---- ENDE KONFIG ----
'
'
'
'Public Function ShowPrintPicker() As Collection
'    Dim dlg As Object
'
'    On Error GoTo ErrHandler
'    Set dlg = VBA.UserForms.Add("UserForm13") ' <-- lädt "UserForm13"
'
'    dlg.InitWithKeys Kandidaten           ' muss in UserForm13 stehen
'    dlg.show vbModal
'
'    If dlg.DialogOK Then
'        Set ShowPrintPicker = dlg.SelectedKeys ' muss in UserForm13 stehen
'    Else
'        Set ShowPrintPicker = Nothing
'    End If
'
'Cleanup:
'    On Error Resume Next
'    If Not dlg Is Nothing Then Unload dlg
'    Set dlg = Nothing
'    Exit Function
'
'ErrHandler:
'    MsgBox "Druck-Dialog konnte nicht geöffnet werden: " & Err.Description, vbExclamation
'    Set ShowPrintPicker = Nothing
'    Resume Cleanup
'End Function
'
'
'Private Function BlattSafely(ByVal name As String) As Worksheet
'    On Error Resume Next
'    Set BlattSafely = ThisWorkbook.Sheets(name)
'    If Err.Number <> 0 Then Set BlattSafely = Nothing
'    On Error GoTo 0
'End Function
'
'' === DEIN BUTTON-MAKRO: jetzt mit Auswahl-Dialog ===
'Public Sub DruckMakro_neu()
'    Dim bestätigung As VbMsgBoxResult
'    Dim ws As Worksheet
'    Dim keys As Collection
'    Dim k As Variant
'
'    ' 1) Auswahl anzeigen
'    Set keys = ShowPrintPicker()
'    If keys Is Nothing Then Exit Sub             ' Abbrechen
'    If keys.count = 0 Then Exit Sub              ' Nichts gewählt
'
'    ' 2) Sicherheitsabfrage
'    bestätigung = MsgBox("Möchten Sie den Druckvorgang starten?", vbQuestion + vbYesNo, "Sicherheitsabfrage")
'    If bestätigung <> vbYes Then Exit Sub
'
'    On Error GoTo Ende
'    Application.ScreenUpdating = False
'    Application.EnableEvents = False
'
'    ' 3) Druck entsprechend der Auswahl
'    For Each k In keys
'        Select Case CStr(k)
'            Case "PP"
'                Set ws = BlattSafely("Produktionsplan")
'                If Not ws Is Nothing Then
'                    ws.PrintOut From:=1, To:=1, Copies:=1
'                    ws.PrintOut From:=2, To:=2, Copies:=1
'                End If
'
'            Case "SK"
'                Set ws = BlattSafely("Sonderkostformen")
'                If Not ws Is Nothing Then ws.PrintOut From:=1, To:=1, Copies:=1
'
'            Case "TR"
'                Set ws = BlattSafely("Transport")
'                If Not ws Is Nothing Then ws.PrintOut From:=1, To:=1, Copies:=1
'
'            Case Else
'                Set ws = BlattSafely(CStr(k))
'                If Not ws Is Nothing Then ws.PrintOut Copies:=1
'        End Select
'        Set ws = Nothing
'    Next k
'
'Ende:
'    Application.ScreenUpdating = True
'    Application.EnableEvents = True
'
'    If Err.Number <> 0 Then
'        MsgBox "Fehler beim Drucken: " & Err.Description, vbExclamation
'    End If
'End Sub
' ============================================================
' KONFIGURATION DER DRUCKAUSWAHL
'
' PP = Produktionsplan, Seiten 1 bis 3
' SK = Sonderkostformen, Seite 1
' SM = Sonderkost-Matrix, Bereich AK1:AY28
' TR = Transport, Seite 1
'
' Alle anderen Einträge werden als Blattnamen behandelt.
' ============================================================
Private Function Kandidaten() As Variant
    Kandidaten = Array( _
        "PP", _
        "Kita1", _
        "Kita2_4", _
        "Bäume", _
        "Liedbach", _
        "Strolche", _
        "Lünern", _
        "Lünern2", _
        "Cluster1", _
        "Cluster2", _
        "Cluster3", _
        "Cluster4", _
        "Kontakt", _
        "SK", _
        "SM", _
        "TR" _
    )
End Function
' ============================================================
' DRUCKAUSWAHL ÖFFNEN
' ============================================================
Public Function ShowPrintPicker() As Collection
    Dim dlg As Object
    On Error GoTo ErrHandler
    Set dlg = VBA.UserForms.Add("UserForm13")
    dlg.InitWithKeys Kandidaten
    dlg.show vbModal
    If dlg.DialogOK Then
        Set ShowPrintPicker = dlg.SelectedKeys
    Else
        Set ShowPrintPicker = Nothing
    End If
Cleanup:
    On Error Resume Next
    If Not dlg Is Nothing Then Unload dlg
    Set dlg = Nothing
    Exit Function
ErrHandler:
    MsgBox "Druck-Dialog konnte nicht geöffnet werden: " & _
           Err.Description, vbExclamation
    Set ShowPrintPicker = Nothing
    Resume Cleanup
End Function
' ============================================================
' BLATT SICHER ERMITTELN
' ============================================================
Private Function BlattSafely(ByVal name As String) As Worksheet
    On Error Resume Next
    Set BlattSafely = ThisWorkbook.Sheets(name)
    If Err.Number <> 0 Then
        Set BlattSafely = Nothing
    End If
    On Error GoTo 0
End Function
' ============================================================
' BUTTON-MAKRO
' ALLES DRUCKEN MIT AUSWAHL-DIALOG
' ============================================================
Public Sub DruckMakro_neu()
    Dim bestätigung As VbMsgBoxResult
    Dim ws As Worksheet
    Dim keys As Collection
    Dim k As Variant
    ' Sonderkost-Matrix
    Dim alterDruckbereich As String
    Dim matrixDruckAktiv As Boolean
    ' Fehlernummer und Fehlertext zwischenspeichern
    Dim fehlerNummer As Long
    Dim fehlerText As String
    ' --------------------------------------------------------
    ' 1) AUSWAHL ANZEIGEN
    ' --------------------------------------------------------
    Set keys = ShowPrintPicker()
    If keys Is Nothing Then Exit Sub
    If keys.count = 0 Then Exit Sub
    ' --------------------------------------------------------
    ' 2) SICHERHEITSABFRAGE
    ' --------------------------------------------------------
    bestätigung = MsgBox( _
        "Möchten Sie den Druckvorgang starten?", _
        vbQuestion + vbYesNo, _
        "Sicherheitsabfrage")
    If bestätigung <> vbYes Then Exit Sub
    ' --------------------------------------------------------
    ' 3) DRUCK VORBEREITEN
    ' --------------------------------------------------------
    On Error GoTo Fehler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' --------------------------------------------------------
    ' 4) AUSGEWÄHLTE POSITIONEN DRUCKEN
    ' --------------------------------------------------------
    For Each k In keys
        Select Case CStr(k)
            ' =================================================
            ' PRODUKTIONSPLAN
            ' Seiten 1, 2 und 3
            ' =================================================
            Case "PP"
                Set ws = BlattSafely("Produktionsplan")
                If Not ws Is Nothing Then
                    ws.PrintOut _
                        From:=1, _
                        To:=1, _
                        Copies:=1
                    ws.PrintOut _
                        From:=2, _
                        To:=2, _
                        Copies:=1
                    ws.PrintOut _
                        From:=3, _
                        To:=3, _
                        Copies:=1
                End If
            ' =================================================
            ' SONDERKOSTFORMEN
            ' Nur Seite 1
            ' =================================================
            Case "SK"
                Set ws = BlattSafely("Sonderkostformen")
                If Not ws Is Nothing Then
                    ws.PrintOut _
                        From:=1, _
                        To:=1, _
                        Copies:=1
                End If
            ' =================================================
            ' SONDERKOST-MATRIX
            ' Blatt: Sonderkostformen
            ' Bereich: AK1 bis AY28
            ' =================================================
            Case "SM"
                Set ws = BlattSafely("Sonderkostformen")
                If Not ws Is Nothing Then
                    ' Vorhandenen Druckbereich sichern
                    alterDruckbereich = ws.PageSetup.PrintArea
                    matrixDruckAktiv = True
                    ' Nur Matrix als Druckbereich festlegen
                    ws.PageSetup.PrintArea = "$AK$1:$AY$28"
                    ' Matrix drucken
                    ws.PrintOut Copies:=1
                    ' Ursprünglichen Druckbereich zurücksetzen
                    ws.PageSetup.PrintArea = alterDruckbereich
                    matrixDruckAktiv = False
                End If
            ' =================================================
            ' TRANSPORT
            ' Nur Seite 1
            ' =================================================
            Case "TR"
                KC_TransportPrint "TOURENPLAN"
            ' =================================================
            ' ALLE ÜBRIGEN BLÄTTER
            '
            ' Kita1
            ' Kita2_4
            ' Bäume
            ' Liedbach
            ' Strolche
            ' Lünern
            ' Lünern2
            ' Cluster1
            ' Cluster2
            ' Cluster3
            ' Cluster4
            ' Kontakt
            ' =================================================
            Case Else
                Set ws = BlattSafely(CStr(k))
                If Not ws Is Nothing Then
                    ws.PrintOut Copies:=1
                End If
        End Select
        Set ws = Nothing
    Next k
    GoTo Ende
' ============================================================
' FEHLERBEHANDLUNG
' ============================================================
Fehler:
    fehlerNummer = Err.Number
    fehlerText = Err.Description
    ' Falls der Fehler während des Matrixdrucks aufgetreten ist,
    ' ursprünglichen Druckbereich trotzdem wiederherstellen.
    On Error Resume Next
    If matrixDruckAktiv Then
        If Not ws Is Nothing Then
            ws.PageSetup.PrintArea = alterDruckbereich
        End If
        matrixDruckAktiv = False
    End If
    On Error GoTo 0
' ============================================================
' AUFRÄUMEN
' ============================================================
Ende:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If fehlerNummer <> 0 Then
        MsgBox "Fehler beim Drucken: " & _
               fehlerText, vbExclamation
    End If
End Sub
