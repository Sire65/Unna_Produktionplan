Attribute VB_Name = "KC_Verbuchung"
Option Explicit
Private mBooking As Boolean

Public Sub KC_AutoBook()
    Dim q As Worksheet, r As Long
    If KC_Sheet("_KC_Config").Cells(2, 2).Value <> "EIN" Then Exit Sub
    Set q = KC_Sheet("_KC_Q")
    For r = 2 To q.Cells(q.Rows.Count, 1).End(xlUp).row
        If q.Cells(r, 12).Value = KC_READY Then
            KC_BookRow r, True
            Exit For
        End If
    Next r
End Sub

Public Function KC_BookRow(ByVal r As Long, ByVal automatic As Boolean) As Boolean
    Dim q As Worksheet, txn As Worksheet, e As Object, plan As Collection, item As Variant
    Dim i As Long, cell As Range, previousEvents As Boolean, originalStatus As String, errorText As String
    Dim journalStarted As Boolean, auditRow As Long, logRow As Long, snapshots As Variant
    Dim recalcSheets As Object, sheetName As Variant
    If mBooking Or ThisWorkbook.ReadOnly Then Exit Function
    If automatic And KC_Sheet("_KC_Config").Cells(2, 2).Value <> "EIN" Then Exit Function
    Set q = KC_Sheet("_KC_Q")
    If r < 2 Or r > q.Cells(q.Rows.Count, 1).End(xlUp).row Then Exit Function
    If q.Cells(r, 12).Value = KC_BOOKED Or q.Cells(r, 12).Value = "VERBUCHUNG LÄUFT" Then Exit Function
    If automatic And q.Cells(r, 12).Value <> KC_READY Then Exit Function
    ' Re-evaluate from immutable internal capture, never trust the displayed worksheet.
    Set e = KC_Evaluate(CStr(q.Cells(r, 7).Value), KC_Body(CLng(q.Cells(r, 1).Value)))
    KC_StoreEvaluation r, e: KC_CheckHistory r, e
    If q.Cells(r, 12).Value = KC_ERROR Or Not e("bookable") Then
        If Not automatic Then MsgBox "Keine Verbuchung: " & q.Cells(r, 13).Value, vbExclamation, "KitaFino"
        Exit Function
    End If
    If automatic And q.Cells(r, 12).Value <> KC_READY Then Exit Function
    If Not automatic Then
        If MsgBox("Diese Bestellung verbuchen?" & vbCrLf & q.Cells(r, 7).Value & vbCrLf & q.Cells(r, 13).Value & vbCrLf & "Abweichende vorhandene Werte und freigegebene Sonderkost werden vollständig ersetzt.", vbYesNo + vbQuestion, "KitaFino – manuell verbuchen") <> vbYes Then Exit Function
    End If
    mBooking = True: previousEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set txn = KC_Sheet("_KC_Txn")
    If txn.Cells(txn.Rows.Count, 1).End(xlUp).row > 1 Then Err.Raise vbObjectError + 650, , "Offene Transaktion zuerst wiederherstellen"
    Set plan = e("plan")
    If plan.Count = 0 Then Err.Raise vbObjectError + 651, , "Leerer Schreibplan"
    originalStatus = CStr(q.Cells(r, 12).Value)
    i = 2
    For Each item In plan
        Set cell = ThisWorkbook.Worksheets(item(0)).Range(item(1))
        txn.Cells(i, 1) = q.Cells(r, 1).Value
        txn.Cells(i, 2) = item(0): txn.Cells(i, 3) = item(1)
        txn.Cells(i, 4).Value2 = cell.Value2: txn.Cells(i, 5) = item(2)
        txn.Cells(i, 6) = VarType(cell.Value2)
        i = i + 1
    Next item
    snapshots = txn.Range("A2:F" & i - 1).Value2
    q.Cells(r, 12) = "VERBUCHUNG LÄUFT"
    journalStarted = True
    ThisWorkbook.Save
    ' No production writes before journal has been persisted successfully.
    For Each item In plan
        ThisWorkbook.Worksheets(item(0)).Range(item(1)).Value2 = CLng(item(2))
    Next item
    For Each item In plan
        Set cell = ThisWorkbook.Worksheets(item(0)).Range(item(1))
        If IsError(cell.Value2) Then Err.Raise vbObjectError + 652, , "Fehler beim Rücklesen"
        If Not IsNumeric(cell.Value2) Then Err.Raise vbObjectError + 653, , "Nichtnumerischer Wert beim Rücklesen"
        If CDbl(cell.Value2) <> CDbl(item(2)) Then Err.Raise vbObjectError + 654, , "Rücklesen stimmt nicht überein"
    Next item
    Set recalcSheets = CreateObject("Scripting.Dictionary")
    For Each item In plan
        recalcSheets(CStr(item(0))) = True
    Next item
    For Each sheetName In recalcSheets.Keys
        If Left(CStr(sheetName), 4) <> "_KC_" Then ThisWorkbook.Worksheets(CStr(sheetName)).UsedRange.Calculate
    Next sheetName
    KC_Sheet("Produktionsplan").Range("A17:K68").Calculate
    KC_Sheet("Transport").Range("AZ3:BN250").Calculate
    KC_Sheet("Transport").Range("A6:C31").Calculate
    q.Cells(r, 12) = KC_BOOKED: q.Cells(r, 14) = Now
    q.Cells(r, 13) = "Vollständig verbucht und rückgelesen"
    auditRow = KC_WriteLegacyLog(r, plan)
    logRow = KC_Sheet("_KC_Log").Cells(KC_Sheet("_KC_Log").Rows.Count, 1).End(xlUp).row + 1
    KC_Log r, "VERBUCHT", "Atomare Verbuchung; " & plan.Count & " Zielzellen"
    txn.Rows("2:" & i - 1).ClearContents
    ' Snapshot rows remain in memory until successful final save for rollback.
    KC_Refresh
    ThisWorkbook.Save
    KC_BookRow = True
Done:
    Application.EnableEvents = previousEvents
    mBooking = False
    Exit Function
Failed:
    errorText = Err.Description
    On Error Resume Next
    If journalStarted Then
        txn.Range("A2:F" & plan.Count + 1).Value2 = snapshots
        KC_RestoreSnapshots
        q.Cells(r, 12) = KC_ERROR: q.Cells(r, 14).ClearContents
        q.Cells(r, 13) = "Verbuchung abgebrochen/zurückgesetzt: " & errorText
        If auditRow > 0 Then KC_Sheet("Import_Protokoll").Rows(auditRow).ClearContents
        If logRow > 0 Then KC_Sheet("_KC_Log").Rows(logRow).ClearContents
        KC_Log r, "VERBUCHUNGSFEHLER", errorText
        ThisWorkbook.Save
    End If
    If Not automatic Then MsgBox "Verbuchung fehlgeschlagen: " & errorText, vbCritical
    Resume Done
End Function

Private Function KC_WriteLegacyLog(ByVal r As Long, plan As Collection) As Long
    Dim ws As Worksheet, q As Worksheet, lr As Long, item As Variant, targets As String, key As String
    Set ws = KC_Sheet("Import_Protokoll"): Set q = KC_Sheet("_KC_Q")
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    key = CStr(q.Cells(r, 5).Value): If key = "" Then key = CStr(q.Cells(r, 3).Value)
    If key = "" Then key = "KC-AUTARK-" & q.Cells(r, 1).Value
    ws.Cells(lr, 1) = "KC-AUTARK-" & q.Cells(r, 1).Value
    ws.Cells(lr, 2) = Date: ws.Cells(lr, 3) = Time: ws.Cells(lr, 4) = Environ$("USERNAME")
    ws.Cells(lr, 5).NumberFormat = "@": ws.Cells(lr, 5) = q.Cells(r, 7).Value
    ws.Cells(lr, 6).NumberFormat = "@": ws.Cells(lr, 6) = key
    ws.Cells(lr, 7) = q.Cells(r, 8).Value: ws.Cells(lr, 8) = q.Cells(r, 10).Value
    ws.Cells(lr, 10) = q.Cells(r, 9).Value: ws.Cells(lr, 11) = "Interner geprüfter Schreibplan"
    For Each item In plan
        If targets <> "" Then targets = targets & "; "
        targets = targets & item(0) & "!" & item(1) & "=" & item(2)
    Next item
    ws.Cells(lr, 14) = Left(targets, 32000): ws.Cells(lr, 15) = "JA": ws.Cells(lr, 16) = "GRUEN"
    ws.Cells(lr, 17) = "Autark, atomar, rückgelesen": ws.Cells(lr, 18) = "Outlook/VBA intern"
    ws.Cells(lr, 20) = Left(targets, 32000): ws.Cells(lr, 22) = q.Cells(r, 15).Value
    ws.Cells(lr, 23) = KC_V20.DateText(q.Cells(r, 10).Value) & " - " & KC_V20.DateText(q.Cells(r, 11).Value)
    KC_WriteLegacyLog = lr
End Function

Public Sub KC_Recover()
    Dim q As Worksheet, txn As Worksheet, r As Long, id As Variant, events As Boolean
    Set txn = KC_Sheet("_KC_Txn"): Set q = KC_Sheet("_KC_Q")
    If txn.Cells(txn.Rows.Count, 1).End(xlUp).row < 2 Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    events = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    id = txn.Cells(2, 1).Value
    KC_RestoreSnapshots
    For r = 2 To q.Cells(q.Rows.Count, 1).End(xlUp).row
        If q.Cells(r, 1).Value = id Then
            q.Cells(r, 12) = KC_ERROR
            q.Cells(r, 13) = "Offene Verbuchung nach Neustart zurückgesetzt; manuelle Prüfung erforderlich"
            q.Cells(r, 14).ClearContents
            KC_Log r, "WIEDERHERSTELLUNG", CStr(q.Cells(r, 13).Value)
        End If
    Next r
    ThisWorkbook.Save
    Application.EnableEvents = events
    KC_Refresh
    Exit Sub
Failed:
    Application.EnableEvents = events
    Err.Raise vbObjectError + 655, , "Wiederherstellung fehlgeschlagen: " & Err.Description
End Sub

Public Sub KC_RestoreSnapshots()
    Dim txn As Worksheet, r As Long, cell As Range, v As Variant
    Set txn = KC_Sheet("_KC_Txn")
    For r = 2 To txn.Cells(txn.Rows.Count, 1).End(xlUp).row
        Set cell = ThisWorkbook.Worksheets(CStr(txn.Cells(r, 2).Value)).Range(CStr(txn.Cells(r, 3).Value))
        v = txn.Cells(r, 4).Value2
        Select Case CLng(txn.Cells(r, 6).Value)
            Case vbEmpty: cell.ClearContents
            Case vbString: cell.Value2 = CStr(v)
            Case vbBoolean: cell.Value2 = CBool(v)
            Case Else: cell.Value2 = CDbl(v)
        End Select
    Next r
    If r > 2 Then txn.Rows("2:" & r - 1).ClearContents
End Sub



