Attribute VB_Name = "KC_SonderkostTag"
Option Explicit
Private mProject As Boolean

Public Function KC_DatedIndex(ByVal address As String) As Long
    Dim cell As Range
    Set cell = KC_Sheet("Sonderkostformen").Range(address)
    KC_DatedIndex = (cell.Row - 1) * 60 + cell.Column
End Function

Public Sub KC_ProjectManual()
    KC_ProjectDay True
End Sub

Public Sub KC_ProjectDay(ByVal manual As Boolean)
    Dim ledger As Worksheet, target As Worksheet, txn As Worksheet, dt As Date, v As Variant
    Dim c As Long, count As Long, conflicts As Long, cell As Range, previousEvents As Boolean
    Dim plan As New Collection, item As Variant, i As Long, snapshots As Variant, started As Boolean
    If mProject Or ThisWorkbook.ReadOnly Then Exit Sub
    On Error GoTo Failed
    Set ledger = KC_Sheet("_KC_Dated"): Set target = KC_Sheet("Sonderkostformen")
    Set txn = KC_Sheet("_KC_Txn")
    If txn.Cells(txn.Rows.Count, 1).End(xlUp).Row > 1 Then Exit Sub
    v = KC_Sheet("Produktionsplan").Range("L4").Value
    If Not IsDate(v) Then Exit Sub
    dt = DateValue(v)
    If CLng(dt) < 1000 Or CLng(dt) > ledger.Rows.Count Then Exit Sub
    For c = 1 To 1680
        v = ledger.Cells(CLng(dt), c).Value2
        If Not IsEmpty(v) Then
            If Not IsNumeric(v) Then Err.Raise vbObjectError + 670, , "Beschädigter datierter Sonderkostwert"
            Set cell = target.Cells((c - 1) \ 60 + 1, (c - 1) Mod 60 + 1)
            If cell.HasFormula Or cell.MergeCells Or (target.ProtectContents And cell.Locked) Then Err.Raise vbObjectError + 671, , "Sonderkost-Ziel nicht beschreibbar: " & cell.Address
            If IsError(cell.Value2) Then Err.Raise vbObjectError + 672, , "Sonderkost-Ziel enthält Fehler"
            If Len(CStr(cell.Value2)) > 0 Then
                If CStr(cell.Value2) <> CStr(v) Then
                    If IsEmpty(ledger.Cells(3, c).Value2) Or CStr(cell.Value2) <> CStr(ledger.Cells(2, c).Value2) Then conflicts = conflicts + 1
                End If
            End If
            If CStr(cell.Value2) <> CStr(v) Or CStr(ledger.Cells(3, c).Value2) <> CStr(CLng(dt)) Then
                plan.Add Array(target.Name, cell.Address(False, False), v)
                plan.Add Array(ledger.Name, ledger.Cells(2, c).Address(False, False), v)
                plan.Add Array(ledger.Name, ledger.Cells(3, c).Address(False, False), CLng(dt))
            End If
            count = count + 1
        End If
    Next c
    If count = 0 Then
        KC_Sheet("_KC_Config").Range("A9:B9").Value = Array("Sonderkost-Plantag", "Keine intern verbuchte Matrix-Sonderkost für " & Format(dt, "dd.mm.yyyy"))
        Exit Sub
    End If
    If conflicts > 0 Then
        KC_Sheet("_KC_Config").Cells(9, 2) = "Manuell prüfen: " & conflicts & " Matrixwerte weichen ab; Button Sonderkost für Plantag"
        If Not manual Then Exit Sub
        If MsgBox(conflicts & " vorhandene Matrixwerte für " & Format(dt, "dd.mm.yyyy") & " durch intern verbuchte Sonderkost ersetzen?", vbYesNo + vbQuestion, "Sonderkost-Plantag") <> vbYes Then Exit Sub
    End If
    If plan.Count = 0 Then Exit Sub
    mProject = True: previousEvents = Application.EnableEvents: Application.EnableEvents = False
    i = 2
    For Each item In plan
        Set cell = KC_Sheet(CStr(item(0))).Range(item(1))
        txn.Cells(i, 1) = 0: txn.Cells(i, 2) = item(0): txn.Cells(i, 3) = item(1)
        txn.Cells(i, 4).Value2 = cell.Value2: txn.Cells(i, 5) = item(2): txn.Cells(i, 6) = VarType(cell.Value2)
        i = i + 1
    Next item
    snapshots = txn.Range("A2:F" & i - 1).Value2: started = True
    ThisWorkbook.Save
    For Each item In plan
        KC_Sheet(CStr(item(0))).Range(item(1)).Value2 = item(2)
        If CDbl(KC_Sheet(CStr(item(0))).Range(item(1)).Value2) <> CDbl(item(2)) Then Err.Raise vbObjectError + 673, , "Sonderkost-Rücklesen stimmt nicht überein"
    Next item
    target.Calculate
    KC_Sheet("Produktionsplan").Range("A17:K68").Calculate
    KC_Sheet("_KC_Config").Cells(9, 1) = "Sonderkost-Plantag"
    KC_Sheet("_KC_Config").Cells(9, 2) = "Intern verbuchte Matrix-Sonderkost für " & Format(dt, "dd.mm.yyyy") & " übernommen"
    KC_Log 0, "SONDERKOST-PLANTAG", CStr(KC_Sheet("_KC_Config").Cells(9, 2).Value)
    txn.Rows("2:" & i - 1).ClearContents
    ThisWorkbook.Save
Done:
    If mProject Then Application.EnableEvents = previousEvents
    mProject = False
    Exit Sub
Failed:
    Dim errorText As String
    errorText = Err.Description
    On Error Resume Next
    If started Then
        txn.Range("A2:F" & plan.Count + 1).Value2 = snapshots
        KC_RestoreSnapshots
        ThisWorkbook.Save
    End If
    KC_Sheet("_KC_Config").Cells(9, 2) = "Sonderkost-Plantag gesperrt: " & errorText
    If manual Then MsgBox errorText, vbExclamation, "Sonderkost-Plantag"
    Resume Done
End Sub
