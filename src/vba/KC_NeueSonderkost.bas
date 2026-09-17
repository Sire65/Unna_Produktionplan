Attribute VB_Name = "KC_NeueSonderkost"
Option Explicit
Private mRegistering As Boolean
Private Const LAST_FORM_ROW As Long = 200

Public Function KC_FormKey(ByVal raw As String) As String
    KC_FormKey = LCase(CStr(KC_V20.FlatText(raw)))
End Function

Public Sub KC_FormSetup()
    Dim ws As Worksheet, names As Variant, n As Variant
    names = Array("_KC_Forms", "_KC_FormTxn")
    For Each n In names
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(n))
        On Error GoTo 0
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CStr(n)
        End If
        ws.Visible = xlSheetVeryHidden
    Next n
    Set ws = KC_Sheet("_KC_Forms")
    ws.Range("A1:F1").Value = Array("Wortlaut-Schlüssel", "Wortlaut", "Matrixzeile", "Bestätigt am", "Bestätigt von", "Status")
    ws.Columns("A:B").NumberFormat = "@": ws.Columns("E:F").NumberFormat = "@"
    ws.Range("J1:L1").Value = Array("Kundenzeile", "Original AE", "Original AF")
    ws.Columns("K:L").NumberFormat = "@"
    Set ws = KC_Sheet("_KC_FormTxn")
    ws.Range("A1:E1").Value = Array("Blatt", "Zelle", "Vorherige Formel/Wert", "War Formel", "Werttyp")
    ws.Columns("A:C").NumberFormat = "@"
    KC_FormRecover
End Sub

Public Sub KC_FormRecover()
    Dim txn As Worksheet, r As Long, cell As Range, old As Variant
    Set txn = KC_Sheet("_KC_FormTxn")
    If txn.Cells(txn.Rows.Count, 1).End(xlUp).Row < 2 Then Exit Sub
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 701, , "Offene Sonderkost-Freigabe; Schreibzugriff zur Wiederherstellung erforderlich"
    For r = 2 To txn.Cells(txn.Rows.Count, 1).End(xlUp).Row
        Set cell = KC_Sheet(CStr(txn.Cells(r, 1).Value2)).Range(CStr(txn.Cells(r, 2).Value2))
        old = txn.Cells(r, 3).Value2
        If IsEmpty(old) Then
            cell.ClearContents
        ElseIf txn.Cells(r, 4).Value2 = True Then
            cell.Formula = old
        Else
            If txn.Cells(r, 5).Value2 = vbDouble Or txn.Cells(r, 5).Value2 = vbInteger Or txn.Cells(r, 5).Value2 = vbLong Then
                cell.Value2 = CDbl(old)
            Else
                cell.Value2 = old
            End If
        End If
    Next r
    txn.Rows("2:" & txn.Cells(txn.Rows.Count, 1).End(xlUp).Row).ClearContents
    ThisWorkbook.Save
End Sub

Public Function KC_FormRow(ByVal raw As String) As Long
    Dim ws As Worksheet, r As Long, found As Long, row As Long
    Set ws = KC_Sheet("_KC_Forms")
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value2) = KC_FormKey(raw) Then
            If found > 0 Then Err.Raise vbObjectError + 702, , "Doppelte Sonderkost-Freigabe"
            found = r: row = CLng(ws.Cells(r, 3).Value2)
        End If
    Next r
    If found = 0 Then Exit Function
    If row < 29 Or row > LAST_FORM_ROW Or ws.Cells(found, 6).Value2 <> "AKTIV" Then Err.Raise vbObjectError + 703, , "Sonderkost-Freigabe beschädigt"
    If KC_FormKey(KC_Sheet("Sonderkostformen").Cells(row, 37).Value2) <> KC_FormKey(raw) Then Err.Raise vbObjectError + 704, , "Bestätigter Sonderkost-Wortlaut wurde in der Matrix verändert"
    KC_FormRow = row
End Function

Public Function KC_FormLastRow() As Long
    Dim ws As Worksheet, r As Long, row As Long
    KC_FormLastRow = 28: Set ws = KC_Sheet("_KC_Forms")
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        row = KC_FormRow(CStr(ws.Cells(r, 2).Value2))
        If row > KC_FormLastRow Then KC_FormLastRow = row
    Next r
End Function

Public Function KC_FormFreeRow() As Long
    Dim ws As Worksheet, r As Long, cell As Range, available As Boolean
    Set ws = KC_Sheet("Sonderkostformen")
    If ws.ProtectContents Then Exit Function
    For r = 29 To LAST_FORM_ROW
        available = True
        For Each cell In ws.Range(ws.Cells(r, 37), ws.Cells(r, 51))
            If cell.HasFormula Or cell.MergeCells Or Not IsEmpty(cell.Value2) Then available = False: Exit For
        Next cell
        If available Then KC_FormFreeRow = r: Exit Function
    Next r
End Function

Private Sub Snapshot(ByVal cell As Range)
    Dim txn As Worksheet, r As Long
    Set txn = KC_Sheet("_KC_FormTxn"): r = txn.Cells(txn.Rows.Count, 1).End(xlUp).Row + 1
    txn.Cells(r, 1).Value2 = cell.Parent.Name: txn.Cells(r, 2).Value2 = cell.Address(False, False)
    If Not IsEmpty(cell.Value2) Or cell.HasFormula Then txn.Cells(r, 3).Value2 = CStr(cell.Formula)
    txn.Cells(r, 4).Value2 = CBool(cell.HasFormula): txn.Cells(r, 5).Value2 = VarType(cell.Value2)
End Sub

Private Sub UpdateTotals()
    Dim forms As Worksheet, matrix As Worksheet, r As Long, c As Long, cr As Long, b As Long
    Dim label As String, expr As String, original As String, col As Long, cell As Range
    Set forms = KC_Sheet("_KC_Forms"): Set matrix = KC_Sheet("Sonderkostformen")
    For c = 38 To 50
        label = CStr(matrix.Cells(4, c).Value2): cr = 0
        For r = 5 To 23
            If CStr(matrix.Cells(r, 1).Value2) = label Then
                If cr > 0 Then Err.Raise vbObjectError + 705, , "Matrix-Kundenzeile doppelt"
                cr = r
            End If
        Next r
        If cr > 0 Then
            b = c - 36
            If IsEmpty(forms.Cells(b, 10).Value2) Then
                forms.Cells(b, 10).Value2 = cr
                forms.Cells(b, 11).Value2 = CStr(matrix.Cells(cr, 31).Formula)
                forms.Cells(b, 12).Value2 = CStr(matrix.Cells(cr, 32).Formula)
            ElseIf CLng(forms.Cells(b, 10).Value2) <> cr Then
                Err.Raise vbObjectError + 706, , "Matrix-Kundenzeile wurde verschoben"
            End If
            expr = ""
            For r = 2 To forms.Cells(forms.Rows.Count, 1).End(xlUp).Row
                If forms.Cells(r, 6).Value2 = "AKTIV" Then
                    If Len(expr) > 0 Then expr = expr & ","
                    expr = expr & matrix.Cells(CLng(forms.Cells(r, 3).Value2), c).Address(False, False)
                End If
            Next r
            If Len(expr) > 0 Then
                For col = 31 To 32
                    Set cell = matrix.Cells(cr, col)
                    If cell.MergeCells Then Err.Raise vbObjectError + 707, , "Sonderkost-Summe verbunden"
                    original = CStr(forms.Cells(b, col - 20).Value2)
                    If original = "" Then original = "0"
                    If Left(original, 1) = "=" Then original = Mid(original, 2)
                    If Not cell.HasFormula And Not IsEmpty(cell.Value2) Then
                        If Not IsNumeric(cell.Value2) Then Err.Raise vbObjectError + 708, , "Sonderkost-Summe nicht numerisch"
                    End If
                    cell.Formula = "=(" & original & ")+SUM(" & expr & ")"
                Next col
            End If
        End If
    Next c
End Sub

Private Sub ValidateTotals()
    Dim forms As Worksheet, matrix As Worksheet, r As Long, c As Long, cr As Long, b As Long
    Dim label As String, expr As String, original As String, col As Long, cell As Range
    Set forms = KC_Sheet("_KC_Forms"): Set matrix = KC_Sheet("Sonderkostformen")
    For c = 38 To 50
        label = CStr(matrix.Cells(4, c).Value2): cr = 0
        For r = 5 To 23
            If CStr(matrix.Cells(r, 1).Value2) = label Then
                If cr > 0 Then Err.Raise vbObjectError + 705, , "Matrix-Kundenzeile doppelt"
                cr = r
            End If
        Next r
        If cr > 0 Then
            b = c - 36
            If Not IsEmpty(forms.Cells(b, 10).Value2) Then
                If CLng(forms.Cells(b, 10).Value2) <> cr Then Err.Raise vbObjectError + 706, , "Matrix-Kundenzeile wurde verschoben"
            expr = ""
            For r = 2 To forms.Cells(forms.Rows.Count, 1).End(xlUp).Row
                If forms.Cells(r, 6).Value2 = "AKTIV" Then
                    If Len(expr) > 0 Then expr = expr & ","
                    expr = expr & matrix.Cells(CLng(forms.Cells(r, 3).Value2), c).Address(False, False)
                End If
            Next r
            If Len(expr) > 0 Then
                For col = 31 To 32
                    Set cell = matrix.Cells(cr, col)
                    original = CStr(forms.Cells(b, col - 20).Value2)
                    If original = "" Then original = "0"
                    If Left(original, 1) = "=" Then original = Mid(original, 2)
                    If CStr(cell.Formula) <> "=(" & original & ")+SUM(" & expr & ")" Then Err.Raise vbObjectError + 715, , "Sonderkost-Summe wurde nach der letzten Freigabe manuell verändert: " & cell.Address
                Next col
            End If
            End If
        End If
    Next c
End Sub

' Called only after explicit manual confirmation; never from the monitor.
Public Function KC_FormRegister(ByVal raw As String, ByVal confirmed As Boolean, ByRef reason As String) As Boolean
    Dim row As Long, rr As Long, r As Long, col As Long, ws As Worksheet, forms As Worksheet
    Dim previousEvents As Boolean, started As Boolean, errText As String, cell As Range, savedJournal As Variant, lastJournal As Long
    Dim kw As Worksheet, day As Long, hcol As Long, cr As Long, customer As String, header As String, audit As Long
    reason = ""
    If Not confirmed Then reason = "Nicht bestätigt": Exit Function
    If mRegistering Or ThisWorkbook.ReadOnly Then reason = "Mappe nicht beschreibbar": Exit Function
    On Error GoTo Failed
    raw = CStr(KC_V20.FlatText(raw))
    If Len(raw) = 0 Or Len(raw) > 255 Or Left(raw, 1) = "=" Then Err.Raise vbObjectError + 709, , "Ungültiger Sonderkost-Wortlaut"
    row = KC_FormRow(raw)
    If row > 0 Then KC_FormRegister = True: Exit Function
    ValidateTotals
    r = KC_FormLastRow()
    row = KC_FormFreeRow()
    If row = 0 Then Err.Raise vbObjectError + 710, , "Keine freie, beschreibbare Matrixzeile zwischen AK29 und AY200"
    If KC_Sheet("_KC_Txn").Cells(KC_Sheet("_KC_Txn").Rows.Count, 1).End(xlUp).Row > 1 Then Err.Raise vbObjectError + 711, , "Offene Verbuchung zuerst wiederherstellen"
    If KC_Sheet("_KC_FormTxn").Cells(KC_Sheet("_KC_FormTxn").Rows.Count, 1).End(xlUp).Row > 1 Then Err.Raise vbObjectError + 712, , "Offene Sonderkost-Freigabe zuerst wiederherstellen"
    mRegistering = True: previousEvents = Application.EnableEvents: Application.EnableEvents = False
    Set ws = KC_Sheet("Sonderkostformen"): Set forms = KC_Sheet("_KC_Forms")
    rr = forms.Cells(forms.Rows.Count, 1).End(xlUp).Row + 1
    For Each cell In forms.Range("A" & rr & ":F" & rr): Snapshot cell: Next cell
    For Each cell In forms.Range("J2:L14"): Snapshot cell: Next cell
    Snapshot ws.Cells(row, 37): Snapshot ws.Cells(row, 51)
    For r = 5 To 23
        Snapshot ws.Cells(r, 31): Snapshot ws.Cells(r, 32)
    Next r
    For Each kw In ThisWorkbook.Worksheets
        If KC_V20.RxTest(kw.Name, "^KW\s*\d+$") Then
            For col = 38 To 50
                customer = CStr(ws.Cells(4, col).Value2): header = customer
                If header = "Liedbach1" Then header = "Liedbach"
                hcol = KC_V20.HeaderCol(kw, header): cr = 0
                For r = 5 To 23
                    If CStr(ws.Cells(r, 1).Value2) = customer Then cr = r
                Next r
                If hcol > 0 And cr > 0 Then
                    For day = 0 To 4
                        Set cell = kw.Cells(16 + day * 13, hcol)
                        If Not cell.HasFormula Or cell.MergeCells Or (kw.ProtectContents And cell.Locked) Then Err.Raise vbObjectError + 714, , "Manuelle Sonderkost-Summe nicht als freie Formel verfügbar: " & kw.Name & "!" & cell.Address
                        Snapshot cell
                    Next day
                End If
            Next col
        End If
    Next kw
    audit = KC_Sheet("_KC_Log").Cells(KC_Sheet("_KC_Log").Rows.Count, 1).End(xlUp).Row + 1
    For Each cell In KC_Sheet("_KC_Log").Range("A" & audit & ":I" & audit): Snapshot cell: Next cell
    lastJournal = KC_Sheet("_KC_FormTxn").Cells(KC_Sheet("_KC_FormTxn").Rows.Count, 1).End(xlUp).Row
    savedJournal = KC_Sheet("_KC_FormTxn").Range("A2:E" & lastJournal).Value2
    started = True: ThisWorkbook.Save
    ws.Cells(row, 37).Value2 = "'" & raw
    ws.Cells(row, 37).WrapText = True: ws.Rows(row).AutoFit
    ws.Cells(row, 51).Formula = "=SUM(AL" & row & ":AX" & row & ")"
    forms.Cells(rr, 1).Value2 = KC_FormKey(raw): forms.Cells(rr, 2).Value2 = raw
    forms.Cells(rr, 3).Value2 = row: forms.Cells(rr, 4).Value = Now
    forms.Cells(rr, 5).Value2 = Application.UserName: forms.Cells(rr, 6).Value2 = "AKTIV"
    UpdateTotals
    For Each kw In ThisWorkbook.Worksheets
        If KC_V20.RxTest(kw.Name, "^KW\s*\d+$") Then
            For col = 38 To 50
                customer = CStr(ws.Cells(4, col).Value2): header = customer
                If header = "Liedbach1" Then header = "Liedbach"
                hcol = KC_V20.HeaderCol(kw, header): cr = 0
                For r = 5 To 23
                    If CStr(ws.Cells(r, 1).Value2) = customer Then cr = r
                Next r
                If hcol > 0 And cr > 0 Then
                    For day = 0 To 4
                        kw.Cells(16 + day * 13, hcol).Formula = "=IF(Produktionsplan!$L$4=$A$8+" & day & ",Sonderkostformen!$AF$" & cr & ",""!"")"
                    Next day
                End If
            Next col
        End If
    Next kw
    KC_Log 0, "SONDERKOST-FREIGABE", "Manuell bestätigt: " & raw & "; Matrixzeile " & row
    ws.Calculate
    For r = 2 To 14
        If Not IsEmpty(forms.Cells(r, 10).Value2) Then
            For col = 31 To 32
                Set cell = ws.Cells(CLng(forms.Cells(r, 10).Value2), col)
                If IsError(cell.Value2) Then Err.Raise vbObjectError + 716, , "Fehler in neuer Sonderkost-Summenverknüpfung"
                If Not IsNumeric(cell.Value2) Then Err.Raise vbObjectError + 717, , "Neue Sonderkost-Summe ist nicht numerisch"
            Next col
        End If
    Next r
    If KC_FormRow(raw) <> row Or ws.Cells(row, 37).HasFormula Then Err.Raise vbObjectError + 713, , "Sonderkost-Freigabe konnte nicht rückgelesen werden"
    ThisWorkbook.Save
    KC_Sheet("_KC_FormTxn").Rows("2:" & KC_Sheet("_KC_FormTxn").Cells(KC_Sheet("_KC_FormTxn").Rows.Count, 1).End(xlUp).Row).ClearContents
    ThisWorkbook.Save
    started = False
    KC_FormRegister = True
Done:
    If mRegistering Then Application.EnableEvents = previousEvents
    mRegistering = False
    Exit Function
Failed:
    errText = Err.Description
    On Error Resume Next
    If started Then
        KC_Sheet("_KC_FormTxn").Range("A2:E" & lastJournal).Value2 = savedJournal
        KC_FormRecover
    End If
    reason = errText
    On Error GoTo 0
    GoTo Done
End Function

Public Function KC_FormAsk(e As Object) As Boolean
    Dim unknown As Collection, item As Variant, seen As Object, key As String, reason As String, row As Long
    Set unknown = e("unknown"): Set seen = CreateObject("Scripting.Dictionary")
    On Error GoTo Failed
    For Each item In unknown
        key = KC_FormKey(CStr(item(0)))
        If Not seen.Exists(key) Then
            seen.Add key, True
            row = KC_FormFreeRow()
            If row = 0 Then MsgBox "Keine freie Sonderkost-Matrixzeile verfügbar. Die Bestellung bleibt gesperrt.", vbExclamation, "KitaFino": Exit Function
            If MsgBox("Unbekannte Sonderkostform gefunden. Soll ich sie in die Sonderkostform-Matrix übernehmen?" & vbCrLf & vbCrLf & CStr(item(0)) & vbCrLf & "Kunde: " & CStr(item(1)) & vbCrLf & "Lieferdatum: " & Format(CDate(item(2)), "dd.mm.yyyy") & vbCrLf & "Menge: " & item(3) & vbCrLf & "Neue Matrixzeile: " & row & vbCrLf & vbCrLf & "Damit bestätigst du den vollständigen Wortlaut einschließlich aller Kombinationen. Anschließend wird die Bestellung erneut geprüft.", vbYesNo + vbQuestion + vbDefaultButton2, "KitaFino – Sonderkostform bestätigen") <> vbYes Then Exit Function
            If Not KC_FormRegister(CStr(item(0)), True, reason) Then MsgBox "Sonderkostform nicht übernommen: " & reason, vbExclamation, "KitaFino": Exit Function
        End If
    Next item
    KC_FormAsk = True
    Exit Function
Failed:
    MsgBox "Sonderkost-Freigabe gesperrt: " & Err.Description, vbExclamation, "KitaFino"
End Function

