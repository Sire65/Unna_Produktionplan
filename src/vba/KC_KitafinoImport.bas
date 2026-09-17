Attribute VB_Name = "KC_KitafinoImport"

Option Explicit



Public Sub KC_KitaFino_Start()
    KC_CaptureFiles
End Sub



Public Sub KCProto_Alle(): ResetProto: End Sub

Public Sub KCProto_Gruen(): ResetProto: ThisWorkbook.Worksheets("Import_Protokoll").Range("A1").CurrentRegion.AutoFilter Field:=16, Criteria1:="GRUEN": End Sub

Public Sub KCProto_Gelb(): ResetProto: ThisWorkbook.Worksheets("Import_Protokoll").Range("A1").CurrentRegion.AutoFilter Field:=16, Criteria1:="GELB": End Sub

Public Sub KCProto_Rot(): ResetProto: ThisWorkbook.Worksheets("Import_Protokoll").Range("A1").CurrentRegion.AutoFilter Field:=16, Criteria1:="ROT": End Sub

Public Sub KCProto_Heute(): ResetProto: ThisWorkbook.Worksheets("Import_Protokoll").Range("A1").CurrentRegion.AutoFilter Field:=2, Criteria1:=">=" & CLng(Date), Operator:=xlAnd, Criteria2:="<" & CLng(Date + 1): End Sub

Public Sub KCProto_7Tage(): ResetProto: ThisWorkbook.Worksheets("Import_Protokoll").Range("A1").CurrentRegion.AutoFilter Field:=2, Criteria1:=">=" & CLng(Date - 6): End Sub

Public Sub KCProto_Bereinigen()

    Dim ws As Worksheet, se As Worksheet, lr As Long, r As Long, cnt As Long

    Dim dg As Long, dy As Long, dr As Long, age As Long, st As String

    Set ws = ThisWorkbook.Worksheets("Import_Protokoll"): Set se = ThisWorkbook.Worksheets("Import_Einstellungen")

    dg = val(se.Range("B3").Value): dy = val(se.Range("B4").Value): dr = val(se.Range("B5").Value)

    lr = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    For r = 2 To lr

        If IsDate(ws.Cells(r, 2).Value) Then

            age = Date - DateValue(ws.Cells(r, 2).Value): st = UCase(Trim(CStr(ws.Cells(r, 16).Value)))

            If (st = "GRUEN" And age > dg) Or (st = "GELB" And age > dy) Or (st = "ROT" And age > dr) Then cnt = cnt + 1

        End If

    Next r

    If cnt = 0 Then MsgBox "Keine alten Protokolle zum Löschen gefunden.", vbInformation: Exit Sub

    If MsgBox("Es werden " & cnt & " alte Protokolle gelöscht. Fortfahren?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    For r = lr To 2 Step -1

        If IsDate(ws.Cells(r, 2).Value) Then

            age = Date - DateValue(ws.Cells(r, 2).Value): st = UCase(Trim(CStr(ws.Cells(r, 16).Value)))

            If (st = "GRUEN" And age > dg) Or (st = "GELB" And age > dy) Or (st = "ROT" And age > dr) Then ws.Rows(r).Delete

        End If

    Next r

    MsgBox cnt & " Protokolle gelöscht.", vbInformation

End Sub



Private Sub ResetProto(): On Error Resume Next: If ThisWorkbook.Worksheets("Import_Protokoll").FilterMode Then ThisWorkbook.Worksheets("Import_Protokoll").ShowAllData: On Error GoTo 0: End Sub

Private Function Rx1(t As String, p As String) As String: Dim r As Object, m As Object: Set r = CreateObject("VBScript.RegExp"): r.IgnoreCase = True: r.pattern = p: If r.Test(t) Then Set m = r.Execute(t)(0): Rx1 = m.SubMatches(0): End If: End Function

Private Function AddN(a As String, b As String) As String: If Trim(a) = "" Then AddN = b Else AddN = a & "; " & b: End Function

Private Function GermanDate(t As String) As Variant

    Dim r As Object, m As Object, d As Long, y As Long, mo As String, mn As Long

    Set r = CreateObject("VBScript.RegExp"): r.IgnoreCase = True

    r.pattern = "(\d{1,2})\.\s*(Januar|Februar|März|Maerz|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)\s*(\d{4})"

    If Not r.Test(t) Then GermanDate = "": Exit Function

    Set m = r.Execute(t)(0): d = CLng(m.SubMatches(0)): mo = LCase(m.SubMatches(1)): y = CLng(m.SubMatches(2))

    Select Case mo

        Case "januar": mn = 1

        Case "februar": mn = 2

        Case "märz", "maerz": mn = 3

        Case "april": mn = 4

        Case "mai": mn = 5

        Case "juni": mn = 6

        Case "juli": mn = 7

        Case "august": mn = 8

        Case "september": mn = 9

        Case "oktober": mn = 10

        Case "november": mn = 11

        Case "dezember": mn = 12

    End Select

    GermanDate = DateSerial(y, mn, d)

End Function

Private Function Institution(b As String, id As String) As String: Dim r As Object, m As Object, x As String: x = Replace(Replace(b, vbCr, " "), vbLf, " "): Set r = CreateObject("VBScript.RegExp"): r.IgnoreCase = True: r.pattern = "BESTELLUNGEN.*?für\s+(.+?)\s+\(" & id & "\)": If r.Test(x) Then Set m = r.Execute(x)(0): Institution = Trim(m.SubMatches(0)): End Function

Private Function KeyExists(ws As Worksheet, k As String) As Boolean: Dim f As Range: If Trim(k) = "" Then Exit Function: Set f = ws.Columns(6).Find(What:=k, LookIn:=xlValues, LookAt:=xlWhole): KeyExists = Not f Is Nothing: End Function

Private Function TargetID(id As String) As String: Select Case id: Case "59432": TargetID = "Cluster1-4": Case "59431": TargetID = "Liedbach1": Case "59500": TargetID = "Strolche": Case "59430": TargetID = "Lünern / Lünern2": Case "59426": TargetID = "Kita1 / Kita2_4": End Select: End Function

Private Function FindCol(ws As Worksheet, n As String) As Long: Dim c As Long: For c = 3 To 80: If Trim(CStr(ws.Cells(3, c).Value)) = n Then FindCol = c: Exit Function: Next c: End Function

Private Function RawRow(ws As Worksheet, dt As Variant) As Long: Dim ix As Long: If Not IsDate(dt) Or Not IsDate(ws.Range("A8").Value) Then Exit Function: ix = DateDiff("d", CDate(ws.Range("A8").Value), CDate(dt)): If ix >= 0 And ix <= 4 Then RawRow = 4 + ix * 13: End Function

Private Function WriteTot(ws As Worksheet, n As String, dt As Variant, qv As Long) As Boolean: Dim c As Long, r As Long: c = FindCol(ws, n): r = RawRow(ws, dt): If c > 0 And r > 0 Then ws.Cells(r, c).Value = qv: WriteTot = True: End Function

Private Function ReadVals(ws As Worksheet, a As Variant, dt As Variant) As String: Dim j As Long, c As Long, r As Long, z As String: r = RawRow(ws, dt): For j = LBound(a) To UBound(a): c = FindCol(ws, CStr(a(j))): If c > 0 Then If z <> "" Then z = z & "; ": z = z & CStr(a(j)) & "=" & CStr(ws.Cells(r, c).Value): Next j: ReadVals = z: End Function

Private Function GQty(b As String, g As String) As String: Dim r As Object, m As Object, x As String: x = Replace(Replace(Replace(b, vbCr, " "), vbLf, " "), vbTab, " "): Set r = CreateObject("VBScript.RegExp"): r.IgnoreCase = True: r.pattern = g & "\s+Anzahl.*?\s(\d{1,4})(?:\s+Davon|\s+Cluster|\s+Gruppe|\s+\[|$)": If r.Test(x) Then Set m = r.Execute(x)(0): GQty = m.SubMatches(0): End Function

Private Function ImportClustersV(ws As Worksheet, b As String, dt As Variant) As Boolean

    Dim q1 As String, q2 As String, q3 As String, q4 As String

    q1 = GQty(b, "Cluster\s*1"): q2 = GQty(b, "Cluster\s*2"): q3 = GQty(b, "Cluster\s*3"): q4 = GQty(b, "Cluster\s*4")

    If IsNumeric(q1) And IsNumeric(q2) And IsNumeric(q3) And IsNumeric(q4) Then

        If WriteTot(ws, "Cluster1", dt, CLng(q1)) And WriteTot(ws, "Cluster2", dt, CLng(q2)) And WriteTot(ws, "Cluster3", dt, CLng(q3)) And WriteTot(ws, "Cluster4", dt, CLng(q4)) Then ImportClustersV = True

    End If

End Function

Private Sub WriteCurrent(ws As Worksheet, r As Long, f As String, dt As Variant, kw As Long, id As String, ins As String, qv As String, st As String, n As String, tg As String, bk As String, ii As String): ws.Cells(r, 1) = f: ws.Cells(r, 2) = dt: ws.Cells(r, 3) = kw: ws.Cells(r, 4) = id: ws.Cells(r, 5) = ins: ws.Cells(r, 6) = qv: ws.Cells(r, 7) = st: ws.Cells(r, 8) = n: ws.Cells(r, 9) = tg: ws.Cells(r, 10) = bk: ws.Cells(r, 11) = ii: End Sub

Private Sub WriteLog(ws As Worksheet, r As Long, ii As String, u As String, f As String, k As String, id As String, dt As Variant, kw As Long, ins As String, tg As String, qv As String, bef As String, aft As String, bk As String, st As String, n As String): ws.Cells(r, 1) = ii: ws.Cells(r, 2) = Date: ws.Cells(r, 3) = Time: ws.Cells(r, 4) = u: ws.Cells(r, 5) = f: ws.Cells(r, 6) = k: ws.Cells(r, 7) = id: ws.Cells(r, 8) = dt: ws.Cells(r, 9) = kw: ws.Cells(r, 10) = ins: ws.Cells(r, 11) = tg: ws.Cells(r, 12) = "Gesamt=" & qv: ws.Cells(r, 13) = bef: ws.Cells(r, 14) = aft: ws.Cells(r, 15) = bk: ws.Cells(r, 16) = st: ws.Cells(r, 17) = n: ws.Cells(r, 18) = "MSG": End Sub

Private Sub ColorStatus(ws As Worksheet, c As Long, r1 As Long, r2 As Long): Dim r As Long, st As String: If r2 < r1 Then Exit Sub: For r = r1 To r2: st = UCase(Trim(CStr(ws.Cells(r, c).Value))): Select Case st: Case "GRUEN": ws.Cells(r, c).Interior.Color = RGB(198, 239, 206): Case "GELB": ws.Cells(r, c).Interior.Color = RGB(255, 235, 156): Case "ROT": ws.Cells(r, c).Interior.Color = RGB(255, 199, 206): End Select: Next r: End Sub



