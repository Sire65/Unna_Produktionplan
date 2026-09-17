Attribute VB_Name = "KC_Autark"
Option Explicit
Public Const KC_READY As String = "GEPRÜFT – BEREIT"
Public Const KC_WARN As String = "WARNUNG – MANUELL PRÜFEN"
Public Const KC_ERROR As String = "FEHLER – GESPERRT"
Public Const KC_BOOKED As String = "VERBUCHT"
Private mNext As Date, mScheduled As Boolean, mRunning As Boolean, mBusy As Boolean
Private mCallback As String
Private mScanItems As Object, mScanIndex As Long

Public Function KC_Sheet(ByVal name As String) As Worksheet
    Set KC_Sheet = ThisWorkbook.Worksheets(name)
End Function

Private Function EnsureSheet(ByVal name As String, ByVal hidden As Boolean) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.name = name
    End If
    If hidden Then EnsureSheet.Visible = xlSheetVeryHidden
End Function

Public Sub KC_Initialize()
    Dim ws As Worksheet, hdr As Variant, i As Long, button As Shape
    Set ws = EnsureSheet("_KC_Config", True)
    If ws.Cells(1, 1).Value <> "Version" Then
        ws.Range("A1:B8").NumberFormat = "@"
        ws.Cells(1, 1) = "Version": ws.Cells(1, 2) = "1.0-dev"
        ws.Cells(2, 1) = "Auto": ws.Cells(2, 2) = "AUS"
        ws.Cells(3, 1) = "Überwachung": ws.Cells(3, 2) = "EIN"
        ws.Cells(4, 1) = "Intervall Sekunden": ws.Cells(4, 2) = "60"
        ws.Cells(5, 1) = "Outlook Postfach": ws.Cells(5, 2) = "Mensa"
        ws.Cells(6, 1) = "Rückblick Tage": ws.Cells(6, 2) = "30"
        ws.Cells(7, 1) = "Letzter Scan": ws.Cells(7, 2) = "Noch nicht geprüft"
        ws.Cells(8, 1) = "Letzter Fehler": ws.Cells(8, 2) = ""
    End If
    Set ws = EnsureSheet("_KC_Dated", True)
    Set ws = EnsureSheet("_KC_Q", True)
    hdr = Array("Queue-ID", "Empfangen", "EntryID", "StoreID", "Message-ID", "Absender", "Betreff", "KitaFino-ID", "Kunde", "Lieferdatum", "Bis", "Status", "Ergebnis/Fehler", "Verbucht am", "Bestellart", "Sonderkost", "Formen", "Textteile", "Gesendet")
    For i = 0 To UBound(hdr): ws.Cells(1, i + 1) = hdr(i): Next i
    ws.Columns("C:I").NumberFormat = "@": ws.Columns("L:M").NumberFormat = "@"
    ws.Columns("O:P").NumberFormat = "@"
    Set ws = EnsureSheet("_KC_Body", True)
    ws.Range("A1:C1").Value = Array("Queue-ID", "Teil", "Mailtext")
    ws.Columns(3).NumberFormat = "@"
    Set ws = EnsureSheet("_KC_Log", True)
    ws.Range("A1:I1").Value = Array("Zeit", "Queue-ID", "Aktion", "Status", "Ergebnis", "Absender", "Betreff", "Mail-ID", "Verbuchungszeit")
    ws.Columns("C:H").NumberFormat = "@"
    Set ws = EnsureSheet("_KC_Txn", True)
    ws.Range("A1:F1").Value = Array("Queue-ID", "Blatt", "Zelle", "Formel/Wert", "Neuer Wert", "War Formel")
    ws.Columns("B:D").NumberFormat = "@"
    Set ws = EnsureSheet("KitaFino_Warteschlange", False)
    If ws.Cells(1, 1).Value = "" Then
        ws.Range("A1").Value = "KitaFino prüfen und verbuchen"
        ws.Range("A1").Font.Size = 16
        ws.Range("A3").Value = "Auto-Modus": ws.Range("A4").Value = "Outlook-Überwachung"
        ws.Range("A5").Value = "Postfach": ws.Range("A6").Value = "Letzte Prüfung"
        ws.Range("A7").Value = "Fehler": ws.Range("A8").Value = "Sonderkost-Plantag"
        ws.Range("A9:K9").Value = Array("Queue-ID", "Empfangen", "Kunde", "Von", "Bis", "Status", "Menge", "Ergebnis", "Sonderkost", "Betreff", "Verbucht am")
        ws.Range("A9:K9").Interior.Color = RGB(40, 65, 80)
        ws.Range("A9:K9").Font.Color = vbWhite: ws.Range("A9:K9").Font.Bold = True
        ws.Range("A9:K9").WrapText = True: ws.Rows(9).RowHeight = 30
        ws.Columns("A").ColumnWidth = 10: ws.Columns("B").ColumnWidth = 20
        ws.Columns("C").ColumnWidth = 28: ws.Columns("D:E").ColumnWidth = 12
        ws.Columns("F").ColumnWidth = 32: ws.Columns("G").ColumnWidth = 9
        ws.Columns("H").ColumnWidth = 45: ws.Columns("I").ColumnWidth = 35
        ws.Columns("J").ColumnWidth = 45: ws.Columns("K").ColumnWidth = 20
        ws.Columns("H:J").NumberFormat = "@": ws.Columns("H:J").WrapText = True
        ws.Range("B5").NumberFormat = "@": ws.Range("B5").Locked = False
        AddButton ws, "kcScan", "Jetzt prüfen", "KC_CheckNow", 350, 40
        AddButton ws, "kcBook", "Markierte Bestellung verbuchen", "KC_BookSelected", 520, 40
        AddButton ws, "kcAuto", "Auto AUS/EIN", "KC_ToggleAuto", 800, 40
        AddButton ws, "kcWatch", "Überwachung AUS/EIN", "KC_ToggleWatch", 970, 40
        AddButton ws, "kcPick", "Outlook-Auswahl einlesen", "KC_CaptureSelection", 350, 80
        AddButton ws, "kcDetails", "Mail und Zielwerte anzeigen", "KC_ShowSelected", 600, 80
        AddButton ws, "kcDay", "Sonderkost für Plantag", "KC_ProjectManual", 350, 120
        AddButton ws, "kcFile", "MSG-Dateien einlesen", "KC_CaptureFiles", 870, 80
    End If
    ws.Unprotect
    ws.Range("A3") = "Auto": ws.Range("A4") = "Outlook"
    ws.Range("A6") = "Prüfung": ws.Range("A8") = "Plantag"
    ws.Rows(1).RowHeight = 26: ws.Rows(2).RowHeight = 18
    ws.Rows("3:8").RowHeight = 23
    ws.Shapes("kcBook").Width = 250
    ws.Shapes("kcPick").Width = 230
    ws.Shapes("kcDetails").Width = 230
    ws.Shapes("kcDay").Width = 240
    For Each button In ws.Shapes
        If Left(button.Name, 2) = "kc" Then
            button.Placement = xlFreeFloating: button.Height = 30
            Select Case button.Name
                Case "kcScan", "kcBook", "kcAuto", "kcWatch": button.Top = 40
                Case "kcPick", "kcDetails", "kcFile": button.Top = 80
                Case "kcDay": button.Top = 120
            End Select
        End If
        If Left(button.Name, 2) = "kc" And InStr(button.OnAction, "!") > 0 Then
            button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & Mid(button.OnAction, InStrRev(button.OnAction, "!") + 1)
        End If
    Next button
    KC_Refresh
End Sub

Private Sub AddButton(ws As Worksheet, name As String, caption As String, macro As String, x As Long, y As Long)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(5, x, y, 160, 30)
    s.name = name: s.TextFrame.Characters.Text = caption
    s.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & macro
End Sub

Public Sub KC_Log(ByVal queueRow As Long, ByVal action As String, ByVal note As String)
    Dim ws As Worksheet, q As Worksheet, r As Long, j As Long
    Set ws = KC_Sheet("_KC_Log"): Set q = KC_Sheet("_KC_Q")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    ws.Cells(r, 1) = Now: ws.Cells(r, 3) = action: ws.Cells(r, 5) = Left(note, 32000)
    If queueRow > 1 Then
        ws.Cells(r, 2) = q.Cells(queueRow, 1).Value
        ws.Cells(r, 4) = q.Cells(queueRow, 12).Value
        ws.Cells(r, 6) = q.Cells(queueRow, 6).Value: ws.Cells(r, 7) = q.Cells(queueRow, 7).Value
        ws.Cells(r, 8) = q.Cells(queueRow, 5).Value: ws.Cells(r, 9) = q.Cells(queueRow, 14).Value
    End If
End Sub

Public Function KC_Body(ByVal id As Long) As String
    Dim ws As Worksheet, r As Long
    Set ws = KC_Sheet("_KC_Body")
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        If ws.Cells(r, 1).Value = id Then KC_Body = KC_Body & CStr(ws.Cells(r, 3).Value)
    Next r
End Function

Public Function KC_Enqueue(ByVal subject As String, ByVal body As String, ByVal entryId As String, ByVal storeId As String, ByVal messageId As String, ByVal sender As String, ByVal received As Date, ByVal sent As Date) As Long
    Dim q As Worksheet, b As Worksheet, r As Long, j As Long, br As Long, id As Long, p As Long
    Dim evaluation As Object, existing As Boolean
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 610, , "Arbeitsmappe schreibgeschützt"
    Set q = KC_Sheet("_KC_Q"): Set b = KC_Sheet("_KC_Body")
    For j = 2 To q.Cells(q.Rows.Count, 1).End(xlUp).row
        existing = False
        If Len(messageId) > 0 Then existing = (CStr(q.Cells(j, 5).Value) = messageId)
        If Len(entryId) > 0 Then
            If CStr(q.Cells(j, 3).Value) = entryId And CStr(q.Cells(j, 4).Value) = storeId Then existing = True
        End If
        If CStr(q.Cells(j, 7).Value) = subject Then
            If KC_Body(CLng(q.Cells(j, 1).Value)) = body Then existing = True
        End If
        If existing Then KC_Enqueue = j: Exit Function
    Next j
    r = q.Cells(q.Rows.Count, 1).End(xlUp).row + 1
    id = r - 1
    q.Cells(r, 1) = id: q.Cells(r, 2) = received
    q.Cells(r, 3) = entryId: q.Cells(r, 4) = storeId: q.Cells(r, 5) = messageId
    q.Cells(r, 6) = sender: q.Cells(r, 7) = subject: q.Cells(r, 19) = sent
    q.Cells(r, 12) = KC_ERROR: q.Cells(r, 13) = "Einlesen noch nicht abgeschlossen"
    For p = 1 To Len(body) Step 30000
        br = b.Cells(b.Rows.Count, 1).End(xlUp).row + 1
        b.Cells(br, 1) = id: b.Cells(br, 2) = ((p - 1) \ 30000) + 1
        b.Cells(br, 3) = Mid(body, p, 30000)
    Next p
    q.Cells(r, 18) = (Len(body) + 29999) \ 30000
    Set evaluation = KC_Evaluate(subject, body)
    KC_StoreEvaluation r, evaluation
    KC_CheckHistory r, evaluation
    KC_Log r, "EINGELESEN", CStr(q.Cells(r, 13).Value)
    KC_Enqueue = r
End Function

Public Sub KC_StoreEvaluation(ByVal r As Long, e As Object)
    Dim q As Worksheet
    Set q = KC_Sheet("_KC_Q")
    q.Cells(r, 8) = e("id"): q.Cells(r, 9) = e("customer")
    q.Cells(r, 10) = e("date"): q.Cells(r, 11) = e("end")
    q.Cells(r, 12) = e("status"): q.Cells(r, 13) = Left(CStr(e("note")), 32000)
    q.Cells(r, 15) = e("type"): q.Cells(r, 16) = Left(CStr(e("special")), 32000)
    q.Cells(r, 17) = e("forms")
End Sub

Public Sub KC_CheckHistory(ByVal r As Long, e As Object)
    Dim q As Worksheet, log As Worksheet, j As Long, d As Variant, lastD As Variant
    Dim mid As String, eid As String, bookedKey As String, oldDate As Variant
    Set q = KC_Sheet("_KC_Q"): Set log = KC_Sheet("Import_Protokoll")
    mid = CStr(q.Cells(r, 5).Value): eid = CStr(q.Cells(r, 3).Value)
    For j = 2 To log.Cells(log.Rows.Count, 1).End(xlUp).row
        If UCase(CStr(log.Cells(j, 15).Value)) = "JA" Then
            bookedKey = CStr(log.Cells(j, 6).Value)
            If (Len(mid) > 0 And bookedKey = mid) Or (Len(eid) > 0 And bookedKey = eid) Then
                q.Cells(r, 12) = KC_ERROR: q.Cells(r, 13) = "Dublette: bereits im bisherigen Protokoll verbucht": Exit Sub
            End If
            oldDate = log.Cells(j, 8).Value
            If CStr(log.Cells(j, 7).Value) = CStr(e("id")) And IsDate(oldDate) And IsDate(e("date")) And IsDate(e("end")) Then
                If CDate(oldDate) >= CDate(e("date")) And CDate(oldDate) <= CDate(e("end")) Then
                    If e("status") <> KC_ERROR Then q.Cells(r, 12) = KC_WARN: q.Cells(r, 13) = KC_V20.AddN(q.Cells(r, 13).Value, "Bereits vorhandene historische Verbuchung für diesen Zeitraum")
                End If
            End If
        End If
    Next j
    If Not IsDate(e("date")) Or Not IsDate(e("end")) Then Exit Sub
    For j = 2 To q.Cells(q.Rows.Count, 1).End(xlUp).row
        If j <> r And CStr(q.Cells(j, 8).Value) = CStr(e("id")) Then
            d = q.Cells(j, 10).Value: lastD = q.Cells(j, 11).Value
            If IsDate(d) And IsDate(lastD) Then
                If CDate(d) <= CDate(e("end")) And CDate(lastD) >= CDate(e("date")) Then
                    If q.Cells(j, 12).Value <> KC_ERROR Then
                        If CDbl(q.Cells(j, 2).Value) >= CDbl(q.Cells(r, 2).Value) Then
                            q.Cells(r, 12) = KC_ERROR: q.Cells(r, 13) = "Ältere oder zeitgleiche Bestellung für überlappende Liefertage": Exit Sub
                        ElseIf e("status") <> KC_ERROR Then
                            q.Cells(r, 12) = KC_WARN: q.Cells(r, 13) = KC_V20.AddN(q.Cells(r, 13).Value, "Neuere Bestellung: überlappenden Vorgänger manuell prüfen")
                        End If
                    End If
                End If
            End If
        End If
    Next j
End Sub

Public Sub KC_Refresh()
    Dim ws As Worksheet, q As Worksheet, c As Worksheet, r As Long, v As Long, n As Long
    Set ws = KC_Sheet("KitaFino_Warteschlange"): Set q = KC_Sheet("_KC_Q"): Set c = KC_Sheet("_KC_Config")
    ws.Unprotect
    ws.Range("B3") = c.Cells(2, 2).Value: ws.Range("B4") = c.Cells(3, 2).Value
    ws.Range("B5") = c.Cells(5, 2).Value: ws.Range("B6") = c.Cells(7, 2).Value
    ws.Range("B7") = c.Cells(8, 2).Value
    ws.Range("B8").Value = KC_Sheet("_KC_Config").Cells(9, 2).Value
    n = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If n >= 10 Then ws.Range("A10:K" & n).ClearContents
    For r = 2 To q.Cells(q.Rows.Count, 1).End(xlUp).row
        v = r + 8
        ws.Cells(v, 1) = q.Cells(r, 1).Value: ws.Cells(v, 2) = q.Cells(r, 2).Value
        ws.Cells(v, 3) = q.Cells(r, 9).Value: ws.Cells(v, 4) = q.Cells(r, 10).Value
        ws.Cells(v, 5) = q.Cells(r, 11).Value: ws.Cells(v, 6) = q.Cells(r, 12).Value
        ws.Cells(v, 7) = KC_V20.Rx1(CStr(q.Cells(r, 7).Value), "kitafino\s*-\s*(\d+)\s+Bestellungen")
        ws.Cells(v, 8) = Left(CStr(q.Cells(r, 13).Value), 220): ws.Cells(v, 9) = Left(CStr(q.Cells(r, 16).Value), 220)
        ws.Cells(v, 10) = q.Cells(r, 7).Value: ws.Cells(v, 11) = q.Cells(r, 14).Value
        ws.Rows(v).RowHeight = 65
        Select Case CStr(q.Cells(r, 12).Value)
            Case KC_READY, KC_BOOKED: ws.Cells(v, 6).Interior.Color = RGB(198, 239, 206)
            Case KC_WARN: ws.Cells(v, 6).Interior.Color = RGB(255, 235, 156)
            Case Else: ws.Cells(v, 6).Interior.Color = RGB(255, 199, 206)
        End Select
    Next r
    ws.Columns("B").NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Columns("D:E").NumberFormat = "dd.mm.yyyy": ws.Columns("K").NumberFormat = "dd.mm.yyyy hh:mm"
    ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Public Sub KC_OnOpen()
    On Error GoTo Failed
    KC_Initialize
    KC_Recover
    If KC_Sheet("_KC_Config").Cells(3, 2).Value = "EIN" Then KC_Start
    Exit Sub
Failed:
    ' Importfehler dürfen die bisherigen Workbook-Ereignisse nicht abbrechen.
    On Error Resume Next
    KC_Sheet("_KC_Config").Cells(8, 2) = Err.Description
End Sub

Private Function CallbackName() As String
    CallbackName = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!KC_Tick"
End Function

Public Sub KC_Start()
    KC_Stop
    If ThisWorkbook.ReadOnly Then Exit Sub
    mRunning = True
    Schedule 5
End Sub

Public Sub KC_Stop()
    mRunning = False
    On Error Resume Next
    If mScheduled Then Application.OnTime EarliestTime:=mNext, Procedure:=mCallback, Schedule:=False
    mScheduled = False: Set mScanItems = Nothing: mScanIndex = 0
    On Error GoTo 0
End Sub

Private Sub Schedule(ByVal seconds As Long)
    If Not mRunning Then Exit Sub
    mNext = Now + TimeSerial(0, 0, seconds)
    mCallback = CallbackName()
    Application.OnTime EarliestTime:=mNext, Procedure:=mCallback
    mScheduled = True
End Sub

Public Sub KC_Tick()
    mScheduled = False
    If Not mRunning Then Exit Sub
    KC_CheckNow
    KC_ProjectDay False
    Schedule CLng(Application.Min(3600, Application.Max(10, Val(KC_Sheet("_KC_Config").Cells(4, 2).Value))))
End Sub

Private Function OutlookApp() As Object
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If OutlookApp Is Nothing Then Err.Raise vbObjectError + 620, , "Outlook nicht erreichbar; klassisches Outlook erforderlich"
End Function

Public Sub KC_CheckNow()
    Dim ol As Object, ns As Object, store As Object, folder As Object, mail As Object, c As Worksheet
    Dim query As String, utc As Date, count As Long, r As Long, t As Single, storeName As String
    If mBusy Or ThisWorkbook.ReadOnly Then Exit Sub
    mBusy = True
    On Error GoTo Failed
    Set c = KC_Sheet("_KC_Config")
    storeName = Trim(CStr(KC_Sheet("KitaFino_Warteschlange").Range("B5").Value))
    If storeName <> CStr(c.Cells(5, 2).Value) Then
        c.Cells(5, 2) = storeName: Set mScanItems = Nothing
    End If
    Set ol = OutlookApp(): Set ns = ol.GetNamespace("MAPI")
    If mScanItems Is Nothing Then
        For Each store In ns.Stores
            If StrComp(store.DisplayName, storeName, vbTextCompare) = 0 Then Set folder = store.GetDefaultFolder(6): Exit For
        Next store
        If folder Is Nothing Then Err.Raise vbObjectError + 621, , "Outlook-Postfach nicht eindeutig gefunden: " & storeName
        utc = folder.PropertyAccessor.LocalTimeToUTC(DateAdd("d", -CLng(Application.Max(1, Val(c.Cells(6, 2).Value))), Now))
        query = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%kitafino%' AND ""urn:schemas:httpmail:datereceived"" >= '" & Format(utc, "yyyy-mm-dd hh:nn") & "'"
        Set mScanItems = folder.Items.Restrict(query)
        mScanItems.Sort "[ReceivedTime]", True: mScanIndex = 1
    End If
    t = VBA.Timer
    Do While mScanIndex <= mScanItems.Count And count < 20
        Set mail = mScanItems.Item(mScanIndex): mScanIndex = mScanIndex + 1: count = count + 1
        If mail.Class = 43 Then
            If KC_V20.RxTest(CStr(mail.subject), "^\s*kitafino\s*-\s*\d+\s+Bestellungen") Then r = Capture(mail)
        End If
        If VBA.Timer - t > 3 Or VBA.Timer < t Then Exit Do
    Loop
    If mScanIndex > mScanItems.Count Then Set mScanItems = Nothing
    c.Cells(7, 2) = Format(Now, "dd.mm.yyyy hh:nn:ss"): c.Cells(8, 2) = ""
    KC_Refresh
    ThisWorkbook.Save
    If c.Cells(2, 2).Value = "EIN" Then KC_AutoBook
Done:
    mBusy = False
    Exit Sub
Failed:
    Dim errorText As String
    errorText = Err.Description
    On Error Resume Next
    KC_Sheet("_KC_Config").Cells(8, 2) = Left(errorText, 32000)
    KC_Log 0, "OUTLOOK/SCAN-FEHLER", errorText
    Set mScanItems = Nothing
    KC_Refresh
    Resume Done
End Sub

Private Function Capture(mail As Object) As Long
    Dim mid As String, sid As String
    If mail.Class <> 43 Then Exit Function
    On Error Resume Next
    mid = CStr(mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F"))
    If Len(mid) = 0 Then mid = CStr(mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001E"))
    sid = mail.Parent.StoreID
    On Error GoTo 0
    Capture = KC_Enqueue(CStr(mail.subject), CStr(mail.body), CStr(mail.entryId), sid, mid, CStr(mail.SenderEmailAddress), CDate(mail.ReceivedTime), CDate(mail.SentOn))
End Function

Public Sub KC_CaptureSelection()
    Dim ol As Object, selection As Object, mail As Object
    If mBusy Or ThisWorkbook.ReadOnly Then Exit Sub
    mBusy = True
    On Error GoTo Failed
    Set ol = OutlookApp(): Set selection = ol.ActiveExplorer.selection
    For Each mail In selection
        If mail.Class = 43 Then Capture mail
    Next mail
    KC_Refresh: ThisWorkbook.Save
    mBusy = False
    KC_Sheet("KitaFino_Warteschlange").Activate
    Exit Sub
Failed:
    mBusy = False
    MsgBox "Einlesen fehlgeschlagen: " & Err.Description, vbExclamation
End Sub

Public Sub KC_CaptureFiles()
    Dim fd As FileDialog, ol As Object, mail As Object, i As Long
    If mBusy Or ThisWorkbook.ReadOnly Then Exit Sub
    Set fd = Application.FileDialog(3)
    fd.AllowMultiSelect = True: fd.Filters.Clear: fd.Filters.Add "Outlook-Nachrichten", "*.msg"
    If fd.show <> -1 Then Exit Sub
    mBusy = True
    On Error GoTo Failed
    Set ol = OutlookApp()
    For i = 1 To fd.SelectedItems.Count
        Set mail = ol.Session.OpenSharedItem(fd.SelectedItems(i))
        If mail.Class = 43 Then Capture mail
    Next i
    KC_Refresh: ThisWorkbook.Save: mBusy = False
    KC_Sheet("KitaFino_Warteschlange").Activate
    Exit Sub
Failed:
    mBusy = False: MsgBox "MSG-Einlesen fehlgeschlagen: " & Err.Description, vbExclamation
End Sub

Public Sub KC_ToggleAuto()
    Dim c As Worksheet
    Set c = KC_Sheet("_KC_Config")
    If c.Cells(2, 2).Value = "EIN" Then
        c.Cells(2, 2) = "AUS"
    Else
        If MsgBox("Auto-Modus einschalten? Nur vollständig eindeutige, konfliktfreie Bestellungen werden verbucht.", vbYesNo + vbQuestion, "KitaFino") <> vbYes Then Exit Sub
        c.Cells(2, 2) = "EIN"
    End If
    KC_Log 0, "AUTO-MODUS", CStr(c.Cells(2, 2).Value)
    KC_Refresh: ThisWorkbook.Save
End Sub

Public Sub KC_ToggleWatch()
    Dim c As Worksheet
    Set c = KC_Sheet("_KC_Config")
    If c.Cells(3, 2).Value = "EIN" Then c.Cells(3, 2) = "AUS": KC_Stop Else c.Cells(3, 2) = "EIN": KC_Start
    KC_Refresh: ThisWorkbook.Save
End Sub

Public Sub KC_BookSelected()
    Dim id As Variant, q As Worksheet, r As Long
    If ActiveSheet.name <> "KitaFino_Warteschlange" Or ActiveCell.row < 10 Then Exit Sub
    id = ActiveSheet.Cells(ActiveCell.row, 1).Value
    Set q = KC_Sheet("_KC_Q")
    For r = 2 To q.Cells(q.Rows.Count, 1).End(xlUp).row
        If q.Cells(r, 1).Value = id Then KC_BookRow r, False: Exit For
    Next r
    KC_Refresh
End Sub

Public Sub KC_ShowSelected()
    Dim id As Variant, q As Worksheet, r As Long, e As Object, plan As Collection, item As Variant
    Dim ws As Worksheet, text As String, line As Long
    If ActiveSheet.name <> "KitaFino_Warteschlange" Or ActiveCell.row < 10 Then Exit Sub
    id = ActiveSheet.Cells(ActiveCell.row, 1).Value
    Set q = KC_Sheet("_KC_Q")
    For r = 2 To q.Cells(q.Rows.Count, 1).End(xlUp).row
        If q.Cells(r, 1).Value = id Then Exit For
    Next r
    If r > q.Cells(q.Rows.Count, 1).End(xlUp).row Then Exit Sub
    Set ws = EnsureSheet("KitaFino_Details", False)
    ws.Cells.ClearContents: ws.Columns("A:C").NumberFormat = "@"
    ws.Range("A1") = q.Cells(r, 7).Value
    Set e = KC_Evaluate(CStr(q.Cells(r, 7).Value), KC_Body(CLng(id)))
    ws.Range("A3") = e("status"): ws.Range("A4") = e("note")
    ws.Range("A6:C6").Value = Array("Zielblatt", "Zelle", "Geprüfter Wert")
    Set plan = e("plan"): line = 7
    For Each item In plan
        ws.Cells(line, 1) = item(0): ws.Cells(line, 2) = item(1): ws.Cells(line, 3) = item(2): line = line + 1
    Next item
    line = line + 2: text = KC_Body(CLng(id))
    Dim p As Long
    For p = 1 To Len(text) Step 30000
        ws.Cells(line, 1) = Mid(text, p, 30000): line = line + 1
    Next p
    ws.Columns("A").ColumnWidth = 105: ws.Columns("A").WrapText = True
    ws.Columns("B:C").ColumnWidth = 18
    ws.Rows.AutoFit: ws.Activate
End Sub
