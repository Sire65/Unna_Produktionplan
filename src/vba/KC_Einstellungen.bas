Attribute VB_Name = "KC_Einstellungen"
Option Explicit
Private Const INTERVAL_NAME As String = "KC_OutlookIntervallSekunden"

Public Sub KC_SetupInterval()
    Dim oldCell As Range, mailCell As Range, ws As Worksheet, seconds As Long, mailbox As String
    On Error Resume Next
    Set oldCell = ThisWorkbook.Names(INTERVAL_NAME).RefersToRange
    Set mailCell = ThisWorkbook.Names("KC_OutlookPostfach").RefersToRange
    On Error GoTo 0
    Set ws = KC_Sheet("Stammdaten")
    If Not oldCell Is Nothing And Not mailCell Is Nothing Then
        If oldCell.Address = "$Q$34" And mailCell.Address = "$Q$33" Then Exit Sub
    End If
    If ws.ProtectContents Then Err.Raise vbObjectError + 682, , "Stammdaten für erstmalige Outlook-Einrichtung geschützt"
    seconds = 60
    If ValidInterval(KC_Sheet("_KC_Config").Cells(4, 2).Value2) Then seconds = CLng(KC_Sheet("_KC_Config").Cells(4, 2).Value2)
    If Not oldCell Is Nothing Then
        If ValidInterval(oldCell.Value2) Then seconds = CLng(oldCell.Value2)
    End If
    mailbox = Trim(CStr(KC_Sheet("_KC_Config").Cells(5, 2).Value))
    If mailbox = "" Then mailbox = "Mensa"
    If IsEmpty(ws.Range("Q33").Value2) Then ws.Range("Q33").Value2 = mailbox
    If IsEmpty(ws.Range("Q34").Value2) Then ws.Range("Q34").Value2 = seconds
    ws.Range("M33") = "Outlook-Postfach"
    ws.Range("M34") = "Outlook-Abfrage alle"
    ws.Range("R34") = "Sekunden"
    ws.Range("Q33").NumberFormat = "@": ws.Range("Q34").NumberFormat = "0"
    ws.Range("Q33:Q34").Locked = False
    ws.Range("Q33:Q34").Interior.Color = RGB(255, 255, 204)
    ws.Range("Q33:Q34").Font.Bold = True
    With ws.Range("Q34").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="10", Formula2:="3600"
        .IgnoreBlank = False: .ShowError = True: .ShowInput = True
        .InputTitle = "Outlook-Abfrageintervall"
        .InputMessage = "Ganze Sekunden von 10 bis 3600; Standard 60."
        .ErrorTitle = "Ungültiges Abfrageintervall"
        .ErrorMessage = "Bitte eine ganze Zahl von 10 bis 3600 Sekunden eintragen."
    End With
    With ws.Range("Q33").Validation
        .Delete
        .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="255"
        .IgnoreBlank = False: .ShowError = True: .ShowInput = True
        .InputTitle = "Outlook-Postfachname"
        .InputMessage = "Name wie in Outlook angezeigt, z.B. Mensa. Nur dessen Posteingang wird gelesen."
        .ErrorTitle = "Postfachname fehlt"
        .ErrorMessage = "Bitte den Outlook-Postfachnamen eintragen."
    End With
    ThisWorkbook.Names.Add Name:=INTERVAL_NAME, RefersTo:="='Stammdaten'!$Q$34"
    ThisWorkbook.Names.Add Name:="KC_OutlookPostfach", RefersTo:="='Stammdaten'!$Q$33"
    If Not oldCell Is Nothing Then
        If oldCell.Worksheet.Name = ws.Name And oldCell.Address = "$Q$23" And ws.Range("M23").Value = "Outlook-Abfrage alle" Then
            oldCell.ClearContents: oldCell.Validation.Delete
            oldCell.Interior.Pattern = xlNone: oldCell.Font.Bold = False
            oldCell.NumberFormat = "General": oldCell.Locked = True
            ws.Range("M23").ClearContents
            If ws.Range("R23").Value = "Sekunden" Then ws.Range("R23").ClearContents
            If Left(CStr(ws.Range("M24").Value), 21) = "10 bis 3600 Sekunden; " Then ws.Range("M24").ClearContents
        End If
    End If
End Sub

Public Function KC_MailboxName() As String
    Dim v As Variant
    On Error GoTo Invalid
    v = ThisWorkbook.Names("KC_OutlookPostfach").RefersToRange.Value2
    If IsError(v) Or IsEmpty(v) Or VarType(v) = vbBoolean Then GoTo Invalid
    KC_MailboxName = Trim(CStr(v))
    If Len(KC_MailboxName) = 0 Or Len(KC_MailboxName) > 255 Then GoTo Invalid
    Exit Function
Invalid:
    Err.Raise vbObjectError + 683, , "Ungültiger Outlook-Postfachname in Stammdaten Q33; kein Postfach abgefragt"
End Function

Private Function ValidInterval(v As Variant) As Boolean
    Dim n As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    n = CDbl(v)
    ValidInterval = (n >= 10 And n <= 3600 And n = Fix(n))
End Function

Public Function KC_IntervalSeconds() As Long
    Dim intervalCell As Range, c As Worksheet, v As Variant
    KC_IntervalSeconds = 60
    On Error GoTo Invalid
    Set c = KC_Sheet("_KC_Config")
    If ValidInterval(c.Cells(4, 2).Value2) Then KC_IntervalSeconds = CLng(c.Cells(4, 2).Value2)
    Set intervalCell = ThisWorkbook.Names(INTERVAL_NAME).RefersToRange
    v = intervalCell.Value2
    If Not ValidInterval(v) Then GoTo Invalid
    KC_IntervalSeconds = CLng(v)
    c.Cells(4, 2) = KC_IntervalSeconds
    Exit Function
Invalid:
    On Error Resume Next
    c.Cells(8, 2) = "Ungültiges Outlook-Intervall in Stammdaten Q34; bisheriges Intervall " & KC_IntervalSeconds & " Sekunden bleibt aktiv"
    On Error GoTo 0
End Function
