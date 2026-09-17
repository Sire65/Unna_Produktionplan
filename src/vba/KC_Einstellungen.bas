Attribute VB_Name = "KC_Einstellungen"
Option Explicit
Private Const INTERVAL_NAME As String = "KC_OutlookIntervallSekunden"

Public Sub KC_SetupInterval()
    Dim intervalCell As Range, ws As Worksheet, seconds As Long
    On Error Resume Next
    Set intervalCell = ThisWorkbook.Names(INTERVAL_NAME).RefersToRange
    On Error GoTo 0
    If Not intervalCell Is Nothing Then Exit Sub
    Set ws = KC_Sheet("Stammdaten")
    If Application.WorksheetFunction.CountA(ws.Range("M23:R24")) > 0 Then Err.Raise vbObjectError + 681, , "Stammdaten M23:R24 für Outlook-Intervall bereits belegt"
    If ws.ProtectContents Then Err.Raise vbObjectError + 682, , "Stammdaten für erstmalige Intervall-Einrichtung geschützt"
    seconds = 60
    If ValidInterval(KC_Sheet("_KC_Config").Cells(4, 2).Value2) Then seconds = CLng(KC_Sheet("_KC_Config").Cells(4, 2).Value2)
    ws.Range("M23") = "Outlook-Abfrage alle"
    ws.Range("Q23").NumberFormat = "0": ws.Range("Q23") = seconds
    ws.Range("Q23").Locked = False
    ws.Range("Q23").Interior.Color = RGB(255, 255, 204)
    ws.Range("Q23").Font.Bold = True
    ws.Range("R23") = "Sekunden"
    ws.Range("M24") = "10 bis 3600 Sekunden; Standard 60. Gilt ab nächstem Timerlauf."
    With ws.Range("Q23").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="10", Formula2:="3600"
        .IgnoreBlank = False: .ShowError = True: .ShowInput = True
        .InputTitle = "Outlook-Abfrageintervall"
        .InputMessage = "Ganze Sekunden von 10 bis 3600; Standard 60."
        .ErrorTitle = "Ungültiges Abfrageintervall"
        .ErrorMessage = "Bitte eine ganze Zahl von 10 bis 3600 Sekunden eintragen."
    End With
    ThisWorkbook.Names.Add Name:=INTERVAL_NAME, RefersTo:="='Stammdaten'!$Q$23"
End Sub

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
    c.Cells(8, 2) = "Ungültiges Outlook-Intervall in Stammdaten Q23; bisheriges Intervall " & KC_IntervalSeconds & " Sekunden bleibt aktiv"
    On Error GoTo 0
End Function
