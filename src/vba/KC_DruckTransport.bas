Attribute VB_Name = "KC_DruckTransport"
Option Explicit

Public Sub KC_TransportSetup()
    With ThisWorkbook.Worksheets("Transport").PageSetup
        .PrintArea = "$A$1:$K$36"
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
    End With
End Sub

' Both the individual buttons and the combined print selection use this path.
' A PDF filename exports exactly the same document for verification.
Public Sub KC_TransportPrint(ByVal document As String, Optional ByVal pdfPath As String = "")
    Dim ws As Worksheet, area As String, oldArea As String, oldZoom As Variant
    Dim oldWide As Variant, oldTall As Variant, oldOrientation As Long, oldPaper As Long
    Dim captured As Boolean, errorNumber As Long, errorText As String, restoreError As String
    Select Case document
        Case "TOURENPLAN": area = "$A$1:$K$36"
        Case "MENGENBEWERTUNG": area = "$M$5:$AB$32"
        Case Else: Err.Raise vbObjectError + 740, "KitaFino Druck", "Unbekanntes Transport-Dokument"
    End Select
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("Transport")
    With ws.PageSetup
        oldArea = .PrintArea: oldZoom = .Zoom
        oldWide = .FitToPagesWide: oldTall = .FitToPagesTall
        oldOrientation = .Orientation: oldPaper = .PaperSize
        captured = True
        .PrintArea = area: .Zoom = False
        .FitToPagesWide = 1: .FitToPagesTall = 1
        .Orientation = xlLandscape: .PaperSize = xlPaperA4
    End With
    If Len(pdfPath) > 0 Then
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Else
        ws.PrintOut From:=1, To:=1, Copies:=1, Collate:=True, IgnorePrintAreas:=False
    End If
Restore:
    If captured Then
        ' Attempt every restoration even when a printer-setting operation fails.
        On Error Resume Next
        With ws.PageSetup
            .PrintArea = oldArea
            If Err.Number <> 0 Then restoreError = Err.Description: Err.Clear
            .Orientation = oldOrientation
            If Err.Number <> 0 Then restoreError = Err.Description: Err.Clear
            .PaperSize = oldPaper
            If Err.Number <> 0 Then restoreError = Err.Description: Err.Clear
            .FitToPagesWide = oldWide
            If Err.Number <> 0 Then restoreError = Err.Description: Err.Clear
            .FitToPagesTall = oldTall
            If Err.Number <> 0 Then restoreError = Err.Description: Err.Clear
            .Zoom = oldZoom
            If Err.Number <> 0 Then restoreError = Err.Description: Err.Clear
        End With
        On Error GoTo 0
    End If
    If Len(restoreError) > 0 Then
        errorText = errorText & "; Druckeinstellungen konnten nicht vollständig wiederhergestellt werden: " & restoreError
        If errorNumber = 0 Then errorNumber = vbObjectError + 741
    End If
    If errorNumber <> 0 Then Err.Raise errorNumber, "Transport-Druck", errorText
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume Restore
End Sub

