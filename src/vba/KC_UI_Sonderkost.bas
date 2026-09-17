Attribute VB_Name = "KC_UI_Sonderkost"

Option Explicit

Public Sub KC_UI_ImportKitafino()
  KC_Initialize
  KC_Sheet("KitaFino_Warteschlange").Activate
End Sub

Public Sub KC_UI_ImportProtokoll()

  ThisWorkbook.Worksheets("Import_Protokoll").Visible = xlSheetVisible

  ThisWorkbook.Worksheets("Import_Protokoll").Activate

End Sub

Public Sub KC_UI_Druck_SonderkostMatrix()

  Dim w As Worksheet, lastRow As Long, oldArea As String

  Set w = ThisWorkbook.Worksheets("Sonderkostformen")

  lastRow = w.Cells(w.Rows.count, "AK").End(xlUp).row

  If lastRow < 5 Then MsgBox "Keine Sonderkost-Matrix gefunden.", vbExclamation: Exit Sub

  oldArea = w.PageSetup.PrintArea

  On Error GoTo EH

  With w.PageSetup

    .PrintArea = w.Range("AK1:AX" & lastRow).Address

    .Orientation = xlLandscape

    .zoom = False

    .FitToPagesWide = 1

    .FitToPagesTall = False

  End With

  w.PrintPreview

DONE:

  On Error Resume Next

  w.PageSetup.PrintArea = oldArea

  Exit Sub

EH:

  MsgBox "Druckvorschau konnte nicht geoeffnet werden: " & Err.Description, vbExclamation

  Resume DONE

End Sub

