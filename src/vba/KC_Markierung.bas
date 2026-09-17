Attribute VB_Name = "KC_Markierung"
Option Explicit

Private Function ImportCell(cell As Range) As Boolean
    ImportCell = (cell.Worksheet.Name = "Sonderkostformen") Or (KC_V20.IsRealWeekSheetName(cell.Worksheet.Name) And Not KC_V20.IsWeekCopySheet(cell.Worksheet.Name))
End Function

Private Function BorderTheme(border As Object) As Long
    On Error Resume Next
    BorderTheme = border.ThemeColor
    On Error GoTo 0
End Function

Public Sub KC_SnapshotBorders(txn As Worksheet, ByVal row As Long, cell As Range)
    Dim edges As Variant, k As Long, col As Long, border As Object
    If Not ImportCell(cell) Then Exit Sub
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    txn.Cells(row, 7) = "RAHMEN"
    For k = 0 To 3
        col = 8 + k * 6: Set border = cell.Borders(edges(k))
        txn.Cells(row, col) = border.LineStyle
        txn.Cells(row, col + 1) = border.Weight
        txn.Cells(row, col + 2) = border.Color
        txn.Cells(row, col + 3) = border.ColorIndex
        txn.Cells(row, col + 4) = BorderTheme(border)
        txn.Cells(row, col + 5) = border.TintAndShade
    Next k
End Sub

Public Sub KC_MarkImported(cell As Range)
    Dim edge As Variant
    If Not ImportCell(cell) Then Exit Sub
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With cell.Borders(edge)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(255, 0, 0)
        End With
    Next edge
End Sub

Public Sub KC_RestoreBorders(txn As Worksheet, ByVal row As Long, cell As Range)
    Dim edges As Variant, k As Long, col As Long, border As Object, colorIndex As Long, theme As Long
    If CStr(txn.Cells(row, 7).Value) <> "RAHMEN" Then Exit Sub
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    For k = 0 To 3
        col = 8 + k * 6: Set border = cell.Borders(edges(k))
        colorIndex = CLng(txn.Cells(row, col + 3).Value)
        theme = CLng(txn.Cells(row, col + 4).Value)
        If colorIndex = xlColorIndexAutomatic Or colorIndex = xlColorIndexNone Then
            border.ColorIndex = colorIndex
        ElseIf theme > 0 Then
            border.ThemeColor = theme
            border.TintAndShade = CDbl(txn.Cells(row, col + 5).Value)
        Else
            border.Color = CLng(txn.Cells(row, col + 2).Value)
        End If
        border.Weight = CLng(txn.Cells(row, col + 1).Value)
        ' Zuletzt setzen: Farbe/Gewicht können eine zuvor unsichtbare Kante aktivieren.
        border.LineStyle = CLng(txn.Cells(row, col).Value)
    Next k
End Sub
