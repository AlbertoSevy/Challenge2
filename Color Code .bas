Attribute VB_Name = "Module2"
Sub ColorPositiveNegative()

    Dim ws As Worksheet
    Dim cell As Range
    
    For Each ws In ThisWorkbook.Worksheets
        For Each cell In ws.Range("J1", ws.Cells(ws.Rows.Count, "J").End(xlUp))
            If IsNumeric(cell.Value) Then
                If cell.Value < 0 Then
                    cell.Interior.Color = RGB(255, 0, 0)  ' Red
                ElseIf cell.Value > 0 Then
                    cell.Interior.Color = RGB(0, 255, 0)  ' Green
                End If
            End If
        Next cell
    Next ws

End Sub
