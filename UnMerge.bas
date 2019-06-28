Sub UnMergeFillAnotherOne()

Dim cell As Range

'It works on active worksheet
For Each cell In ThisWorkbook.ActiveSheet.UsedRange
    If cell.MergeCells Then
    'write value to cell in column 4
        Cells(cell.Row, 4) = cell.Value
    'unmerge cells
        cell.MergeCells = False
    'clear merge cells
        cell.Clear
    End If
Next
End Sub
