Function getMaxRow(Optional sh As Worksheet = Nothing)
    Dim maxRow As Long
    Dim i As Long
    Dim curColumnMaxRow As Long
    
    If sh Is Nothing Then Set sh = ActiveSheet
    
    maxRow = 0
    For i = 1 To sh.UsedRange.Column + sh.UsedRange.Columns.Count - 1
        curColumnMaxRow = sh.Cells(sh.Cells.Rows.Count, i).End(xlUp).Row
        maxRow = WorksheetFunction.MAX(maxRow, curColumnMaxRow)
    Next i
    
    getMaxRow = maxRow
End Function