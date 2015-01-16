Function getMinRow(Optional sh As Worksheet = Nothing)
    If sh Is Nothing Then Set sh = ActiveSheet
    getMinRow = 0
    On Error GoTo Err
    getMinRow = sh.Cells.Find("*", sh.Cells(sh.Cells.Rows.Count, sh.Cells.Columns.Count), xlFormulas, , xlByRows, xlNext).Row
Err:
End Function