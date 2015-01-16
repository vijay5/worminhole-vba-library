Function getMinCol(Optional sh As Worksheet = Nothing)
    If sh Is Nothing Then Set sh = ActiveSheet
    getMinCol = 0
    On Error GoTo Err
    getMinCol = sh.Cells.Find("*", sh.Cells(sh.Cells.Rows.Count, sh.Cells.Columns.Count), xlFormulas, , xlByColumns, xlNext).Column
Err:
End Function
