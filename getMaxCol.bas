Function getMaxCol(Optional sh As Worksheet = Nothing)
    If sh Is Nothing Then Set sh = ActiveSheet
    getMaxCol = 0
    On Error GoTo Err
    getMaxCol = sh.Cells.Find("*", sh.Range("A1"), xlFormulas, , xlByColumns, xlPrevious).Column
Err:
End Function