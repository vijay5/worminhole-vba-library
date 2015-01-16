Function getMaxRow(Optional sh As Worksheet = Nothing)
    If sh Is Nothing Then Set sh = ActiveSheet
    getMaxRow = 0
    On Error GoTo Err
    getMaxRow = sh.Cells.Find("*", sh.Range("A1"), xlFormulas, , xlByRows, xlPrevious).Row
Err:
End Function
