Function getMaxRow(Optional sh As Worksheet = Nothing)
    
    If sh Is Nothing Then Set sh = ActiveSheet
    getMaxRow = sh.Cells.Find("*", sh.Range("A1"), xlFormulas, , xlByRows, xlPrevious).Row
End Function