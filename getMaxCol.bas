Function getMaxCol(Optional sh As Worksheet = Nothing)
    
    If sh Is Nothing Then Set sh = ActiveSheet
    getMaxCol = sh.Cells.Find("*", sh.Range("A1"), xlFormulas, , xlByColumns, xlPrevious).Column
End Function