Function getMaxRow(Optional sh As Worksheet = Nothing)
    Dim sh As Worksheet
    
    If sh Is Nothing Then Set sh = ActiveSheet
    realLastRow = sh.Cells.Find("*", sh.Range("A1"), xlFormulas, , xlByRows, xlPrevious).Row
End Function
