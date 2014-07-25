' не проверено
Function getMinCol(Optional sh As Worksheet = Nothing)
    
    If sh Is Nothing Then Set sh = ActiveSheet
    getMinCol = sh.Cells.Find("*", sh.Range("A1"), xlFormulas, , xlByColumns, xlNext).Column
End Function
