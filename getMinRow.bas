' не проверено
Function getMinRow(Optional sh As Worksheet = Nothing)
    
    If sh Is Nothing Then Set sh = ActiveSheet
    getMinRow = sh.Cells.Find("*", sh.Range("A1"), xlFormulas, , xlByRows, xlNext).Row
End Function