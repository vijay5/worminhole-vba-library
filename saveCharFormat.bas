''' сохран€ет цвет букв €чейки в массив
Function saveCharFormat(cl As Range) As Variant
    Dim colorArr As Variant
    Dim k As Long
    
    colorArr = "" ' по умолчанию
    If Len(cl.value) > 0 Then
        ReDim colorArr(1 To Len(cl.value))
        
        For k = 1 To Len(cl.value)
            colorArr(k) = cl.Characters(start:=k, Length:=1).Font.color
        Next k
    Else
        ' pass
    End If
    
    saveCharFormat = colorArr
End Function