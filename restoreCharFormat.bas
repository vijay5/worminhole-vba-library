''' восстанавливает цвет букв из массива
Sub restoreCharFormat(cl As Range, colorArr As Variant)
    Dim k As Long
    Dim maxLen As Long
    
    If IsArray(colorArr) Then ' если есть цвета для букв
        maxLen = WorksheetFunction.Min(arrayLength(colorArr), Len(cl.value)) ' символ, до которого можем прописать формат
        
        For k = 1 To maxLen
            cl.Characters(start:=k, Length:=1).Font.color = colorArr(k)
        Next k
    Else
        ' pass
    End If
    
End Sub