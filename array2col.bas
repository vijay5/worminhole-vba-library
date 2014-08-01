' REQUIRES: isInCollection
''' преобразует xD-массив в коллекцию
Function array2col(arr As Variant, Optional uniqnessCheck As Boolean = True) As Collection
    Dim tmpCol As Collection
    Dim Item As Variant
    Dim key As String
    Dim el As Variant
    
    Set tmpCol = New Collection
    
    If IsArray(arr) Then
        For Each el In arr
            If uniqnessCheck Then
                key = CStr(el)
                Item = el
                If Not isInCollection(key, tmpCol) Then
                    tmpCol.Add Item, key
                Else
                    ' pass
                End If
            Else ' без проверки повтора ключей - тупо перенос в колелкцию
                tmpCol.Add el
            End If
        Next el
    Else
        ' pass
    End If

    Set array2col = tmpCol ' возвращаем коллекцию уникальных элементов
    
End Function