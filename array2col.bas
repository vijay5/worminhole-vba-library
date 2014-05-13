' REQUIRES: isInCollection
Function array2col(arr As Variant) As Collection
    Dim tmpCol As Collection
    Dim Item As Variant
    Dim key As String
    Dim el As Variant
    
    
    Set tmpCol = New Collection
    
    For Each el In arr
        key = CStr(el)
        Item = el
        If Not isInCollection(key, tmpCol) Then
            tmpCol.Add Item, key
        Else
            ' pass
        End If
    Next el

    Set array2col = tmpCol ' возвращаем коллекцию уникальных элементов
    
End Function