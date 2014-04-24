' REQUIRE: isInCollection
Sub buildIndex(col As Collection, itemKey As Variant, itemValue As Variant)
    Dim tmpCol As Collection
    Dim itemKeyStr As String
    Dim itemValueStr As String
    ' создаёт индекс
    
    itemKeyStr = CStr(itemKey)
    itemValueStr = CStr(itemValue)
    If isInCollection(itemKeyStr, col) Then ' есть в коллекции
        Set tmpCol = col.Item(itemKeyStr)(1)
        col.Remove itemKeyStr
    Else ' нет в коллекции - новое значение
        Set tmpCol = New Collection
    End If
    
    If Not isInCollection(itemValueStr, tmpCol) Then
        tmpCol.Add itemValue, itemValueStr
    End If
    
    col.Add Array(itemKeyStr, tmpCol), itemKeyStr ' по ключу можно узнать ключ
    
End Sub