' REQUIRES: isInCollection
Sub addUniqToCol(col As Collection, Item As Variant, Optional key As String = "")
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    If key = "" Then
        On Error Resume Next
        key = CStr(Item)
        On Error GoTo 0
    End If
    
    If key = "" Then
        col.Add Item
    Else
        If Not isInCollection(key, col) Then
            col.Add Item, key
        End If
    End If

End Sub
