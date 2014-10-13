''' эквивалент dictionary.exists(key)
Function isInCollection(key As String, col As Variant)
    Dim tmp As Variant
    
    isInCollection = False
    If Not (TypeName(col) = "Collection" Or _
            TypeName(col) = "Sheets" Or _
            TypeName(col) = "CollectionExtended" Or _
            TypeName(col) = "Workbooks" Or _
            TypeName(col) = "ListObjects" Or _
            TypeName(col) = "PivotTables" Or _
            TypeName(col) = "Names" Or _
            TypeName(col) = "Shapes") Then Exit Function
    
    tmp = "abrakadabra"
    On Error Resume Next
        If TypeName(col) = "CollectionExtended" Then
            tmp = col.GetCollection(key)
            Stop
        Else
            If IsObject(col(key)) Then
                Set tmp = col.Item(key)
            Else
                tmp = col.Item(key)
            End If
        End If
    On Error GoTo 0
    
    If IsArray(tmp) Or IsObject(tmp) Or IsError(tmp) Then
        isInCollection = True
    Else
        isInCollection = (tmp <> "abrakadabra")
    End If
End Function