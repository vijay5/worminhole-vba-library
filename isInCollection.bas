' эквивалент dictionary.exists(key)
Function isInCollection(key As String, col As Variant)
    Dim tmp As Variant
    
    isInCollection = False
    If Not (TypeName(col) = "Collection" Or _
            TypeName(col) = "Sheets" Or _
            TypeName(col) = "CollectionExtended" Or _
            TypeName(col) = "Workbooks" Or _
            TypeName(col) = "ListObjects" Or _
            TypeName(col) = "PivotTables") Then Exit Function
    
    tmp = "abrakadabra"
    On Error Resume Next
        If TypeName(col) = "CollectionExtended" Then
            tmp = col.GetCollection(key)
            Stop
        Else
            If IsObject(col(key)) Then
                Set tmp = col(key)
            Else
                tmp = col(key)
            End If
        End If
    On Error GoTo 0
    
    If IsArray(tmp) Or IsObject(tmp) Then
        isInCollection = True
    Else
        isInCollection = (tmp <> "abrakadabra")
    End If
End Function