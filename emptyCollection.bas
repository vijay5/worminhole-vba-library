''' очищает коллекцию
Sub emptyCollection(col As Collection)
    Dim i As Long
    If Not col Is Nothing Then
        For i = 1 To col.Count
            col.Remove 1
        Next i
    End If
End Sub