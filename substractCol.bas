' REQUIRES: isInCollection
Function substractCol(sourceCol As Collection, substrCol As Collection, Optional posIndex As Integer = -1) As Collection
    Dim el As Variant
    Dim destCol As New Collection
    Dim key As String
    
    Set destCol = sourceCol
    For Each el In substrCol ' перебор элементов вычитаемого множества
        If posIndex <> -1 Then
            key = CStr(el(posIndex)) ' если коллекция содержит 1D-массив, где posIndex - позиция в массиве с кодом ключа
        Else
            key = CStr(el) ' если в коллекции key=item
        End If
        
        If isInCollection(key, destCol) Then
            destCol.Remove key
        Else
            ' pass
        End If
    Next el
    
    Set substractCol = destCol
End Function
