' REQUIRES: isInCollection
Function intersectCol(aCol As Collection, bCol As Collection, Optional posIndex As Integer = -1) As Collection
    Dim el As Variant
    Dim destCol As New Collection
    Dim key As String
    
    Set destCol = New Collection
    For Each el In aCol ' перебор элементов вычитаемого множества
        If posIndex <> -1 Then
            key = CStr(el(posIndex)) ' если коллекция содержит 1D-массив, где posIndex - позиция в массиве с кодом ключа
        Else
            key = CStr(el) ' если в коллекции key=item
        End If
            
        
        If isInCollection(key, bCol) Then
            destCol.Add el, key
        Else
            ' pass
        End If
    Next el
    
    Set intersectCol = destCol
End Function