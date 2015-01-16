' возвращает список с номерами выделенных элементов. Нумерация с 0.
Function getListOfSelected(listBox As Object) As Collection
    Dim tmp As Variant, i As Integer
    Dim tmpCol As New Collection
    
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) Then
            tmpCol.Add i
        End If
    Next i
    Set getListOfSelected = tmpCol
End Function