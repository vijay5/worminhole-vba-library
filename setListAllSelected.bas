''' выделяет все элементы ListBox'а / снимает выделение
Sub setListAllSelected(listBox As MSForms.listBox, Optional valueToSet As Boolean = True)
    Dim i As Integer
    
    For i = 0 To listBox.ListCount - 1
        listBox.Selected(i) = valueToSet
    Next i
End Sub
