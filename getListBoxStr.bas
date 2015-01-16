''' возвращает текстовое значение элементов из ListBox'а
''' REQUIRES: addToText
Function getListBoxStr(lbox As MSForms.listBox) As String
    Dim listBoxStr As String
    Dim rowStr As String
    Dim i As Long
    Dim j As Long
    
    listBoxStr = ""
    For i = 0 To lbox.ListCount - 1
        rowStr = ""
        For j = 0 To 9
            rowStr = addToText(rowStr, IIf(IsNull(lbox.List(i, j)), "", lbox.List(i, j)), Chr(9))
        Next j
        listBoxStr = addToText(listBoxStr, rowStr, Chr(10))
    Next i
    getListBoxStr = listBoxStr
End Function