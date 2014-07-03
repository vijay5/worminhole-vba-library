''' простенькая функция для добавления одной или нескольких записей в конец массива
Sub appendTo(arr As Variant, ParamArray itemsToAppend() As Variant)
    Dim numOfEls As Long, i As Long
    Dim initSize As Long
    numOfEls = UBound(itemsToAppend) - LBound(itemsToAppend) + 1
    
    If Not IsArray(arr) Then
        ReDim arr(0 To numOfEls - 1)
        For i = 0 To numOfEls - 1
            arr(i) = itemsToAppend(i)
        Next i
    Else
        initSize = UBound(arr) ' первоначальный размер
        ReDim Preserve arr(LBound(arr) To UBound(arr) + numOfEls)
        For i = 0 To numOfEls - 1
            arr(initSize + i + 1) = itemsToAppend(i)
        Next i
    End If
End Sub
