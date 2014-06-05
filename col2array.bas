''' преобразует коллекцию в 1D-массив
Function col2Array(col As Collection) As Variant
    Dim tmpArr As Variant
    Dim el As Variant
    Dim cnt As Long
    
    tmpArr = ""
    If col.Count > 0 Then
        ReDim tmpArr(0 To col.Count - 1)
        cnt = 0
        For Each el In col
             tmpArr(cnt) = el
             cnt = cnt + 1
        Next el
    End If
    
    col2Array = tmpArr
End Function