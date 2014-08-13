''' преобразует коллекцию в 1D/2D-массив
Function col2Array(col As Collection, Optional force2D As Boolean = False) As Variant
    Dim tmpArr As Variant
    Dim el As Variant
    Dim cnt As Long
    
    tmpArr = ""
    If col.Count > 0 Then
        cnt = 0
        If force2D Then ' 2D-массив
            ReDim tmpArr(1 To col.Count, 1 To 1)
            For Each el In col
                 cnt = cnt + 1
                 tmpArr(cnt, 1) = el
            Next el
        Else ' 1D-массив
            ReDim tmpArr(1 To col.Count)
            For Each el In col
                 cnt = cnt + 1
                 tmpArr(cnt) = el
            Next el
        End If
    End If
    
    col2Array = tmpArr
End Function