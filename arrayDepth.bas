''' возвращает число измерений у входящего массива. Если не массив - возвращает 0
Function arrayDepth(arr As Variant) As Byte
    Dim tmp As Variant
    If InStr(TypeName(arr), "()") > 0 Then ' перед нами массив
        On Error Resume Next
            For i = 1 To 200 ' цикл по числу измерений
                tmp = -1.5
                tmp = UBound(arr, i)
                If tmp <> -1.5 And tmp >= 0 Then
                    arrayDepth = i
                Else
                    Exit For
                End If
            Next i
        On Error GoTo 0
    Else ' перед нами не массив
        arrayDepth = 0
    End If
End Function
