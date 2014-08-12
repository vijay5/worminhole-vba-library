''' возвращеает длину массива по какому-либо из измерений
Function arrayLength(arr As Variant, Optional degree As Byte = 1) As Long
    Dim tmpDim As Single
    arrayLength = 0 ' по умолчанию
    If InStr(TypeName(arr), "()") > 0 Then
        tmpDim = 0.5
        On Error Resume Next
            tmpDim = UBound(arr, degree)
        On Error GoTo 0
        
        If tmpDim <> 0.5 Then ' есть такая ось в массиве
            arrayLength = UBound(arr, degree) - LBound(arr, degree) + 1
        Else ' такой оси у массива нет
            Exit Function
        End If
    ElseIf TypeName(arr) = "Collection" Then
        arrayLength = arr.Count
    Else
        ' pass
    End If
End Function