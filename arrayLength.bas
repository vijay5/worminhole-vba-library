' REQUIRES: arrayDepth
''' возвращеает длину массива по какому-либо из измерений
Function arrayLength(arr As Variant, Optional degree As Byte = 1) As Long
    arrayLength = 0 ' по умолчанию
    If InStr(TypeName(arr), "()") > 0 Then
        If arrayDepth(arr) >= 1 Then ' Перед нами массив
            arrayLength = UBound(arr, degree) - LBound(arr, degree) + 1
        End If
    End If
End Function