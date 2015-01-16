' склеивает в один вектор несколько 1-D массивов
' REQUIRES: arrayLength
Function mergeVectors(ParamArray arr() As Variant) As Variant
    Dim totalLength As Long
    Dim curLength As Long
    Dim i As Long
    Dim j As Long
    
    totalLength = 0
    
    For i = LBound(arr) To UBound(arr)
        totalLength = totalLength + UBound(arr(i)) - LBound(arr(i)) + 1
    Next i
    
    ReDim outArr(0 To totalLength - 1)
    
    curLength = -1
    For i = LBound(arr) To UBound(arr)
        For j = LBound(arr(i)) To UBound(arr(i))
            curLength = curLength + 1
            outArr(curLength) = arr(i)(j)
        Next j
    Next i
    
    mergeVectors = outArr
End Function
