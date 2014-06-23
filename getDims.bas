' REQUIRES: arrayDepth
Function getDims(arr As Variant) As Variant
    Dim axis As Byte
    Dim numOfAxis As Byte
    Dim dimsArr() As Long
    
    numOfAxis = arrayDepth(arr)
    
    ReDim dimsArr(1 To 2, 1 To numOfAxis)
    
    For axis = 1 To numOfAxis
        dimsArr(1, axis) = LBound(arr, axis)
        dimsArr(1, axis) = UBound(arr, axis)
    Next i
    
    getDims = dimsArr
End Function
