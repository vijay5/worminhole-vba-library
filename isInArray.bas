Function isInArray(value As Variant, arr() As Variant) As Boolean
    Dim chk As Boolean
    Dim i As Long
    
    chk = False
    i = LBound(arr)
    Do While i <= UBound(arr) And Not chk
        chk = (value = arr(i))
        i = i + 1
    Loop
    isInArray = chk
End Function