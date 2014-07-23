' REQUIRES: arrayDepth
Function transposeArr(inArr As Variant, Optional forObjects As Boolean = False) As Variant
    Dim outArr As Variant
    Dim rowNum As Long
    Dim colNum As Long
    
    outArr = ""
    
    If arrayDepth(inArr) = 2 Then
        ReDim outArr(LBound(inArr, 2) To UBound(inArr, 2), LBound(inArr, 1) To UBound(inArr, 1))
        If forObjects Then
            For rowNum = LBound(inArr, 2) To UBound(inArr, 2)
                For colNum = LBound(inArr, 1) To UBound(inArr, 1)
                    Set outArr(rowNum, colNum) = inArr(colNum, rowNum) ' переносим объект
                Next colNum
            Next rowNum
        
        Else
            For rowNum = LBound(inArr, 2) To UBound(inArr, 2)
                For colNum = LBound(inArr, 1) To UBound(inArr, 1)
                    outArr(rowNum, colNum) = inArr(colNum, rowNum) ' переносим значение
                Next colNum
            Next rowNum
        End If
    Else
        ' pass
    End If
    
    transposeArr = outArr
End Function