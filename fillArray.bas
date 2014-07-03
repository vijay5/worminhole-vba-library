''' создаёт массив заданной размерности, заполненный значениями или объектами
Function fillArray(value As Variant, ParamArray dimensions_in())

    Dim tmp As Variant
    Dim dimensions As Variant
    Dim j As Long
    Dim i1 As Long, i2 As Long, i3 As Long, i4 As Long, i5 As Long, i6 As Long, i7 As Long, i8 As Long, i9 As Long, i10 As Long
    Dim numOfDims As Long
    Dim funcName As String
    
    dimensions = dimensions_in
    
    funcName = "fillArray"
    
    numOfDims = arrayLength(dimensions) ' узнаём фактический размер
    ReDim Preserve dimensions(0 To 9) ' расширяем до 9 и заполнеяем нулями
    For j = numOfDims To 9
        dimensions(j) = 0
    Next j
    
    fillArray = "" ' значение по умолчанию
    
    Select Case numOfDims
    Case 1:  ReDim tmp(0 To dimensions(0))
    Case 2:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1))
    Case 3:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2))
    Case 4:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3))
    Case 5:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4))
    Case 6:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5))
    Case 7:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5), 0 To dimensions(6))
    Case 8:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5), 0 To dimensions(6), 0 To dimensions(7))
    Case 9:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5), 0 To dimensions(6), 0 To dimensions(7), 0 To dimensions(8))
    Case 10: ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5), 0 To dimensions(6), 0 To dimensions(7), 0 To dimensions(8), 0 To dimensions(9))
    Case Else
        addJournal "funcName", "[Warning]", "Число измерений должно быть целым числом в интервале от 1 до 10"
        Exit Function
    End Select
        
    
    For i10 = 0 To dimensions(9)
    For i9 = 0 To dimensions(8)
    For i8 = 0 To dimensions(7)
    For i7 = 0 To dimensions(6)
    For i6 = 0 To dimensions(5)
    For i5 = 0 To dimensions(4)
    For i4 = 0 To dimensions(3)
    For i3 = 0 To dimensions(2)
    For i2 = 0 To dimensions(1)
    For i1 = 0 To dimensions(0)
        If IsObject(value) Then ' присваиваем объект
            Select Case numOfDims
            Case 1:  Set tmp(i1) = value
            Case 2:  Set tmp(i1, i2) = value
            Case 3:  Set tmp(i1, i2, i3) = value
            Case 4:  Set tmp(i1, i2, i3, i4) = value
            Case 5:  Set tmp(i1, i2, i3, i4, i5) = value
            Case 6:  Set tmp(i1, i2, i3, i4, i5, i6) = value
            Case 7:  Set tmp(i1, i2, i3, i4, i5, i6, i7) = value
            Case 8:  Set tmp(i1, i2, i3, i4, i5, i6, i7, i8) = value
            Case 9:  Set tmp(i1, i2, i3, i4, i5, i6, i7, i8, i9) = value
            Case 10: Set tmp(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10) = value
            End Select
        Else
            Select Case numOfDims ' присваиваем значение
            Case 1:  tmp(i1) = value
            Case 2:  tmp(i1, i2) = value
            Case 3:  tmp(i1, i2, i3) = value
            Case 4:  tmp(i1, i2, i3, i4) = value
            Case 5:  tmp(i1, i2, i3, i4, i5) = value
            Case 6:  tmp(i1, i2, i3, i4, i5, i6) = value
            Case 7:  tmp(i1, i2, i3, i4, i5, i6, i7) = value
            Case 8:  tmp(i1, i2, i3, i4, i5, i6, i7, i8) = value
            Case 9:  tmp(i1, i2, i3, i4, i5, i6, i7, i8, i9) = value
            Case 10: tmp(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10) = value
            End Select
        End If
    Next i1
    Next i2
    Next i3
    Next i4
    Next i5
    Next i6
    Next i7
    Next i8
    Next i9
    Next i10
    
    fillArray = tmp
    
End Function
