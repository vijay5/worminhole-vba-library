''' склеивает диапазоны/2D массивы и возвращает одну большую 2D-таблицу
' REQUIRES: arrayLength
Function mergeMatrixes(direction As XlDirection, ParamArray rangeArray() As Variant) As Variant
    Dim rng As Variant
    Dim maxRow As Long
    Dim maxCol As Long
    Dim colNum As Long
    Dim rowNum As Long
    Dim maxRowPrev As Long
    Dim maxColPrev As Long
    Dim outArr As Variant
    Dim tmpArr As Variant
    Dim rngArr As Variant
    Dim rangeValueArray As Variant ' список таблиц (rng -> rng.Value)
    
    Dim startCnt As Long
    Dim endCnt As Long
    Dim stepCnt As Long
    Dim i As Long
    
    Dim tmpDim2 As Single
    Dim tmpDim3 As Single
    Dim chk3 As Boolean
    Dim outRowNum As Long
    Dim outColNum As Long
    
    mergeMatrixes = "" ' по-умолчанию - возвращаем ошибку
    ' в зависимости от направления задаём параметры перебора списка таблиц
    Select Case direction
    Case XlDirection.xlDown, XlDirection.xlToRight ' сверху вниз
        startCnt = LBound(rangeArray)
        endCnt = UBound(rangeArray)
        stepCnt = 1
        
    Case XlDirection.xlUp, XlDirection.xlToLeft   ' снизу вверх
        startCnt = UBound(rangeArray)
        endCnt = LBound(rangeArray)
        stepCnt = -1
        
    Case Else
        Exit Function
    End Select
    
    
    ' преобразуем диапазоны в массив таблиц
    ' одновременно проверяем, чтобы у массивов было ровно 2 оси
    ReDim rangeValueArray(LBound(rangeArray) To UBound(rangeArray))
    
    For i = startCnt To endCnt Step stepCnt
        If TypeName(rangeArray(i)) = "Range" Then ' на входе - диапазон
            ' корректно преобразуем диапазон в 2D-массив
            If rangeArray(i).Cells.Count = 1 Then
                ReDim tmpArray(1 To 1, 1 To 1)
                tmpArray(1, 1) = rangeArray(i).value
                rangeValueArray(i) = tmpArray
            Else
                rangeValueArray(i) = rangeArray(i).value
            End If
            
            
        ElseIf InStr(TypeName(rangeArray(i)), "()") > 0 Then ' на входе - массив
            
            ' / уточняем размерность массива
            tmpDim2 = 0.5
            tmpDim3 = 0.5
            On Error Resume Next
                tmpDim2 = UBound(rangeArray(i), 2)
                tmpDim3 = UBound(rangeArray(i), 3)
            On Error GoTo 0
            chk3 = (tmpDim2 <> 0.5) And (tmpDim3 = 0.5)
            
            If chk3 Then
                rangeValueArray(i) = rangeArray(i)
                
            Else ' размерность массива 1 или >= 3
                MsgBox "На входе должен быть строго 2D-массив"
                Exit Function
            End If
        Else
            ' pass
        End If
    Next i
    
    ' в этой точке все диапазоны преобразованы в массивы
    
    
    ' сначала считаем, сколько строк/столбцов будет
    maxRow = 0
    maxCol = 0
    For i = startCnt To endCnt Step stepCnt
        Select Case direction
        Case XlDirection.xlDown, XlDirection.xlUp ' вверх-вниз
            maxRow = maxRow + arrayLength(rangeValueArray(i), 1)
            maxCol = WorksheetFunction.MAX(maxCol, arrayLength(rangeValueArray(i), 2))
        
        Case XlDirection.xlToRight, XlDirection.xlToLeft   ' влево-вправо
            maxRow = WorksheetFunction.MAX(maxRow, arrayLength(rangeValueArray(i), 1))
            maxCol = maxCol + arrayLength(rangeValueArray(i), 2)
        
        Case Else
            Exit Function
            
        End Select

    Next i
    
    ' выделяем место, сделали принудительный rebase к (1, 1)
    ReDim outArr(1 To maxRow, 1 To maxCol)
    maxRow = LBound(outArr, 1) - 1
    maxCol = LBound(outArr, 2) - 1
    
    ' перенос данных
    For i = startCnt To endCnt Step stepCnt ' цикл по диапазонам
        rngArr = rangeValueArray(i)
        Select Case direction
        Case XlDirection.xlDown, XlDirection.xlUp ' вверх-вниз
            maxRowPrev = maxRow + 1
            maxColPrev = LBound(outArr, 2)
            
            maxRow = maxRow + arrayLength(rngArr, 1)
            maxCol = WorksheetFunction.MAX(maxCol, arrayLength(rngArr, 2))
        
        Case XlDirection.xlToRight, XlDirection.xlToLeft   ' влево-вправо
            maxRowPrev = LBound(outArr, 1)
            maxColPrev = maxCol + 1
            
            maxRow = WorksheetFunction.MAX(maxRow, arrayLength(rngArr, 1))
            maxCol = maxCol + arrayLength(rngArr, 2)
        
        Case Else
            Exit Function
            
        End Select
        
        
        ' физический перенос
        For rowNum = LBound(rngArr, 1) To UBound(rngArr, 1)
            For colNum = LBound(rngArr, 2) To UBound(rngArr, 2)
                outRowNum = maxRowPrev + rowNum - LBound(rngArr, 1)
                outColNum = maxColPrev + colNum - LBound(rngArr, 2)
                outArr(outRowNum, outColNum) = rngArr(rowNum, colNum)
            Next colNum
        Next rowNum
        
    Next i
    
    mergeMatrixes = outArr ' возвращаем массив
    
End Function
