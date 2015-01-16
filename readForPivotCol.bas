''' TODO: нужно вывод сделать не поячеечно, а писать в массив, а затем выводить весь массив целиком...
' REQUIRES: MatrixPart, getProperValArray, getFlatArray, mergeVectors, arrayLength
Function ReadForPivotCol(sourceRange As Range, numOfRowProperties As Integer, numOfColProperties As Integer, Optional numOfRowParams As Integer = 1, Optional numOfColParams As Integer = 1, Optional getNonEmptyOnly As Boolean = False) As Collection
    ''' Преобразует текущий лист в плоскую таблицу
    ''' Внешний вид входящей таблицы:
    '''
    ''' [*]                +-----------------------+-----------------------+
    '''                    |Var1-1                 |Var1-2                 |
    '''                    +-------+-------+-------+-------+-------+-------+
    '''   Col1  Col2  ColN |Var2-1 |Var2-2 |Var2-3 |Var2-1 |Var2-2 |Var2-3 |
    '''  +-----+-----+-----+-------+-------+-------+-------+-------+-------+
    '''  |     |     |     |TLCell | ...   | ...   | ...   | ...   | ...   |
    '''
    ''' где
    ''' [*] - левый верхний угол листа с данными
    ''' Col1-ColN - число фиксированных переменных (numOfHeaderCols)
    ''' Var1 - "непостоянная" переменная (первая из двух, всего переменных неограничено)
    ''' Var2 - "непостоянная" переменная (вторая из двух, подчинена первой)
    ''' TLCell - первая ячейка с данными
    ''' здесь видно, что кол-во строк над TLCell равно кол-ву "непостоянных переменных" (если таблица составлена верно)
    '''
    ''' на выходе будет следующая "простая" таблица
    '''  Col1  Col2  ColN ||Var1  Var2 || Value
    ''' +-----+-----+-----++-----+-----++-------+
    ''' |     |     |     ||     |     ||       |

    Dim height0 As Long, height As Long, width0 As Long, width As Long
    Dim rowLabels As Variant, colLabels As Variant, sourceData As Variant
    Dim minRow As Long, maxRow As Long
    Dim sourceSheet As Variant
    Dim resultSheet As Worksheet
                                  
    Dim rowHdrCol As Collection
    Dim colHdrCol As Collection
    Dim dataCol As Collection
    Dim rowNum As Long
    Dim colNum As Long
    Dim tmpArr As Variant
    Dim cellArr As Variant
    Dim cellFlatArr As Variant
    Dim rowFlatArr As Variant
    Dim colFlatArr As Variant
    Dim numOfRecords As Long
    Dim i As Long
    Dim arrIsEmpty As Variant
    
    ' работаем по исходному выделению
    Set sourceSheet = sourceRange.Parent
    height0 = sourceRange.Rows.Count
    width = sourceRange.Columns.Count
    
    If height0 <= numOfColProperties Or width <= numOfRowProperties Then
       Call MsgBox("Выделена слишком маленькая область")
       Exit Function
    End If
    
    ' TODO: нужно сделать механизм автоматического Unmerge'а ячеек и заполнения их значениями
    
    ' копируем значения в 2D-массив
    rowLabels = getProperValArray(sourceRange.Cells(1 + numOfColProperties, 1).Resize(height0 - numOfColProperties, numOfRowProperties))
    colLabels = getProperValArray(sourceRange.Cells(1, 1 + numOfRowProperties).Resize(numOfColProperties, width - numOfRowProperties))
    
    ' диапазон с цифровыми данными
    sourceData = getProperValArray(sourceRange.Resize(height0 - numOfColProperties, width - numOfRowProperties). _
                                               Offset(numOfColProperties, numOfRowProperties))
                                               
    ' габариты массивов
    height0 = arrayLength(sourceData, 1) ' число строк в исходном массиве
    height = height0 \ numOfRowParams    ' число строк на одну секцию финального массива
    width0 = arrayLength(sourceData, 2)  ' число столбцов в исходном массиве
    width = width0 \ numOfColParams      ' число столбцов на одну секцию финального массива
    
    If width * numOfColParams <> width0 Then
        MsgBox "Массив с данными не кратен numOfColParams"
        Exit Function
    End If
                                  
    If height * numOfRowParams <> height0 Then
        MsgBox "Массив с данными не кратен numOfColParams"
        Exit Function
    End If
                                  


    ' составляем массивы с метками строк и столбцов
    Set rowHdrCol = New Collection
    Set colHdrCol = New Collection
    For rowNum = 1 To height Step numOfRowParams ' номер строки в исходном диапазоне
        ReDim tmpArr(0 To numOfRowProperties - 1)
        tmpArr = MatrixPart(rowLabels, rowNum, rowNum, 1, numOfRowProperties, True, False)
        rowHdrCol.Add tmpArr, CStr(rowNum) ' пишем в коллекцию
    Next rowNum
        
    For colNum = 1 To width Step numOfColParams ' номер столбца в исходном диапазоне
        ReDim tmpArr(0 To numOfRowProperties - 1)
        tmpArr = MatrixPart(colLabels, 1, numOfColProperties, colNum, colNum, True, False)
        colHdrCol.Add tmpArr, CStr(colNum) ' пишем в коллекцию
    Next colNum
        
    ' переписываем массив данных в коллекции
    numOfRecords = 0
    Set dataCol = New Collection
    For rowNum = 1 To height Step numOfRowParams ' номер строки в исходном диапазоне
        
        For colNum = 1 To width Step numOfColParams ' номер столбца в исходном диапазоне
            ' курсор находится в левой верхней ячейке
            
            ' 2-D массив со значением ячейки
            cellArr = MatrixPart(sourceData, rowNum, rowNum + numOfRowParams - 1, colNum, colNum + numOfColParams - 1, , False)
            cellFlatArr = getFlatArray(cellArr, 0) ' 1-d массив значений
            arrIsEmpty = True
            For i = LBound(cellFlatArr) To UBound(cellFlatArr)
                arrIsEmpty = arrIsEmpty And IsEmpty(cellFlatArr(i))
            Next i
            
            If (getNonEmptyOnly And Not arrIsEmpty) Or Not getNonEmptyOnly Then ' если в массиве что-то есть
                rowFlatArr = rowHdrCol.Item(CStr(rowNum))
                colFlatArr = colHdrCol.Item(CStr(colNum))
                
                tmpArr = mergeVectors(rowFlatArr, colFlatArr, cellFlatArr)
                
                numOfRecords = numOfRecords + 1
                'dataCol.Add Array(CStr(numOfRecords), tmpArr), CStr(numOfRecords)
                dataCol.Add tmpArr, CStr(numOfRecords)
            End If
            
        Next colNum
    Next rowNum
    
    Set ReadForPivotCol = dataCol
    
End Function