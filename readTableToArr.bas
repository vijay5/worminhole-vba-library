' читает некоторые столбцы в таблицу, проверяет на уникальность строк, может сделать reBase
' Пример запуска:     readTableToArr Array("WARE", "PO_STATUS", "PO_DATE_DELIVERY", "PLAN_AMOUNT_DELIVERY"), 
'                                    Workbooks("AW15.XLS").Sheets(1).Range("1:1"), outArr, Array("WARE")

'REQUIRES: getMaxRow, arrayLength, getProperValArray, FindCell, addUniqToCol, regroupArray, col2Array, array2col, isInCollection, addToText
Sub readTableToArr(fieldsList As Variant, headerRng As Range, outArr As Variant, Optional keyFieldsList As Variant = False, Optional minRow As Long = 0, Optional minCol As Long = 0)
    
    Dim sh As Worksheet
    Dim maxRow As Long
    Dim colCount As Long
    Dim columnsData() As Variant
    Dim fieldsListRebased As Variant
    Dim hdrColumn As Range
    Dim dtRange As Range
    Dim tmpRng As Range
    Dim rowNum As Long
    Dim colNum As Long
    Dim tmpCol As Collection
    Dim keyFieldsCol As Collection
    Dim item As Variant
    Dim key As String
    Dim curColName As String
    
    
    
    Set sh = headerRng.Parent
    maxRow = getMaxRow(sh)
    colCount = arrayLength(fieldsList)
    
    ReDim columnsData(0 To colCount - 1) ' массив со значениями 1D * 2D
    ' все строки под шапкой таблицы
    Set dtRange = Range(sh.Cells(headerRng.Cells(1, 1).Row + headerRng.Rows.Count, 1), sh.Cells(maxRow, 1)).EntireRow
    
    For i = 0 To colCount - 1
        Set hdrColumn = FindCell(CStr(fieldsList(i)), headerRng, , xlWhole).EntireColumn ' ищем название поля
        If Not hdrColumn Is Nothing Then
            columnsData(i) = getProperValArray(Intersect(hdrColumn, dtRange))
        Else
            columnsData(i) = -1
        End If
    Next i
    
    '
    allUniqueRecords = False
    severalUniqueRecords = False
    If IsArray(keyFieldsList) Then ' задан массив уникальных ключей (они должны входить в список копируемых полей)
        severalUniqueRecords = True
        Set keyFieldsCol = array2col(keyFieldsList, True)
        
    ElseIf TypeName(keyFieldsList) = "Boolean" Then ' задан флаг - True(все поля)/False(ни одно поле)
        allUniqueRecords = keyFieldsList
    Else ' задано что-то ещё = False(ни одно поле)
        ' pass
    End If
    
    If allUniqueRecords Or severalUniqueRecords Then ' проверяем на уникальность по всем или некоторым столбцам
        Set tmpCol = New Collection
        
        For rowNum = 0 To dtRange.Rows.Count - 1
            ReDim item(0 To colCount - 1) ' чистим массив перед записью
            key = ""
            For colNum = 0 To colCount - 1
                item(colNum) = columnsData(colNum)(rowNum + 1, 1) ' заполняем строку
                
                If allUniqueRecords Then ' если берём все столбцы
                    key = addToText(key, CStr(item(colNum)), Chr(1))
                Else ' не все столбцы
                    curColName = CStr(fieldsList(colNum))
                    If isInCollection(curColName, keyFieldsCol) Then ' проверяем вхождение текущего столбца в множество ключевых
                        key = addToText(key, CStr(item(colNum)), Chr(1))
                    Else
                        ' pass
                    End If
                End If
            Next colNum
            addUniqToCol tmpCol, item, key ' добавляем в коллекцию
        Next rowNum
        
        outArr = regroupArray(col2Array(tmpCol), True, minRow, minCol)

    Else ' не проверяем на уникальность
    
        ReDim outArr(0 + minRow To dtRange.Rows.Count - 1 + minRow, 0 + minCol To colCount - 1 + minCol)
        
        For rowNum = 0 To dtRange.Rows.Count - 1
            For colNum = 0 To colCount - 1
                outArr(rowNum + minRow, colNumminCol + minCol) = columnsData(colNum)(rowNum + 1, 1)
            Next colNum
        Next rowNum
    End If
    
End Sub
