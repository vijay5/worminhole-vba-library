''' ƒобавл€ет выпадающий список к указанной €чейке (2003 compatible)
''' —писок значений задаЄтс€ из sourceDataRng (возможно, есть ограничени€ на кол-во строк/столбцов, т.е. должна быть либо одна строка, либо один столбец)
' REQUIRES: col2Array, array2col
Sub makeDropDownList(targetRng As Range, sourceData As Variant, Optional ignoreBlank As Boolean = True, Optional showError As Boolean = True)
    Dim shName As String
    Dim sourceDataRng As Range
    Dim firstCellRow As String, firstCellCol As String
    Dim lastCellRow As String, lastCellCol As String
    Dim isRange As Boolean
    Dim validObj As Validation
    Dim oldSelectionAddr As String
    Dim sourceDataCol As Collection
    Dim sourceDataArr As Variant
    Dim formulaStr As String
    Dim i As Long
    
    If targetRng Is Nothing Then Exit Sub ' выходим
    
    isRange = False
    Select Case TypeName(sourceData)
    Case "Range"
        Set sourceDataRng = sourceData
        isRange = True
    Case "Collection"
        Set sourceDataCol = sourceData
        sourceDataArr = col2Array(sourceDataCol)
    Case "Variant()", "String()", "Integer()", "Single()", "Long()", "Double()"
        sourceDataArr = col2Array(array2col(sourceData))
    Case Else
        Exit Sub
    End Select
    
    
    If isRange Then ' задан диапазон
        shName = sourceDataRng.Parent.Name
        firstCellRow = CStr(sourceDataRng.Cells(1, 1).row)
        firstCellCol = CStr(sourceDataRng.Cells(1, 1).Column)
        lastCellRow = CStr(sourceDataRng.Cells(1, 1).row + sourceDataRng.Rows.Count - 1)
        lastCellCol = CStr(sourceDataRng.Cells(1, 1).Column + sourceDataRng.Columns.Count - 1)
        
        formulaStr = "=INDIRECT(ADDRESS(" + firstCellRow + "," + firstCellCol + ",,,""" + shName + """)&"":""&ADDRESS(" + lastCellRow + "," + lastCellCol + "))"
    Else ' задан массив
        
        ReDim sourceDataArrStr(LBound(sourceDataArr) To UBound(sourceDataArr))
        For i = LBound(sourceDataArr) To UBound(sourceDataArr)
            sourceDataArrStr(i) = CStr(sourceDataArr(i))
        Next i
        
        formulaStr = Join(sourceDataArrStr, ",")
        
    End If
    
    ' совместимость со старой версией
    If CDbl(Application.Version) < 12 Then
        oldSelectionAddr = Selection.Address
        targetRng.Cells(1, 1).Select ' дл€ совместимости с 2003 экселем, необходимо выдел€ть €чейку перед изменением валидации, увы...
        Set validObj = Selection.Validation
    Else
        oldSelectionAddr = ""
        Set validObj = targetRng.Cells(1, 1).Validation
    End If
    
    
    With validObj
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
         Formula1:=formulaStr
        .ignoreBlank = ignoreBlank ' игнорировать пропуски
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "ќшибка!"
        .InputMessage = ""
        .ErrorMessage = "¬ведено неверное значение. ¬ыберите значение из выпадающего списка!"
        .ShowInput = True
        .showError = showError
    End With
    
    
    
    If oldSelectionAddr <> "" Then ' восстанавливаем выделеннные €чейки
        Range(oldSelectionAddr).Select
    End If
End Sub
