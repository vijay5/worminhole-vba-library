Sub makeDropDownList(targetRng As Range, sourceDataRng As Range, Optional ignoreBlank As Boolean = True, Optional showError As Boolean = True)
    ''' Добавляет выпадающий список к указанной ячейке (2003 compatible)
    ''' Список значений задаётся из sourceDataRng (возможно, есть ограничения на кол-во строк/столбцов, т.е. должна быть либо одна строка, либо один столбец)
    Dim shName As String
    Dim firstCellRow As String, firstCellCol As String
    Dim lastCellRow As String, lastCellCol As String
    
    If targetRng Is Nothing Or sourceDataRng Is Nothing Then Exit Sub ' выходим
    
    shName = sourceDataRng.Parent.Name
    firstCellRow = CStr(sourceDataRng.Cells(1, 1).Row)
    firstCellCol = CStr(sourceDataRng.Cells(1, 1).Column)
    lastCellRow = CStr(sourceDataRng.Cells(1, 1).Row + sourceDataRng.Rows.Count - 1)
    lastCellCol = CStr(sourceDataRng.Cells(1, 1).Column + sourceDataRng.Columns.Count - 1)
    
    If CDbl(Application.Version) < 12 Then
        targetRng.Cells(1, 1).Select ' для совместимости с 2003 экселем, необходимо выделять ячейку перед изменением валидации, увы...
    Else
        targetRng.Select
    End If
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
         Formula1:="=INDIRECT(ADDRESS(" + firstCellRow + "," + firstCellCol + ",,,""" + shName + """)&"":""&ADDRESS(" + lastCellRow + "," + lastCellCol + "))"
        .ignoreBlank = ignoreBlank ' игнорировать пропуски
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Ошибка!"
        .InputMessage = ""
        .ErrorMessage = "Введено неверное значение. Выберите значение из выпадающего списка!"
        .ShowInput = True
        .showError = showError
    End With
End Sub