''' чтение данных из плосткой таблицы (на одном листе) по колонкам
''' предполагается, что поля в таблице идут в том же порядке, что и в destFieldNames
''' названия колонок в первой строке не нужны
' REQUIRES: col2Array, isInCollection
Sub db_excelToDB(excelRng As Variant, conn As ADODB.Connection, destTblName As String, destFieldNames As Variant)
    Dim dataArr As Variant
    Dim chk0 As Boolean
    Dim chk1 As Boolean
    Dim chk2 As Boolean
    Dim chk3 As Boolean
    Dim rowNum As Long
    Dim tableRs As New ADODB.Recordset
    Dim queryStr As String
    Dim fld As ADODB.Field
    Dim destFieldNamesArr As Variant
    Dim tmpDim2 As Single
    Dim tmpDim3 As Single
    Dim fldNum As Long
    Dim fldName As String
    Dim fldNameAlpha As String
    Dim noFldsList As Collection
    Dim destFieldNamesTmp As Collection
    Dim existingFldsList As Collection
    Dim tmp As Variant
         
    ' ///// проверки
    ' / входящий массив значений - массив или Range
    chk0 = (TypeName(excelRng) = "Range" Or InStr(TypeName(excelRng), "()") > 0)
    If Not chk0 Then
        MsgBox "[db_readExcelToDB] На вход можно подать только 2D-массив или Range"
        Exit Sub
    End If
    
    
    ' / названия полей - список или массив
    chk1 = TypeName(destFieldNames) = "Collection" Or InStr(TypeName(destFieldNames), "()") > 0
    If TypeName(excelRng) = "Range" Then
        chk2 = excelRng.Rows.Count >= 1
    ElseIf InStr(TypeName(excelRng), "()") > 0 Then
        chk2 = (UBound(excelRng) - LBound(excelRng) + 1) > 0
    Else
        chk2 = False
    End If
    
    ' / уточняем размерность массива на входе
    If InStr(TypeName(excelRng), "()") > 0 Then
        tmpDim2 = 0.5
        tmpDim3 = 0.5
        On Error Resume Next
            tmpDim2 = UBound(excelRng, 2)
            tmpDim3 = UBound(excelRng, 3)
        On Error GoTo 0
        chk3 = (tmpDim2 <> 0.5) And (tmpDim3 = 0.5)
    Else
        ' pass
        chk3 = True
    End If
    
    
    
    If Not (chk1 And chk2 And chk3) Then Exit Sub
    ' проверки закончены
    
    If TypeName(destFieldNames) = "Collection" Then
        Set destFieldNamesTmp = destFieldNames
        destFieldNamesArr = col2Array(destFieldNamesTmp) ' 0-based array
    Else
        destFieldNamesArr = destFieldNames
    End If
    
    
    If TypeName(excelRng) = "Range" Then
        ' сохраняем значения в массив
        If excelRng.Cells.Count = 1 Then
            ReDim dataArr(1 To 1, 1 To 1)
            dataArr(1, 1) = excelRng.Value
        Else
            dataArr = excelRng.Value
        End If
    Else ' 2D массив
        dataArr = excelRng
    End If
    
    ' устанавливаем соединение с БД
    queryStr = "SELECT * FROM " + destTblName + ";"
    tableRs.ActiveConnection = conn.ConnectionString
    tableRs.LockType = adLockOptimistic
    tableRs.source = queryStr
    tableRs.Open ' открываем табличку
    
    ' проверка наличия всех полей в таблице
    Set noFldsList = New Collection
    Set existingFldsList = New Collection
    
    For fldNum = LBound(destFieldNamesArr, 1) To UBound(destFieldNamesArr, 1) ' перебор полей
        tmp = "abrakadabra"
        On Error Resume Next
            fldName = destFieldNamesArr(fldNum)
            ' для проверки вхождения, характерно только для Access...
            fldNameAlpha = Replace(Replace(fldName, "[", ""), "]", "")
            tmp = tableRs.Fields.item(fldNameAlpha).Name
        On Error GoTo 0
        If tmp = "abrakadabra" Then ' нет такого поля
            If Not isInCollection(fldNameAlpha, noFldsList) Then
                noFldsList.Add fldNameAlpha
            Else
                ' pass
            End If
        Else ' есть такое поле
            If Not isInCollection(fldName, existingFldsList) Then
                existingFldsList.Add fldName
            Else
                ' pass
            End If
        End If
    Next fldNum
    
    If noFldsList.Count > 0 Then
        tableRs.Close
        MsgBox "[db_readExcelToDB] В таблице отсутствуют поля: " + Join(col2Array(noFldsList), ", ")
        Exit Sub
    Else
        ' pass
    End If
    
    ' переливаем данные
    ' одна строка в массиве - одна запись в БД
    For rowNum = LBound(dataArr, 1) To UBound(dataArr, 1)
        tableRs.AddNew
        
        For fldNum = LBound(destFieldNames, 1) To UBound(destFieldNames, 1) ' перебор полей
            Set fld = tableRs.Fields(destFieldNames(fldNum))
            fld.Value = dataArr(rowNum, fldNum - LBound(destFieldNames, 1) + LBound(dataArr, 2))
        Next fldNum
        
        tableRs.Update
    Next rowNum
    tableRs.Close ' закрываем табличку
    
End Sub
