''' вовзращает 2D-массив, если есть данные в таблице
''' если данных нет - возвращает пустую строку
' REQUIRES: transposeArr, col2Array, isInCollection
Function db_DBToExcel(conn As ADODB.Connection, srcTblName As String, Optional srcFieldNames As Variant = "*", Optional fieldsCol As Collection) As Variant
    Dim srcFieldNamesArr As Variant
    Dim tableRs As New ADODB.Recordset
    Dim tmp As Variant
    Dim fldName As String
    Dim fldNameAlpha As String
    Dim fldNum As Long
    Dim queryStr As String
    Dim noFldsList As Collection
    Dim existingFldsList As Collection
    Dim fldsListStr As String
    Dim srcFieldNamesTmp As Collection
    Dim i As Long
    
    ' по умолчанию - нет данных
    db_DBToExcel = ""


    ' проверка существования укаанных полей
    If Not IsArray(srcFieldNames) Then
        fldsListStr = "*"
    Else
        ' часть 1 - проверка наличия всех полей в БД
        ' устанавливаем соединение с БД
        queryStr = "SELECT * FROM " + srcTblName + ";"
        tableRs.ActiveConnection = conn.ConnectionString
        tableRs.LockType = adLockOptimistic
        tableRs.source = queryStr
        tableRs.Open ' открываем табличку

        
        ' преобразуем список полей в массив
        If TypeName(srcFieldNames) = "Collection" Then
        Set srcFieldNamesTmp = srcFieldNames
            srcFieldNamesArr = col2Array(srcFieldNamesTmp) ' 0-based array
        Else
            srcFieldNamesArr = srcFieldNames
        End If
    
        ' проверка наличия всех полей в таблице
        Set noFldsList = New Collection
        Set existingFldsList = New Collection
        
        For fldNum = LBound(srcFieldNamesArr, 1) To UBound(srcFieldNamesArr, 1) ' перебор полей
            tmp = "abrakadabra"
            On Error Resume Next
                fldName = srcFieldNamesArr(fldNum)
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
            MsgBox "[db_excelToDB] В таблице отсутствуют поля: " + Join(col2Array(noFldsList), ", ")
            Exit Function
        Else
            ' pass
        End If
    
        fldsListStr = Join(col2Array(existingFldsList), ", ")
        tableRs.Close ' закрываем табличку
    End If
    
    
    ' в этой точке мы проверили, что все поля существуют
    queryStr = "SELECT " + fldsListStr + " FROM " + srcTblName + ";"
    tableRs.ActiveConnection = conn.ConnectionString
    tableRs.LockType = adLockOptimistic
    tableRs.source = queryStr
    tableRs.Open ' открываем табличку
    
        ' переливаем список полей в коллекцию
        Set fieldsCol = Nothing
        Set fieldsCol = New Collection
        For i = 1 To tableRs.Fields.Count
            fieldsCol.Add tableRs.Fields.item(i - 1).Name
        Next i
        
        ' переливаем собственно данные
        If Not tableRs.EOF Then
            db_DBToExcel = transposeArr(tableRs.GetRows) ' транспонируем
        Else
            db_DBToExcel = ""
        End If
    tableRs.Close ' закрываем табличку
    
End Function
