''' переливает данные из одной базы в другую
Sub db_DBToDB(srcConn As ADODB.Connection, srcTblName As String, destConn As ADODB.Connection, destTblName As String, Optional srcFldsList As Variant = "", Optional destFldsList As Variant = "", Optional srcQryCondition As String = "")

    Dim srcQueryStr As String
    Dim srcTableRs As New ADODB.Recordset
    Dim destQueryStr As String
    Dim destTableRs As New ADODB.Recordset
    
    Dim srcFieldNamesTmp As Collection
    Dim destFieldNamesTmp As Collection
    Dim srcFieldNamesArr As Variant
    Dim destFieldNamesArr As Variant
    
    Dim chk1 As Boolean
    Dim chk2 As Boolean
    Dim chk3 As Boolean
    Dim chk4 As Boolean
    Dim chk5 As Boolean
    Dim chk6 As Boolean
    Dim chk7 As Boolean
    Dim chk8 As Boolean
    
    Dim srcFieldNamesList As Collection
    Dim destFieldNamesList As Collection
    Dim value As Variant
    
    Dim el As Variant
    Dim i As Long
    Dim srcFldName As String
    Dim destFldName As String
    
    ' ///// Проверки
    ' преобразуем список полей в массив
    If TypeName(srcFldsList) = "Collection" Then
        Set srcFieldNamesTmp = srcFldsList
        srcFieldNamesArr = col2Array(srcFieldNamesTmp) ' 0-based array
    Else
        srcFieldNamesArr = srcFldsList
    End If
    
    ' преобразуем список полей в массив
    If TypeName(destFldsList) = "Collection" Then
        Set destFieldNamesTmp = destFldsList
        destFieldNamesArr = col2Array(destFieldNamesTmp) ' 0-based array
    Else
        destFieldNamesArr = destFldsList
    End If
    
    chk1 = InStr(TypeName(srcFieldNamesArr), "()") > 0
    chk2 = InStr(TypeName(destFieldNamesArr), "()") > 0
    chk3 = (TypeName(srcFieldNamesArr) = "String")
    chk4 = (TypeName(destFieldNamesArr) = "String")
    
    ' можем работать только когда либо оба массивы, либо оба не заданы
    chk5 = (chk1 And chk2) Or (chk3 And chk4)
    
    If Not chk5 Then
        MsgBox "[db_DbToDb] Должны быть заданы или оба списка полей, или ни одного"
        Exit Sub
    End If
    
    ' доп проверка для массивов - совпадение размерностей
    If chk1 And chk2 Then
        chk6 = (LBound(srcFieldNamesArr) = LBound(destFieldNamesArr))
        chk7 = (UBound(srcFieldNamesArr) = UBound(destFieldNamesArr))
    Else
        chk6 = (Len(srcFieldNamesArr) = 0)
        chk7 = (Len(destFieldNamesArr) = 0)
    End If
    
    If chk1 And chk2 And Not (chk6 And chk7) Then
        MsgBox "[db_DbToDb] Список полей в источнике и получателе должен содержать одинаковое число элементов"
        Exit Sub
    End If
    
    If Not (chk1 And chk2) And Not (chk6 And chk7) Then
        MsgBox "[db_DbToDb] Ошибка в одном из списков полей"
        Exit Sub
    End If
    
    ' \ проверки завершены

    ' соединение с источником
    srcQueryStr = "SELECT * FROM " + srcTblName + IIf(srcQryCondition <> "", " WHERE " + srcQryCondition, "")
    srcTableRs.ActiveConnection = srcConn.ConnectionString
    srcTableRs.LockType = adLockOptimistic
    srcTableRs.source = srcQueryStr
    
    ' соединение с приёмником
    destQueryStr = "SELECT * FROM " + destTblName
    destTableRs.ActiveConnection = destConn.ConnectionString
    destTableRs.LockType = adLockOptimistic
    destTableRs.source = destQueryStr
    
    
    
    srcTableRs.Open ' открываем табличку-источник
    destTableRs.Open ' открываем табличку-приёмник
    
    If Not (chk1 And chk2) Then ' массив полей не задан - проверяем, что количество полей в источнике и полчателе равно
        chk8 = (srcTableRs.Fields.Count = destTableRs.Fields.Count)
    Else
        chk8 = True
    End If
    
    If Not srcTableRs.EOF Then ' встали на первую запись
        srcTableRs.MoveFirst
    End If
    
    
    ' готовим унифицированный список полей (из массивов или из самих таблиц)
    ' ранее уже проверили, что в таблицах число полей совпадает
    If chk1 And chk2 Then ' задан массив
        Set srcFieldNamesList = array2col(srcFieldNamesArr)
        Set destFieldNamesList = array2col(destFieldNamesArr)
    Else ' работаем со списком полей из самой таблицы
        Set srcFieldNamesList = New Collection
        Set destFieldNamesList = New Collection
        
        For Each el In srcTableRs.Fields
            srcFieldNamesList.Add CStr(el.Name), CStr(el.Name)
        Next el
        
        For Each el In destTableRs.Fields
            destFieldNamesList.Add CStr(el.Name), CStr(el.Name)
        Next el
    End If
    
    Do While (Not srcTableRs.EOF) And chk8
        destTableRs.AddNew
        
        For i = 1 To srcFieldNamesList.Count ' перебираем список полей
            srcFldName = srcFieldNamesList.item(i)
            destFldName = destFieldNamesList.item(i)
            
            ' искренне надеемся, что типы данных в источнике и получателе совпадают
            Select Case destTableRs.Fields(destFldName).Type
            Case DataTypeEnum.adBSTR ' MEMO
                value = CStr(srcTableRs.Fields(srcFldName).value)
            Case Else
                value = srcTableRs.Fields(srcFldName).value
            End Select
            
            destTableRs.Fields(destFldName).value = value
        Next i
        
        destTableRs.Update
        srcTableRs.MoveNext
    Loop
    
    destTableRs.Close ' закрываем табличку-приёмник
    srcTableRs.Close ' закрываем табличку-источник
    
    

End Sub
