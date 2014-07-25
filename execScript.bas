' REQUIRES: addJournal
Function execScript(scriptStr As String, Optional objects As Collection = Nothing, Optional sc As Object = Nothing) As Variant
    Dim obj As Variant
    Dim objName As String
    Dim objObject As Object
    Dim tmpVal As Variant
    Dim sc1 As New MSScriptControl.ScriptControl
    
    ' инициализируем объект
    If sc Is Nothing Then
        Set sc = CreateObject("MSScriptControl.ScriptControl") ' для вычисления выражений типа "1+2+Sin(1)"
        sc.Language = "VBScript"
    End If
    tmpVal = 0
    
    ' передаём ссылки на объекты (если они заданы)
    On Error GoTo equError1
        For Each obj In objects
            objName = obj(0) ' имя переменной
            If IsObject(obj(1)) Then ' если задан объект - передаём объект
                Set objObject = obj(1)
                Call sc.AddObject(objName, objObject)
            Else ' иначе - присваиваем значение переменной
                Call sc.ExecuteStatement(CStr(obj(0)) & " = " & CStr(obj(1)))
            End If
equError1:
            If sc.Error.Number <> 0 Then ' ошибка произошла :)
                addJournal "[execScript]", "[Warning]", "Не удалось передать объект / выполнить оператор для переменной: " & CStr(obj(0))
            End If
            
        Next obj
    On Error GoTo 0
    
    On Error GoTo equError2
        sc.Error.Clear ' чистим ошибки
        tmpVal = sc.eval(scriptStr) ' вычисляем вес
        
equError2:
        If sc.Error.Number <> 0 Then ' ошибка произошла :)
            addJournal "[execScript]", "[Warning]", "Не удалось вычислить значение выражения: " & scriptStr
            tmpVal = 0 ' по-умолчанию, при ошибке будет нулевой вес
        End If
    On Error GoTo 0
    
    execScript = tmpVal
End Function