''' Функция для ведения журнала (в файле без ограничения длины)
''' Соглашение:
'''
''' При возникновении ошибки вызывается:
'''    addJournal "[имя функции]", "[тип ошибки]", "текст сообщения"
'''    имя функции:
'''      Для каждой функции задаём имя в переменной funcName, потом просто передаём "["+funcName+"]"
'''      Для класса также вводим className, соответственно в addJournal передаётся  "["+className+"."+funcName+"]"
'''    тип ошибки:
'''      "[Error]" - критическая ошибка, полная остановка выполнения, вывод файла журнала на экран
'''      "[Event]" - некое событие, оно пишется в лог, на экран не выводится
'''      "[Warning]" и все остальные значения - некая ошибка, выводится на экран через MsgBox, управление возвращается
Function addJournal(ParamArray items_in() As Variant) As Variant
    Dim curTimeStr As String
    Dim journal As Variant
    Dim fso As Object, file As Object
    Dim jrnPath As Variant
    Dim wf As Variant
    Dim items As Variant
    Dim shortMessage As String, fullMessage As String
    
    ' заглушка
    items = items_in ' иначе мы не можем работать с массивом как с массивом
    
    Set wf = Application.WorksheetFunction
    
    shortMessage = ""
    If arrayLength(items) = 3 Then ' если передано имя функции без квадратных скобок
        If InStr(items(0), "[") = 0 And InStr(items(0), "]") = 0 Then items(0) = "[" + items(0) + "]"
    End If
    
    fullMessage = Join(items, Chr(9)) ' всё сообщение через Табы, чтобы в экселе удобно смотреть
    If (UBound(items) - LBound(items) + 1) >= 3 Then
        shortMessage = items(2) ' 3й элемент
    End If
    
    ' формируем строку
    curTimeStr = wf.text(Now(), "yyyy-mm-dd hh:mm:ss")
    If ActiveSheet Is Nothing Then
        journal = curTimeStr & Chr(9) & "-" & Chr(9) & fullMessage
    Else
        journal = curTimeStr & Chr(9) & ActiveSheet.Name & Chr(9) & fullMessage
    End If
    
    ' пишем в конец файла всё
    Set fso = CreateObject("Scripting.FileSystemObject")
    jrnPath = ThisWorkbook.path + "\" + Mid(ThisWorkbook.Name, 1, InStrRev(ThisWorkbook.Name, ".xl") - 1) + ".journal"
    Set file = fso.OpenTextFile(jrnPath, 8, 1, 0)
    If Not file Is Nothing Then ' если открылся (никем не модифицируется)
        file.WriteLine journal
        file.Close
    Else ' если не открылся для записи
        Call MsgBox("Не удалось создать файл журнала по адресу: " + jrnPath)
        addJournal = False
    End If
    
    If InStr(journal, "[Error]") > 0 Then ' если текст содержит ошибку - открываем в блокноте
        addJournal "", "[Event]", "----------== Session Terminated ==----------"
        Shell "notepad.exe " + jrnPath, vbNormalFocus
    End If
    
    If InStr(journal, "[State]") > 0 Then ' если текст содержит статус - выводим в StatusBar'е
        Application.StatusBar = CStr(shortMessage)
    End If
    
    If shortMessage <> "" And InStr(fullMessage, "[Event]") = 0 And InStr(fullMessage, "[State]") = 0 Then ' печатаем сообщение
        addJournal = MsgBox(shortMessage)
    End If
    
    Set file = Nothing
    Set fso = Nothing
End Function