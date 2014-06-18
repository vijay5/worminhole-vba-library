''' Просматривает заданное количество вложенных папок от текущего адреса
''' Возвращает коллекцию путей к файлам, имя которых совпадает с маской
Function getFileList(ByVal folderPath As String, Optional ByVal mask As String = "", _
                             Optional ByVal maxSearchDepth As Long = 1) As Collection
   ' Получает в качестве параметра путь к папке FolderPath,
   ' маску имени искомых файлов Mask (будут отобраны только файлы с такой маской/расширением)
   ' и глубину поиска maxSearchDepth в подпапках (если maxSearchDepth=1, то подпапки не просматриваются).
   ' Возвращает коллекцию, содержащую полные пути найденных файлов
   ' (применяется рекурсивный вызов процедуры GetAllFileNamesUsingFSO)
    Dim fso As Object
    Dim fileList As New Collection

    Set fileList = New Collection    ' создаём пустую коллекцию
    GetAllFileNamesUsingfso folderPath, mask, fso, fileList, maxSearchDepth ' поиск
    Set fso = Nothing
    
    
    Set getFileList = fileList

End Function

Sub getAllFileNamesUsingfso(ByVal folderPath As String, ByVal mask As String, ByRef fso As Object, ByRef fileList As Collection, ByVal maxSearchDepth As Long)
    Dim curfold As Variant
    Dim fil As Variant
    Dim sfol As Variant
    ' перебирает все файлы и подпапки в папке folderPath, используя объект fso
    ' перебор папок осуществляется в том случае, если maxSearchDepth > 1
    ' добавляет пути найденных файлов в коллекцию fileList
    'On Error Resume Next
        If fso Is Nothing Then
            Set fso = CreateObject("Scripting.FileSystemObject")
        End If
        Set curfold = fso.GetFolder(folderPath)
        If Not curfold Is Nothing Then    ' если удалось получить доступ к папке
            ' раскомментируйте эту строку для вывода пути к просматриваемой
            ' в текущий момент папке в строку состояния Excel
            ' Application.StatusBar = "Поиск в папке: " & folderPath
            
            For Each fil In curfold.Files    ' перебираем все файлы в папке folderPath
                If fil.Name Like mask Then fileList.Add fil.path
            Next fil
            maxSearchDepth = maxSearchDepth - 1    ' уменьшаем глубину поиска в подпапках
            If maxSearchDepth > 0 Then    ' если надо искать глубже
                For Each sfol In curfold.SubFolders    ' перебираем все подпапки в папке folderPath
                    GetAllFileNamesUsingfso sfol.path, mask, fso, fileList, maxSearchDepth
                Next sfol
            End If
            Set fil = Nothing
            Set curfold = Nothing    ' очищаем переменные
        End If
    'On Error GoTo 0
End Sub
