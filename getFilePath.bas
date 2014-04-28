''' открывает диалог и возвращает путь к выбранному файлу
Function getFilePath(Optional initFileName As String = "", Optional isFolder As Boolean = False) As String
    Dim fso As Object
    Dim initFilePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If isFolder Then ' если работаем с папками
        If fso.FolderExists(initFileName) Then
            initFilePath = initFileName
        Else
            initFilePath = ActiveWorkbook.path + "\"
        End If
        
        ' собственно диалог
        With Application.FileDialog(msoFileDialogFolderPicker)   ' диалог - открытие файла
            .InitialFileName = Left(initFileName, InStrRev(initFileName, "\"))
            
            If .Show = True Then                         ' пользователь указал файл (проверка замены существующего файла - автоматом)
                getFilePath = .SelectedItems(1) + "\"         ' возвращаем полный путь к файлу
            Else
                getFilePath = ""
            End If
        End With
    
    Else
        ' определяем ближайшую папку, откуда будем открывать файлы
        If fso.FileExists(initFileName) Then
            initFilePath = Left(initFileName, InStrRev(initFileName, "\"))
        Else
            initFilePath = ActiveWorkbook.path + "\"
        End If
    
        ' собственно диалог
        With Application.FileDialog(msoFileDialogOpen)   ' диалог - открытие файла
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
            .Filters.Add "All Excel Files", "*.xls*"
            .AllowMultiSelect = False                    ' выбираем только один файл
            .InitialFileName = initFilePath              ' путь по-умолчанию - ссылка на текущий или указанный файл
            
            If .Show = True Then                         ' пользователь указал файл (проверка замены существующего файла - автоматом)
                getFilePath = .SelectedItems(1)          ' возвращаем полный путь к файлу
            Else
                getFilePath = ""
            End If
        End With
    End If
End Function