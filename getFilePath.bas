''' открывает диалог и возвращает путь к выбранному файлу
' REQUIRES: appendTo, arrayLength
Function getFilePath(Optional initFileName As String = "", Optional isFolder As Boolean = False, Optional filterList As Variant = "", Optional dialogTitle As String = "") As String
    Dim fso As Object
    Dim initFilePath As String
    Dim chk As Boolean
    Dim el As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If isFolder Then ' если работаем с папками
        If fso.FolderExists(initFileName) Then
            initFilePath = initFileName
        Else
            initFilePath = ThisWorkbook.Path + "\"
        End If
        
        ' собственно диалог
        With Application.FileDialog(msoFileDialogFolderPicker)   ' диалог - открытие файла
            .InitialFileName = Left(initFileName, InStrRev(initFileName, "\"))
            
            ' задаём диалог, елси задан
            If dialogTitle <> "" Then
                .Title = dialogTitle
            End If
            
            If .Show = True Then                         ' пользователь указал файл (проверка замены существующего файла - автоматом)
                getFilePath = .SelectedItems(1)          ' возвращаем полный путь к файлу
                If Right(getFilePath, 1) <> "\" Then
                    getFilePath = getFilePath + "\"
                End If

            Else
                getFilePath = ""
            End If
        End With
    
    Else
        ' определяем ближайшую папку, откуда будем открывать файлы
        If fso.FileExists(initFileName) Then
            initFilePath = Left(initFileName, InStrRev(initFileName, "\"))
        ElseIf fso.FolderExists(initFileName) Then
            initFilePath = initFileName
        Else
            initFilePath = ThisWorkbook.Path + "\"
        End If
    
        If Not IsArray(filterList) Then
            filterList = ""
            appendTo filterList, Array("All Excel Files", "*.xls?, *.xls")
            appendTo filterList, Array("All Files", "*.*")
        Else
            ' проверяем глубже
            chk = True
            For Each el In filterList
                If arrayLength(el) <> 2 Then chk = False
            Next
            If Not chk Then
                filterList = ""
                appendTo filterList, Array("All Excel Files", "*.xls?, *.xls")
                appendTo filterList, Array("All Files", "*.*")
            End If
        End If
    
        ' собственно диалог
        With Application.FileDialog(msoFileDialogOpen)   ' диалог - открытие файла
            .Filters.Clear
            For Each el In filterList
                .Filters.Add el(0), el(1)
            Next el
            .AllowMultiSelect = False                    ' выбираем только один файл
            .InitialFileName = initFilePath              ' путь по-умолчанию - ссылка на текущий или указанный файл
            
            ' задаём диалог, елси задан
            If dialogTitle <> "" Then
                .Title = dialogTitle
            End If
            
            If .Show = True Then                         ' пользователь указал файл (проверка замены существующего файла - автоматом)
                getFilePath = .SelectedItems(1)          ' возвращаем полный путь к файлу
            Else
                getFilePath = ""
            End If
        End With
    End If
End Function