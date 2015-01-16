''' ��������� ������ � ���������� ���� � ���������� �����
' REQUIRES: appendTo, arrayLength
Function getFilePath(Optional initFileName As String = "", Optional isFolder As Boolean = False, Optional filterList As Variant = "", Optional dialogTitle As String = "") As String
    Dim fso As Object
    Dim initFilePath As String
    Dim chk As Boolean
    Dim el As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If isFolder Then ' ���� �������� � �������
        If fso.FolderExists(initFileName) Then
            initFilePath = initFileName
        Else
            initFilePath = ThisWorkbook.Path + "\"
        End If
        
        ' ���������� ������
        With Application.FileDialog(msoFileDialogFolderPicker)   ' ������ - �������� �����
            .InitialFileName = Left(initFileName, InStrRev(initFileName, "\"))
            
            ' ����� ������, ���� �����
            If dialogTitle <> "" Then
                .Title = dialogTitle
            End If
            
            If .Show = True Then                         ' ������������ ������ ���� (�������� ������ ������������� ����� - ���������)
                getFilePath = .SelectedItems(1)          ' ���������� ������ ���� � �����
                If Right(getFilePath, 1) <> "\" Then
                    getFilePath = getFilePath + "\"
                End If

            Else
                getFilePath = ""
            End If
        End With
    
    Else
        ' ���������� ��������� �����, ������ ����� ��������� �����
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
            ' ��������� ������
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
    
        ' ���������� ������
        With Application.FileDialog(msoFileDialogOpen)   ' ������ - �������� �����
            .Filters.Clear
            For Each el In filterList
                .Filters.Add el(0), el(1)
            Next el
            .AllowMultiSelect = False                    ' �������� ������ ���� ����
            .InitialFileName = initFilePath              ' ���� ��-��������� - ������ �� ������� ��� ��������� ����
            
            ' ����� ������, ���� �����
            If dialogTitle <> "" Then
                .Title = dialogTitle
            End If
            
            If .Show = True Then                         ' ������������ ������ ���� (�������� ������ ������������� ����� - ���������)
                getFilePath = .SelectedItems(1)          ' ���������� ������ ���� � �����
            Else
                getFilePath = ""
            End If
        End With
    End If
End Function