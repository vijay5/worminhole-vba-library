''' ��������� ������ � ���������� ���� � ���������� �����
Function getFilePath(Optional initFileName As String = "", Optional isFolder As Boolean = False) As String
    Dim fso As Object
    Dim initFilePath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If isFolder Then ' ���� �������� � �������
        If fso.FolderExists(initFileName) Then
            initFilePath = initFileName
        Else
            initFilePath = ActiveWorkbook.path + "\"
        End If
        
        ' ���������� ������
        With Application.FileDialog(msoFileDialogFolderPicker)   ' ������ - �������� �����
            .InitialFileName = Left(initFileName, InStrRev(initFileName, "\"))
            
            If .Show = True Then                         ' ������������ ������ ���� (�������� ������ ������������� ����� - ���������)
                getFilePath = .SelectedItems(1) + "\"         ' ���������� ������ ���� � �����
            Else
                getFilePath = ""
            End If
        End With
    
    Else
        ' ���������� ��������� �����, ������ ����� ��������� �����
        If fso.FileExists(initFileName) Then
            initFilePath = Left(initFileName, InStrRev(initFileName, "\"))
        Else
            initFilePath = ActiveWorkbook.path + "\"
        End If
    
        ' ���������� ������
        With Application.FileDialog(msoFileDialogOpen)   ' ������ - �������� �����
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
            .Filters.Add "All Excel Files", "*.xls*"
            .AllowMultiSelect = False                    ' �������� ������ ���� ����
            .InitialFileName = initFilePath              ' ���� ��-��������� - ������ �� ������� ��� ��������� ����
            
            If .Show = True Then                         ' ������������ ������ ���� (�������� ������ ������������� ����� - ���������)
                getFilePath = .SelectedItems(1)          ' ���������� ������ ���� � �����
            Else
                getFilePath = ""
            End If
        End With
    End If
End Function