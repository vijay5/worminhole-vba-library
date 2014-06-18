''' ������������� �������� ���������� ��������� ����� �� �������� ������
''' ���������� ��������� ����� � ������, ��� ������� ��������� � ������
Function getFileList(ByVal folderPath As String, Optional ByVal mask As String = "", _
                             Optional ByVal maxSearchDepth As Long = 1) As Collection
   ' �������� � �������� ��������� ���� � ����� FolderPath,
   ' ����� ����� ������� ������ Mask (����� �������� ������ ����� � ����� ������/�����������)
   ' � ������� ������ maxSearchDepth � ��������� (���� maxSearchDepth=1, �� �������� �� ���������������).
   ' ���������� ���������, ���������� ������ ���� ��������� ������
   ' (����������� ����������� ����� ��������� GetAllFileNamesUsingFSO)
    Dim fso As Object
    Dim fileList As New Collection

    Set fileList = New Collection    ' ������ ������ ���������
    GetAllFileNamesUsingfso folderPath, mask, fso, fileList, maxSearchDepth ' �����
    Set fso = Nothing
    
    
    Set getFileList = fileList

End Function

Sub getAllFileNamesUsingfso(ByVal folderPath As String, ByVal mask As String, ByRef fso As Object, ByRef fileList As Collection, ByVal maxSearchDepth As Long)
    Dim curfold As Variant
    Dim fil As Variant
    Dim sfol As Variant
    ' ���������� ��� ����� � �������� � ����� folderPath, ��������� ������ fso
    ' ������� ����� �������������� � ��� ������, ���� maxSearchDepth > 1
    ' ��������� ���� ��������� ������ � ��������� fileList
    'On Error Resume Next
        If fso Is Nothing Then
            Set fso = CreateObject("Scripting.FileSystemObject")
        End If
        Set curfold = fso.GetFolder(folderPath)
        If Not curfold Is Nothing Then    ' ���� ������� �������� ������ � �����
            ' ���������������� ��� ������ ��� ������ ���� � ���������������
            ' � ������� ������ ����� � ������ ��������� Excel
            ' Application.StatusBar = "����� � �����: " & folderPath
            
            For Each fil In curfold.Files    ' ���������� ��� ����� � ����� folderPath
                If fil.Name Like mask Then fileList.Add fil.path
            Next fil
            maxSearchDepth = maxSearchDepth - 1    ' ��������� ������� ������ � ���������
            If maxSearchDepth > 0 Then    ' ���� ���� ������ ������
                For Each sfol In curfold.SubFolders    ' ���������� ��� �������� � ����� folderPath
                    GetAllFileNamesUsingfso sfol.path, mask, fso, fileList, maxSearchDepth
                Next sfol
            End If
            Set fil = Nothing
            Set curfold = Nothing    ' ������� ����������
        End If
    'On Error GoTo 0
End Sub
