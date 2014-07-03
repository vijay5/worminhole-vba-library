Function getFileName(fullPath As String, Optional getPath As Boolean = False) As String
    Dim pos As Long
    
    pos = InStrRev(fullPath, "\")
    If getPath Then
        getFileName = IIf(pos > 0, Left(fullPath, pos), fullPath) ' � ������� ������� �� ���������� �����
    Else
        getFileName = IIf(pos > 0, Mid(fullPath, pos + 1), fullPath) ' �� ���������� ������� ����� ���������� ����� �� ����� ������
    End If
End Function