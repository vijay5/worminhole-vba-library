''' ��������� ������� � ���������� �����
Function fileChecker(tbox As Object) As Integer
    Dim fso As Object
    Dim ts As Object
    Dim chk1 As Boolean, chk2 As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject") ' ������� ����������
    
    If tbox.text <> "" Then ' ���� ������ � �������� ��������
        chk1 = fso.FileExists(tbox.text)  ' ���� ����������
        chk2 = False                      ' � ����� ���� ������ ��� ������
        If chk1 Then
            Set ts = Nothing
            On Error Resume Next
            Set ts = fso.OpenTextFile(tbox.text, 1) ' 1 = for reading
            On Error GoTo 0
            chk2 = Not (ts Is Nothing)
            ts.Close
            Set ts = Nothing
        End If
        fileChecker = IIf(chk1 And chk2, 1, -1)
    Else ' ���� �� �����
        fileChecker = 0
    End If
End Function
