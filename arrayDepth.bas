''' ���������� ����� ��������� � ��������� �������. ���� �� ������ - ���������� 0
Function arrayDepth(arr As Variant) As Byte
    Dim i as Integer
    Dim tmp As Single
    If InStr(TypeName(arr), "()") > 0 Then ' ����� ���� ������
        On Error Resume Next
            For i = 1 To 256 ' ���� �� ����� ���������
                tmp = -1.5
                tmp = UBound(arr, i)
                If tmp <> -1.5 Then
                    arrayDepth = i
                Else
                    Exit For
                End If
            Next i
        On Error GoTo 0
    Else ' ����� ���� �� ������
        arrayDepth = 0
    End If
End Function
