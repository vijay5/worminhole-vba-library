' REQUIRES: arrayDepth
''' ����������� ����� ������� �� ������-���� �� ���������
Function arrayLength(arr As Variant, Optional degree As Byte = 1) As Long
    arrayLength = 0 ' �� ���������
    If InStr(TypeName(arr), "()") > 0 Then
        If arrayDepth(arr) >= 1 Then ' ����� ���� ������
            arrayLength = UBound(arr, degree) - LBound(arr, degree) + 1
        End If
    End If
End Function