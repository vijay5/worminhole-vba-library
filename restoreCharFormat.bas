''' ��������������� ���� ���� �� �������
Sub restoreCharFormat(cl As Range, colorArr As Variant)
    Dim k As Long
    Dim maxLen As Long
    
    If IsArray(colorArr) Then ' ���� ���� ����� ��� ����
        maxLen = WorksheetFunction.Min(arrayLength(colorArr), Len(cl.value)) ' ������, �� �������� ����� ��������� ������
        
        For k = 1 To maxLen
            cl.Characters(start:=k, Length:=1).Font.color = colorArr(k)
        Next k
    Else
        ' pass
    End If
    
End Sub