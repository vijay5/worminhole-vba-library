''' ��� ������ Range.Value �� Range, ������� ������� �� ����� ������, �� ��������� ������������ �������� ������, � �� ������
''' ��� ������� ���������� ������ ������.
Function getProperValArray(rng As Range) As Variant
    Dim tmpArray As Variant
    
    If rng.Cells.Count = 1 Then
        ReDim tmpArray(1 To 1, 1 To 1)
        tmpArray(1, 1) = rng.value
        getProperValArray = tmpArray
    Else
        getProperValArray = rng.value
    End If
    
End Function
