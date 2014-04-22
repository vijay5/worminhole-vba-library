' ����� �� ����� ������ ��� ������� (interval) + ������ ������� �� ����������
' � ������ � ������� ��������� ����������
' ���� ������ � ���������� ���������� � �� ���� ������ ���������� �������� � �������
Function findPlaceInInterval(value As Variant, interval As Range) As Variant
    Dim r As Long
    findPlaceInInterval = -1 ' �� ���������
    
    If interval.Columns.Count <> 3 Then Exit Function
    
    For r = 1 To interval.Rows.Count
        If interval.Cells(r, 2).value <= value And value <= interval.Cells(r, 3).value Then
            findPlaceInInterval = interval.Cells(r, 1).value
            Exit For
        End If
    Next r
End Function