' ������ ����������� ������, � �������� ����� �� ������ �� 1-2 ����
Function stepArray(sourceArray As Variant, Optional rowsBeginAt As Variant = 0, Optional stepForRows As Variant = 1, Optional colsBeginAt As Variant = 0, Optional stepForCols As Variant = 1) As Variant
    Dim rowNumbers As Variant, colNumbers As Variant
    Dim outArray As Variant
    Dim col As Long, row As Long
    
    stepArray = ""
    
    ' �� ����� ����� ���� ������ ��������� � ���������� �������
    If arrayLength(sourceArray, 2) > 0 Then ' ���������
        rowNumbers = stepFunction(MinMax(rowsBeginAt, LBound(sourceArray, 1), UBound(sourceArray, 1)), stepForRows, UBound(sourceArray, 1))
        colNumbers = stepFunction(MinMax(colsBeginAt, LBound(sourceArray, 2), UBound(sourceArray, 2)), stepForCols, UBound(sourceArray, 2))
        
        ReDim outArray(LBound(rowNumbers) To UBound(rowNumbers), LBound(colNumbers) To UBound(colNumbers))
        For row = LBound(rowNumbers) To UBound(rowNumbers)
            For col = LBound(colNumbers) To UBound(colNumbers)
                outArray(row, col) = sourceArray(rowNumbers(row), colNumbers(col))
            Next col
        Next row
        
    ElseIf arrayLength(sourceArray, 1) > 0 Then ' ����������
        rowNumbers = stepFunction(MinMax(rowsBeginAt, LBound(sourceArray, 1), UBound(sourceArray, 1)), stepForRows, UBound(sourceArray, 1))
    
        ReDim outArray(LBound(rowNumbers) To UBound(rowNumbers))
        For row = LBound(rowNumbers) To UBound(rowNumbers)
            outArray(row) = sourceArray(rowNumbers(row))
        Next row
    
    Else ' �������� ������ - �� ����� �� ������ ��� ������ ������ (�� 0 �� -1)
        Exit Function
    End If
    
    stepArray = outArray
    
End Function
