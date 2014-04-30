''' ��������������� ��������� �������
''' fromChainArray - ����������� ��������� �������
'''    true  - �� 1D*1D � 2D
'''    false - �� 2D � 1D*1D
''' minRow - ����� ������ ������ � ����� �������
''' minCol - ����� ������� ������� � ����� �������
Function regroupArray(arr As Variant, Optional fromChainArray As Boolean = True, Optional minRow As Long = 0, Optional minCol As Long = 0) As Variant
    
    Dim i As Long, j As Long
    Dim resultArray As Variant
    Dim maxRow As Long, maxCol As Long
    Dim curRow As Long, curCol As Long
    
    Select Case fromChainArray
    Case True ' �� 1D*1D � 2D
        maxRow = minRow + arrayLength(arr, 1) - 1
        maxCol = minCol + arrayLength(arr(LBound(arr)), 1) - 1 ' ���� ������ �� ������ ������ (����� �� ����� ������� - ��������� �� ���� ���� �� �������)
        ReDim resultArray(minRow To maxRow, minCol To maxCol)
        
        curRow = minRow - 1
        For i = LBound(arr) To UBound(arr) ' ���� �� �������
            curRow = curRow + 1
            curCol = minCol - 1
            For j = LBound(arr(i)) To UBound(arr(i)) ' ���� �� ��������
                curCol = curCol + 1
                resultArray(curRow, curCol) = arr(i)(j)
            Next j
        Next i
        
    Case False ' �� 2D � 1D*1D
        maxRow = minRow + arrayLength(arr, 1) - 1
        maxCol = minCol + arrayLength(arr, 2) - 1
    
        ReDim resultArray(minRow To maxRow)
        curRow = minRow - 1
        For i = LBound(arr, 1) To UBound(arr, 1) ' ���� �� �������
            ReDim rowArray(minCol To maxCol)
            curRow = curRow + 1
            curCol = minCol - 1
            For j = LBound(arr, 2) To UBound(arr, 2) ' ���� �� ��������
                curCol = curCol + 1
                rowArray(curCol) = arr(i, j)
            Next j
            resultArray(curRow) = rowArray
        Next i
    
    End Select
    
    regroupArray = resultArray
    
End Function
