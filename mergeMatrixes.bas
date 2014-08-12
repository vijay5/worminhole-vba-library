''' ��������� ���������/2D ������� � ���������� ���� ������� 2D-�������
' REQUIRES: arrayLength
Function mergeMatrixes(direction As XlDirection, ParamArray rangeArray() As Variant) As Variant
    Dim rng As Variant
    Dim maxRow As Long
    Dim maxCol As Long
    Dim colNum As Long
    Dim rowNum As Long
    Dim maxRowPrev As Long
    Dim maxColPrev As Long
    Dim outArr As Variant
    Dim tmpArr As Variant
    Dim rngArr As Variant
    Dim rangeValueArray As Variant ' ������ ������ (rng -> rng.Value)
    
    Dim startCnt As Long
    Dim endCnt As Long
    Dim stepCnt As Long
    Dim i As Long
    
    Dim tmpDim2 As Single
    Dim tmpDim3 As Single
    Dim chk3 As Boolean
    Dim outRowNum As Long
    Dim outColNum As Long
    
    mergeMatrixes = "" ' ��-��������� - ���������� ������
    ' � ����������� �� ����������� ����� ��������� �������� ������ ������
    Select Case direction
    Case XlDirection.xlDown, XlDirection.xlToRight ' ������ ����
        startCnt = LBound(rangeArray)
        endCnt = UBound(rangeArray)
        stepCnt = 1
        
    Case XlDirection.xlUp, XlDirection.xlToLeft   ' ����� �����
        startCnt = UBound(rangeArray)
        endCnt = LBound(rangeArray)
        stepCnt = -1
        
    Case Else
        Exit Function
    End Select
    
    
    ' ����������� ��������� � ������ ������
    ' ������������ ���������, ����� � �������� ���� ����� 2 ���
    ReDim rangeValueArray(LBound(rangeArray) To UBound(rangeArray))
    
    For i = startCnt To endCnt Step stepCnt
        If TypeName(rangeArray(i)) = "Range" Then ' �� ����� - ��������
            ' ��������� ����������� �������� � 2D-������
            If rangeArray(i).Cells.Count = 1 Then
                ReDim tmpArray(1 To 1, 1 To 1)
                tmpArray(1, 1) = rangeArray(i).value
                rangeValueArray(i) = tmpArray
            Else
                rangeValueArray(i) = rangeArray(i).value
            End If
            
            
        ElseIf InStr(TypeName(rangeArray(i)), "()") > 0 Then ' �� ����� - ������
            
            ' / �������� ����������� �������
            tmpDim2 = 0.5
            tmpDim3 = 0.5
            On Error Resume Next
                tmpDim2 = UBound(rangeArray(i), 2)
                tmpDim3 = UBound(rangeArray(i), 3)
            On Error GoTo 0
            chk3 = (tmpDim2 <> 0.5) And (tmpDim3 = 0.5)
            
            If chk3 Then
                rangeValueArray(i) = rangeArray(i)
                
            Else ' ����������� ������� 1 ��� >= 3
                MsgBox "�� ����� ������ ���� ������ 2D-������"
                Exit Function
            End If
        Else
            ' pass
        End If
    Next i
    
    ' � ���� ����� ��� ��������� ������������� � �������
    
    
    ' ������� �������, ������� �����/�������� �����
    maxRow = 0
    maxCol = 0
    For i = startCnt To endCnt Step stepCnt
        Select Case direction
        Case XlDirection.xlDown, XlDirection.xlUp ' �����-����
            maxRow = maxRow + arrayLength(rangeValueArray(i), 1)
            maxCol = WorksheetFunction.MAX(maxCol, arrayLength(rangeValueArray(i), 2))
        
        Case XlDirection.xlToRight, XlDirection.xlToLeft   ' �����-������
            maxRow = WorksheetFunction.MAX(maxRow, arrayLength(rangeValueArray(i), 1))
            maxCol = maxCol + arrayLength(rangeValueArray(i), 2)
        
        Case Else
            Exit Function
            
        End Select

    Next i
    
    ' �������� �����, ������� �������������� rebase � (1, 1)
    ReDim outArr(1 To maxRow, 1 To maxCol)
    maxRow = LBound(outArr, 1) - 1
    maxCol = LBound(outArr, 2) - 1
    
    ' ������� ������
    For i = startCnt To endCnt Step stepCnt ' ���� �� ����������
        rngArr = rangeValueArray(i)
        Select Case direction
        Case XlDirection.xlDown, XlDirection.xlUp ' �����-����
            maxRowPrev = maxRow + 1
            maxColPrev = LBound(outArr, 2)
            
            maxRow = maxRow + arrayLength(rngArr, 1)
            maxCol = WorksheetFunction.MAX(maxCol, arrayLength(rngArr, 2))
        
        Case XlDirection.xlToRight, XlDirection.xlToLeft   ' �����-������
            maxRowPrev = LBound(outArr, 1)
            maxColPrev = maxCol + 1
            
            maxRow = WorksheetFunction.MAX(maxRow, arrayLength(rngArr, 1))
            maxCol = maxCol + arrayLength(rngArr, 2)
        
        Case Else
            Exit Function
            
        End Select
        
        
        ' ���������� �������
        For rowNum = LBound(rngArr, 1) To UBound(rngArr, 1)
            For colNum = LBound(rngArr, 2) To UBound(rngArr, 2)
                outRowNum = maxRowPrev + rowNum - LBound(rngArr, 1)
                outColNum = maxColPrev + colNum - LBound(rngArr, 2)
                outArr(outRowNum, outColNum) = rngArr(rowNum, colNum)
            Next colNum
        Next rowNum
        
    Next i
    
    mergeMatrixes = outArr ' ���������� ������
    
End Function
