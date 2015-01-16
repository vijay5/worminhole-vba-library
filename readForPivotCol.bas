''' TODO: ����� ����� ������� �� ���������, � ������ � ������, � ����� �������� ���� ������ �������...
' REQUIRES: MatrixPart, getProperValArray, getFlatArray, mergeVectors, arrayLength
Function ReadForPivotCol(sourceRange As Range, numOfRowProperties As Integer, numOfColProperties As Integer, Optional numOfRowParams As Integer = 1, Optional numOfColParams As Integer = 1, Optional getNonEmptyOnly As Boolean = False) As Collection
    ''' ����������� ������� ���� � ������� �������
    ''' ������� ��� �������� �������:
    '''
    ''' [*]                +-----------------------+-----------------------+
    '''                    |Var1-1                 |Var1-2                 |
    '''                    +-------+-------+-------+-------+-------+-------+
    '''   Col1  Col2  ColN |Var2-1 |Var2-2 |Var2-3 |Var2-1 |Var2-2 |Var2-3 |
    '''  +-----+-----+-----+-------+-------+-------+-------+-------+-------+
    '''  |     |     |     |TLCell | ...   | ...   | ...   | ...   | ...   |
    '''
    ''' ���
    ''' [*] - ����� ������� ���� ����� � �������
    ''' Col1-ColN - ����� ������������� ���������� (numOfHeaderCols)
    ''' Var1 - "������������" ���������� (������ �� ����, ����� ���������� ������������)
    ''' Var2 - "������������" ���������� (������ �� ����, ��������� ������)
    ''' TLCell - ������ ������ � �������
    ''' ����� �����, ��� ���-�� ����� ��� TLCell ����� ���-�� "������������ ����������" (���� ������� ���������� �����)
    '''
    ''' �� ������ ����� ��������� "�������" �������
    '''  Col1  Col2  ColN ||Var1  Var2 || Value
    ''' +-----+-----+-----++-----+-----++-------+
    ''' |     |     |     ||     |     ||       |

    Dim height0 As Long, height As Long, width0 As Long, width As Long
    Dim rowLabels As Variant, colLabels As Variant, sourceData As Variant
    Dim minRow As Long, maxRow As Long
    Dim sourceSheet As Variant
    Dim resultSheet As Worksheet
                                  
    Dim rowHdrCol As Collection
    Dim colHdrCol As Collection
    Dim dataCol As Collection
    Dim rowNum As Long
    Dim colNum As Long
    Dim tmpArr As Variant
    Dim cellArr As Variant
    Dim cellFlatArr As Variant
    Dim rowFlatArr As Variant
    Dim colFlatArr As Variant
    Dim numOfRecords As Long
    Dim i As Long
    Dim arrIsEmpty As Variant
    
    ' �������� �� ��������� ���������
    Set sourceSheet = sourceRange.Parent
    height0 = sourceRange.Rows.Count
    width = sourceRange.Columns.Count
    
    If height0 <= numOfColProperties Or width <= numOfRowProperties Then
       Call MsgBox("�������� ������� ��������� �������")
       Exit Function
    End If
    
    ' TODO: ����� ������� �������� ��������������� Unmerge'� ����� � ���������� �� ����������
    
    ' �������� �������� � 2D-������
    rowLabels = getProperValArray(sourceRange.Cells(1 + numOfColProperties, 1).Resize(height0 - numOfColProperties, numOfRowProperties))
    colLabels = getProperValArray(sourceRange.Cells(1, 1 + numOfRowProperties).Resize(numOfColProperties, width - numOfRowProperties))
    
    ' �������� � ��������� �������
    sourceData = getProperValArray(sourceRange.Resize(height0 - numOfColProperties, width - numOfRowProperties). _
                                               Offset(numOfColProperties, numOfRowProperties))
                                               
    ' �������� ��������
    height0 = arrayLength(sourceData, 1) ' ����� ����� � �������� �������
    height = height0 \ numOfRowParams    ' ����� ����� �� ���� ������ ���������� �������
    width0 = arrayLength(sourceData, 2)  ' ����� �������� � �������� �������
    width = width0 \ numOfColParams      ' ����� �������� �� ���� ������ ���������� �������
    
    If width * numOfColParams <> width0 Then
        MsgBox "������ � ������� �� ������ numOfColParams"
        Exit Function
    End If
                                  
    If height * numOfRowParams <> height0 Then
        MsgBox "������ � ������� �� ������ numOfColParams"
        Exit Function
    End If
                                  


    ' ���������� ������� � ������� ����� � ��������
    Set rowHdrCol = New Collection
    Set colHdrCol = New Collection
    For rowNum = 1 To height Step numOfRowParams ' ����� ������ � �������� ���������
        ReDim tmpArr(0 To numOfRowProperties - 1)
        tmpArr = MatrixPart(rowLabels, rowNum, rowNum, 1, numOfRowProperties, True, False)
        rowHdrCol.Add tmpArr, CStr(rowNum) ' ����� � ���������
    Next rowNum
        
    For colNum = 1 To width Step numOfColParams ' ����� ������� � �������� ���������
        ReDim tmpArr(0 To numOfRowProperties - 1)
        tmpArr = MatrixPart(colLabels, 1, numOfColProperties, colNum, colNum, True, False)
        colHdrCol.Add tmpArr, CStr(colNum) ' ����� � ���������
    Next colNum
        
    ' ������������ ������ ������ � ���������
    numOfRecords = 0
    Set dataCol = New Collection
    For rowNum = 1 To height Step numOfRowParams ' ����� ������ � �������� ���������
        
        For colNum = 1 To width Step numOfColParams ' ����� ������� � �������� ���������
            ' ������ ��������� � ����� ������� ������
            
            ' 2-D ������ �� ��������� ������
            cellArr = MatrixPart(sourceData, rowNum, rowNum + numOfRowParams - 1, colNum, colNum + numOfColParams - 1, , False)
            cellFlatArr = getFlatArray(cellArr, 0) ' 1-d ������ ��������
            arrIsEmpty = True
            For i = LBound(cellFlatArr) To UBound(cellFlatArr)
                arrIsEmpty = arrIsEmpty And IsEmpty(cellFlatArr(i))
            Next i
            
            If (getNonEmptyOnly And Not arrIsEmpty) Or Not getNonEmptyOnly Then ' ���� � ������� ���-�� ����
                rowFlatArr = rowHdrCol.Item(CStr(rowNum))
                colFlatArr = colHdrCol.Item(CStr(colNum))
                
                tmpArr = mergeVectors(rowFlatArr, colFlatArr, cellFlatArr)
                
                numOfRecords = numOfRecords + 1
                'dataCol.Add Array(CStr(numOfRecords), tmpArr), CStr(numOfRecords)
                dataCol.Add tmpArr, CStr(numOfRecords)
            End If
            
        Next colNum
    Next rowNum
    
    Set ReadForPivotCol = dataCol
    
End Function