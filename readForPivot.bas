Sub ReadForPivot(sourceRange As Range, numOfRowProperties As Integer, numOfColProperties As Integer, Optional numOfRowParams As Integer = 1, Optional numOfColParams As Integer = 1)
    ''' ������ ������� ���� � ������� �������
    ''' �� ����� ����� ���� ������ ������� �������
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
    ''' ����� �����, ��� ���-�� ����� ��� �������� ����� ���-�� "������������ ����������" (���� ������� ���������� �����)
    '''
    ''' �� ������ ����� ��������� "�������" �������
    '''  Col1  Col2  ColN ||Var1  Var2 || Value
    ''' +-----+-----+-----++-----+-----++-------+
    ''' |     |     |     ||     |     ||       |

    Dim height0 As Long, height As Long, width As Long
    Dim rowLabels As Variant, colLabels As Variant, sourceData As Variant
    Dim minRow As Long, maxRow As Long
    Dim lastFixedCol As Long, lastVarCol As Long, dataCol As Long
    Dim sourceSheet As Variant
    
    ' �������� �� ��������� ���������
    Set sourceSheet = sourceRange.Parent
    height0 = sourceRange.Rows.Count
    width = sourceRange.Columns.Count
    
    If height0 <= numOfColProperties Or width <= numOfRowProperties Then
       Call MsgBox("�������� ������� ��������� �������")
       Exit Sub
    End If
    
    ' �������� �������� � ������
    rowLabels = sourceRange.Cells(1 + numOfColProperties, 1).Resize(height0 - numOfColProperties, numOfRowProperties).value
    colLabels = sourceRange.Cells(1, 1 + numOfRowProperties).Resize(numOfColProperties, width - numOfRowProperties).value
    
    ' �������� � ��������� �������
    Set sourceRange = sourceRange.Resize(height0 - numOfColProperties, width - numOfRowProperties). _
                                  Offset(numOfColProperties, numOfRowProperties)
    sourceData = sourceRange.value
    height0 = sourceRange.Rows.Count  ' ����� ����� � �������� �������
    height = height0 \ numOfRowParams ' ����� ����� �� ���� ������ ���������� �������
    width = sourceRange.Columns.Count
                                  
    ' ������� �� ����� �����
    Set resultSheet = Sheets.Add(after:=ActiveSheet)
    ' ���������
    originRow = 1
    originCol = 1
    lastFixedCol = originCol + numOfRowProperties - 1 ' ����� ���������� ������� � �������������� �������
    lastVarCol = lastFixedCol + numOfColProperties    ' ����� ���������� ������� � ����������� ������� (����� �� colLabels)
    dataCol = lastVarCol + 1                          ' ����� �������, ������� � �������� ������ ��������� ������
    
    ' ������� ���� �������� �������� �������, ������ ������� ������� ����
    ' ������������, ��� ��� ������� ��������� ����� ������� ������ ����� ������ � ������� ���� � �� �� (�������� �� ������)
    colBlock = 0
    For col = 1 To UBound(sourceData, 2) Step numOfColParams
        colBlock = colBlock + 1
        minRow = originRow + (colBlock - 1) * height
        maxRow = originRow + colBlock * height - 1
        
        ' ������� ������������ ����� �����
        Range(Cells(minRow, originCol), Cells(maxRow, lastFixedCol)) = stepArray(rowLabels, 1, numOfRowParams, 1, 1) ' ��� ����� ������ ����� �����������
        For j = 1 To numOfColProperties     ' ������� ���������� �����
            Range(Cells(minRow, lastFixedCol + j), Cells(maxRow, lastFixedCol + j)).value = colLabels(j, col)
        Next j
        ' ������� �������� �����
        For colParam = 1 To numOfColParams
            outArray = MatrixPart(sourceData, 1, height0, col + colParam - 1, col + colParam - 1, False, False) ' ����������� ���� ������� �� ��������� �������
            For rowParam = 1 To numOfRowParams
                colNum = dataCol + (colParam - 1) * numOfRowParams + rowParam - 1
                Range(Cells(minRow, colNum), Cells(maxRow, colNum)) = stepArray(outArray, rowParam, numOfRowParams, 1, 1)
            Next rowParam
        Next colParam
    Next col
    
    ' ��������� ������� � ����� ��������
    Cells(1, 1).EntireRow.Insert Shift:=xlShiftDown
    
    ' ������� �� ����������, ������� ���� ����� �� �������
    For i = 1 To numOfRowProperties
        With Cells(1, originCol + i - 1)
            tmp = sourceRange.Cells(1, 1).Offset(-numOfColProperties, -numOfRowProperties + i - 1).value
            If tmp <> "" Then ' ���� ���� �������� - �� ����������� ���� ���
                .value = tmp
            Else
                .value = "RowAttribute_" & CStr(i)
            End If
            .Interior.Color = RGB(255, 255, 0)
        End With
    Next i
    
    ' ������� �� ����������, ������� ���� ��� ��������
    For i = 1 To numOfColProperties
        With Cells(1, lastFixedCol + i)
            .value = "ColumnAttribute_" & CStr(i)
            .Interior.Color = RGB(196, 215, 155)
        End With
    Next i
    
    ' ������� �� ��������� �����
    If numOfColParams = 1 And numOfRowParams = 1 Then
        With Cells(1, lastVarCol + 1)
            .value = "Value"
            .Interior.Color = RGB(255, 0, 255)
        End With
    Else
        For i = 1 To numOfColParams
            For j = 1 To numOfRowParams
                With Cells(1, lastVarCol + (i - 1) * numOfRowParams + j)
                    .value = "Value_Col" & CStr(i) & "_Row" & CStr(j)
                    .Interior.Color = RGB(255, 0, 255)
                End With
            Next j
        Next i
    End If
    
    Cells(1, 1).Select
    Cells(2, 1).EntireRow.Select
    ActiveWindow.FreezePanes = True
    
End Sub