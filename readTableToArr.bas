' ������ ��������� ������� � �������, ��������� �� ������������ �����, ����� ������� reBase
' ������ �������:     readTableToArr Array("WARE", "PO_STATUS", "PO_DATE_DELIVERY", "PLAN_AMOUNT_DELIVERY"), 
'                                    Workbooks("AW15.XLS").Sheets(1).Range("1:1"), outArr, Array("WARE")

'REQUIRES: getMaxRow, arrayLength, getProperValArray, FindCell, addUniqToCol, regroupArray, col2Array, array2col, isInCollection, addToText
Sub readTableToArr(fieldsList As Variant, headerRng As Range, outArr As Variant, Optional keyFieldsList As Variant = False, Optional minRow As Long = 0, Optional minCol As Long = 0)
    
    Dim sh As Worksheet
    Dim maxRow As Long
    Dim colCount As Long
    Dim columnsData() As Variant
    Dim fieldsListRebased As Variant
    Dim hdrColumn As Range
    Dim dtRange As Range
    Dim tmpRng As Range
    Dim rowNum As Long
    Dim colNum As Long
    Dim tmpCol As Collection
    Dim keyFieldsCol As Collection
    Dim item As Variant
    Dim key As String
    Dim curColName As String
    
    
    
    Set sh = headerRng.Parent
    maxRow = getMaxRow(sh)
    colCount = arrayLength(fieldsList)
    
    ReDim columnsData(0 To colCount - 1) ' ������ �� ���������� 1D * 2D
    ' ��� ������ ��� ������ �������
    Set dtRange = Range(sh.Cells(headerRng.Cells(1, 1).Row + headerRng.Rows.Count, 1), sh.Cells(maxRow, 1)).EntireRow
    
    For i = 0 To colCount - 1
        Set hdrColumn = FindCell(CStr(fieldsList(i)), headerRng, , xlWhole).EntireColumn ' ���� �������� ����
        If Not hdrColumn Is Nothing Then
            columnsData(i) = getProperValArray(Intersect(hdrColumn, dtRange))
        Else
            columnsData(i) = -1
        End If
    Next i
    
    '
    allUniqueRecords = False
    severalUniqueRecords = False
    If IsArray(keyFieldsList) Then ' ����� ������ ���������� ������ (��� ������ ������� � ������ ���������� �����)
        severalUniqueRecords = True
        Set keyFieldsCol = array2col(keyFieldsList, True)
        
    ElseIf TypeName(keyFieldsList) = "Boolean" Then ' ����� ���� - True(��� ����)/False(�� ���� ����)
        allUniqueRecords = keyFieldsList
    Else ' ������ ���-�� ��� = False(�� ���� ����)
        ' pass
    End If
    
    If allUniqueRecords Or severalUniqueRecords Then ' ��������� �� ������������ �� ���� ��� ��������� ��������
        Set tmpCol = New Collection
        
        For rowNum = 0 To dtRange.Rows.Count - 1
            ReDim item(0 To colCount - 1) ' ������ ������ ����� �������
            key = ""
            For colNum = 0 To colCount - 1
                item(colNum) = columnsData(colNum)(rowNum + 1, 1) ' ��������� ������
                
                If allUniqueRecords Then ' ���� ���� ��� �������
                    key = addToText(key, CStr(item(colNum)), Chr(1))
                Else ' �� ��� �������
                    curColName = CStr(fieldsList(colNum))
                    If isInCollection(curColName, keyFieldsCol) Then ' ��������� ��������� �������� ������� � ��������� ��������
                        key = addToText(key, CStr(item(colNum)), Chr(1))
                    Else
                        ' pass
                    End If
                End If
            Next colNum
            addUniqToCol tmpCol, item, key ' ��������� � ���������
        Next rowNum
        
        outArr = regroupArray(col2Array(tmpCol), True, minRow, minCol)

    Else ' �� ��������� �� ������������
    
        ReDim outArr(0 + minRow To dtRange.Rows.Count - 1 + minRow, 0 + minCol To colCount - 1 + minCol)
        
        For rowNum = 0 To dtRange.Rows.Count - 1
            For colNum = 0 To colCount - 1
                outArr(rowNum + minRow, colNumminCol + minCol) = columnsData(colNum)(rowNum + 1, 1)
            Next colNum
        Next rowNum
    End If
    
End Sub
