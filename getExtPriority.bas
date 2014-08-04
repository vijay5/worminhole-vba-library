''' ���������� ����� ����������, ���� ����������� �������, � ���������� max+1 ����� ����������, ���� �� ����������� �� ���� �������
''' ������� �������� � �������� ������� conditionRng
''' ������ �� conditionRng �������� ������� � ������������ priorRng
''' testCondRng - ������ �� ����������

Public Function getExtPriority(testCondRng As Range, conditionRng As Range, priorRng As Range) As Long
    Dim chk As Boolean
    Dim chk1 As Boolean
    Dim conditionArr As Variant
    Dim priorArr As Variant
    Dim testCondArr As Variant
    Dim prior As Single
    Dim allAreBlank As Boolean
    
    Dim curRow As Long
    
    ' �������� �������
    chk = True
    chk = chk And (conditionRng.Columns.Count = testCondRng.Columns.Count)
    chk = chk And (testCondRng.Rows.Count = 1)
    chk = chk And (priorRng.Columns.Count = 1)
    chk = chk And (priorRng.Rows.Count = conditionRng.Rows.Count)
    If Not chk Then
        getExtPriority = "#N/A"
        Exit Function
    End If
    
    conditionArr = conditionRng.value ' ������ c ��������� ��� �����������
    priorArr = priorRng.value         ' ������ � ������������
    testCondArr = testCondRng.value   ' ������ ��� ��������
    
    Dim rowNum As Long
    Dim colNum As Long
    
    prior = 0
    For rowNum = LBound(conditionArr, 1) To UBound(conditionArr, 1) ' ���� �� ���� �����������
        chk = True ' ������� ��� ������ ������
        allAreBlank = True
        For colNum = LBound(testCondArr, 2) To UBound(testCondArr, 2) ' ���� �� ��������
            allAreBlank = allAreBlank And (conditionArr(rowNum, colNum) = "")
            chk1 = IIf(conditionArr(rowNum, colNum) = "", True, conditionArr(rowNum, colNum) = testCondArr(1, colNum))
            chk = chk And chk1
        Next colNum
        
        If chk And Not allAreBlank Then ' ����� ������, ������� ���������
            prior = priorArr(rowNum, 1) ' ���������� ���������
            Exit For
        End If
    Next rowNum
    If prior = 0 Then ' �� ��������� �� ���� ������� = ������������ ��������� + 1
        prior = WorksheetFunction.MAX(priorArr) + 1
    End If
    getExtPriority = prior
End Function