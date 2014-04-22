Sub reorderColumns(tableRng As Range, columnsOrder As Collection)
    ' ������� �������� �������� � ���������
    ' ��������� ������ � �������� �������� ������
    ' � ������ ����� ������� ��������, �� ��������� - ������������ ����� + 1
    ' ��������� �������
    ' ������� ������
    
    Dim sh As Worksheet
    Dim fullRng As Range
    Dim hdrRng As Range
    Dim orderRng As Range
    Dim itemVal As Variant
    Dim maxValue As Single
    Dim item As Variant
    Dim cl As Range
    Dim key As String
    
    Set sh = tableRng.Parent
    
    Set hdrRng = Intersect(sh.Rows(tableRng.row), tableRng)
    sh.Rows(hdrRng.row).Insert Shift:=xlDown ' �������� ������
    Set orderRng = hdrRng.Offset(-1, 0)
    Set fullRng = tableRng.Offset(-1, 0).Resize(tableRng.Rows.count + 1, tableRng.columns.count) ' �������� � ����������� �������
    
    ' ����� ��������
    ' ����� ������������ ��������
    maxValue = -1
    For Each item In columnsOrder ' ���� �� ���������
        maxValue = WorksheetFunction.Max(maxValue, CSng(item))
    Next item
    ' ���������� ����� � ������ ��������, �� ������� ����� �����������
    For Each cl In orderRng
        key = cl.Offset(1, 0).value
        If isInCollection(key, columnsOrder) Then
            itemVal = columnsOrder.item(key)
        Else
            itemVal = maxValue + 1
        End If
        cl.value = itemVal
    Next cl
        
    sh.Sort.SortFields.Clear
    sh.Sort.SortFields.Add key:=orderRng, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With sh.Sort
        .SetRange fullRng ' ��������� �� ������� + 1 ������
        .Header = xlNo
        .matchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    Rows(hdrRng.row - 1).Delete Shift:=xlUp ' ������� ������
End Sub