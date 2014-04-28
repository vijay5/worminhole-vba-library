''' ������ ������� ������� (���������) �� ��������� ��������� ��� pivotCache'� �� �������/�������� ����� �����
Function createPivotTable(Optional ByVal destRange As Range, Optional ByVal sourceDataRange As Range, Optional ptCache As PivotCache) As Object
    Dim ptTable As PivotTable
    Dim pivotTableName As Variant
    Dim tmp As Variant
    Dim sel As String
    
    
    If Not (ptCache Is Nothing) Then
        ' pass
    Else
        Set ptCache = ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, sourceData:=sourceDataRange.Address(, , xlR1C1)) ' ������� ���
    End If
    
    pivotTableName = MakeRandomName(10) ' ��� ������� ������� - ����� ���������
    If destRange Is Nothing Then
        ' ������ Pivot �� ����� �����
        Set ptTable = ptCache.createPivotTable(TableDestination:="", DefaultVersion:=xlPivotTableVersion10, TableName:=pivotTableName) ' ������� ���� � Pivot'��
        Set destRange = ActiveSheet.Cells(3, 1) ' ������ ������� � ������ A3 (������ ����� � ������ ������)
        ActiveSheet.PivotTableWizard TableDestination:=destRange.Cells(1, 1)
    Else
        destRange.Parent.Select ' ������� ����
        ' ������ Pivot �� ������������ �����
        Set ptTable = ptCache.createPivotTable(TableDestination:=destRange, DefaultVersion:=xlPivotTableVersion10, TableName:=pivotTableName) ' ������� ���� � Pivot'��
    End If
    
    ' ����� ��� ������ ����
    destRange.Cells(1, 1).Select ' ������ �� ����� ������� ������ ����������� ���������
    
    ptTable.ColumnGrand = False
    ptTable.RowGrand = False
    ptTable.HasAutoFormat = False ' ����� ��� ���������� �������� ������ ��������
    
    Set createPivotTable = ptTable ' ���������� ������ �� Pivot
    
End Function