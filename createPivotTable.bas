''' ������ ������� ������� (���������) �� ��������� ��������� ��� pivotCache'� �� �������/�������� ����� �����
'REQUIRES: MakeRandomName
Function createPivotTable(Optional ByVal destRange As Range, Optional ByVal sourceDataRange As Range, Optional ptCache As PivotCache) As Object
    Dim textAddr As String
    Dim ptTable As PivotTable
    Dim pivotTableName As Variant
    Dim tmp As Variant
    Dim sel As String
    
    
    If Not (ptCache Is Nothing) Then
        ' pass
    Else
        ' '[������� �� 18-06.xlsx]Sheet1'!$A$1:$C$25866
        textAddr = "'[" & sourceDataRange.Parent.Parent.Name & "]" & sourceDataRange.Parent.Name & "'!" & sourceDataRange.Address(, , xlR1C1)
        Set ptCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, textAddr, xlPivotTableVersion14) ' ������� ���
    End If
    
    pivotTableName = MakeRandomName(10) ' ��� ������� ������� - ����� ���������
    If destRange Is Nothing Then
        ' ������ Pivot �� ����� �����
        Set ptTable = ptCache.createPivotTable(TableDestination:="", DefaultVersion:=ptCache.Version, tableName:=pivotTableName) ' ������� ���� � Pivot'��
        Set destRange = ActiveSheet.Cells(3, 1) ' ������ ������� � ������ A3 (������ ����� � ������ ������)
        ActiveSheet.PivotTableWizard TableDestination:=destRange.Cells(1, 1)
    Else
        destRange.Parent.Select ' ������� ����
        ' ������ Pivot �� ������������ �����
        Set ptTable = ptCache.createPivotTable(TableDestination:=destRange, DefaultVersion:=ptCache.Version, tableName:=pivotTableName) ' ������� ���� � Pivot'��
    End If
    
    ' ����� ��� ������ ����
    destRange.Cells(1, 1).Select ' ������ �� ����� ������� ������ ����������� ���������
    
    ptTable.ColumnGrand = False
    ptTable.RowGrand = False
    ptTable.HasAutoFormat = False ' ����� ��� ���������� �������� ������ ��������
    
    Set createPivotTable = ptTable ' ���������� ������ �� Pivot
    
End Function
