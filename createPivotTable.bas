''' создаЄт сводную таблицу (заготовку) по заданному диапазону или pivotCache'у на текущем/заданном листе листе
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
        ' '[ќстатки на 18-06.xlsx]Sheet1'!$A$1:$C$25866
        textAddr = "'[" & sourceDataRange.Parent.Parent.Name & "]" & sourceDataRange.Parent.Name & "'!" & sourceDataRange.Address(, , xlR1C1)
        Set ptCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, textAddr, xlPivotTableVersion14) ' создали кэш
    End If
    
    pivotTableName = MakeRandomName(10) ' им€ сводной таблицы - почти случайное
    If destRange Is Nothing Then
        ' создаЄм Pivot на новом листе
        Set ptTable = ptCache.createPivotTable(TableDestination:="", DefaultVersion:=ptCache.Version, tableName:=pivotTableName) ' создали лист с Pivot'ом
        Set destRange = ActiveSheet.Cells(3, 1) ' создаЄм таблицу в €чейке A3 (фильтр будет в первой строке)
        ActiveSheet.PivotTableWizard TableDestination:=destRange.Cells(1, 1)
    Else
        destRange.Parent.Select ' выбрали лист
        ' создаЄм Pivot на существующем листе
        Set ptTable = ptCache.createPivotTable(TableDestination:=destRange, DefaultVersion:=ptCache.Version, tableName:=pivotTableName) ' создали лист с Pivot'ом
    End If
    
    ' здесь уже создан лист
    destRange.Cells(1, 1).Select ' встали на левую верхнюю €чейку выделенного диапазона
    
    ptTable.ColumnGrand = False
    ptTable.RowGrand = False
    ptTable.HasAutoFormat = False ' иначе при обновлении измен€ет ширину столбцов
    
    Set createPivotTable = ptTable ' возвращаем ссылку на Pivot
    
End Function
