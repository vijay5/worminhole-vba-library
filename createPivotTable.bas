''' создаЄт сводную таблицу (заготовку) по заданному диапазону или pivotCache'у на текущем/заданном листе листе
Function createPivotTable(Optional ByVal destRange As Range, Optional ByVal sourceDataRange As Range, Optional ptCache As PivotCache) As Object
    Dim ptTable As PivotTable
    Dim pivotTableName As Variant
    Dim tmp As Variant
    Dim sel As String
    
    
    If Not (ptCache Is Nothing) Then
        ' pass
    Else
        Set ptCache = ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, sourceData:=sourceDataRange.Address(, , xlR1C1)) ' создали кэш
    End If
    
    pivotTableName = MakeRandomName(10) ' им€ сводной таблицы - почти случайное
    If destRange Is Nothing Then
        ' создаЄм Pivot на новом листе
        Set ptTable = ptCache.createPivotTable(TableDestination:="", DefaultVersion:=xlPivotTableVersion10, TableName:=pivotTableName) ' создали лист с Pivot'ом
        Set destRange = ActiveSheet.Cells(3, 1) ' создаЄм таблицу в €чейке A3 (фильтр будет в первой строке)
        ActiveSheet.PivotTableWizard TableDestination:=destRange.Cells(1, 1)
    Else
        destRange.Parent.Select ' выбрали лист
        ' создаЄм Pivot на существующем листе
        Set ptTable = ptCache.createPivotTable(TableDestination:=destRange, DefaultVersion:=xlPivotTableVersion10, TableName:=pivotTableName) ' создали лист с Pivot'ом
    End If
    
    ' здесь уже создан лист
    destRange.Cells(1, 1).Select ' встали на левую верхнюю €чейку выделенного диапазона
    
    ptTable.ColumnGrand = False
    ptTable.RowGrand = False
    ptTable.HasAutoFormat = False ' иначе при обновлении измен€ет ширину столбцов
    
    Set createPivotTable = ptTable ' возвращаем ссылку на Pivot
    
End Function