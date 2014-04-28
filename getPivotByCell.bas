''' по адресу на листе возвращает сводную таблицу
Function getPivotByCell(rng As Range) As PivotTable
    Dim sh As Worksheet
    Dim pivot As Variant
    
    
    Set getPivotByCell = Nothing ' значение по умолчанию
    Set sh = rng.Parent ' лист, на котором указан адрес
    
    For Each pivot In sh.PivotTables
        If Not Intersect(rng.Cells(1, 1), pivot.TableRange2) Is Nothing Then ' если есть пересечение
            Set getPivotByCell = pivot
            Exit For
        Else
            ' идём дальше
        End If
    Next pivot
End Function