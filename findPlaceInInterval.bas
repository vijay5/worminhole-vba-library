' когда на листе заданы два столбца (interval) + третий столбец со значениями
' с левыми и правыми границами диапазонов
' ищет строку с подходящим диапазоном и из этой строки возвращает значение в столбце
Function findPlaceInInterval(value As Variant, interval As Range) As Variant
    Dim r As Long
    findPlaceInInterval = -1 ' по умолчанию
    
    If interval.Columns.Count <> 3 Then Exit Function
    
    For r = 1 To interval.Rows.Count
        If interval.Cells(r, 2).value <= value And value <= interval.Cells(r, 3).value Then
            findPlaceInInterval = interval.Cells(r, 1).value
            Exit For
        End If
    Next r
End Function