''' «ахватывает максиамльно возможный сплошной диапазон €чеек относительно заданной €чейки
''' direction - задаЄт направление движени€
''' ≈сли задана €чейка на границе листа, за границу листа не идЄт - возвращает заданную €чейку
''' ≈сли задана пуста€ €чейка - возввращает заданную €чейку
Function enlargeRange(rng As Range, Optional direction As XlDirection = xlDown) As Range
    Dim resultRange As Range
    
    If rng.Cells(1, 1).value = "" Then ' текуща€ €чейка пуста€
        ' pass - плохо
        Set resultRange = rng.Cells(1, 1) ' текущую €чейку и возвращаем
    Else
        Select Case direction
        Case xlDown
            If rng.Cells(1, 1).Row = rng.Parent.Cells.Rows.Count Then ' €чейка в последней строке
                Set resultRange = rng.Cells(1, 1)
            ElseIf rng.Cells(1, 1).Offset(1, 0).value <> "" Then ' не последн€€ в сплошном диапазоне
                Set resultRange = Range(rng.Cells(1, 1), rng.Cells(1, 1).End(xlDown))
            Else ' если текуща€ €чейка последн€€ в сплошном диапазоне
                Set resultRange = rng.Cells(1, 1)
            End If
            
        Case xlUp
            If rng.Cells(1, 1).Row = 1 Then ' €чейка в первой строке
                Set resultRange = rng.Cells(1, 1)
            ElseIf rng.Cells(1, 1).Offset(-1, 0).value <> "" Then ' текуща€ €чейка не последн€€
                Set resultRange = Range(rng.Cells(1, 1), rng.Cells(1, 1).End(xlUp))
            Else ' если текуща€ €чейка последн€€
                Set resultRange = rng.Cells(1, 1)
            End If
        
        Case xlToLeft
            If rng.Cells(1, 1).Column = 1 Then ' €чейка в первой колонке
                Set resultRange = rng.Cells(1, 1)
            ElseIf rng.Cells(1, 1).Offset(0, -1).value <> "" Then ' не последн€€ в сплошном диапазоне
                Set resultRange = Range(rng.Cells(1, 1), rng.Cells(1, 1).End(xlToLeft))
            Else ' последн€€ в сплошном диапазоне
                Set resultRange = rng.Cells(1, 1)
            End If
            
        Case xlToRight
            If rng.Cells(1, 1).Column = rng.Parent.Cells.Columns.Count Then ' €чейка в последнем столбце
                Set resultRange = rng.Cells(1, 1)
            ElseIf rng.Cells(1, 1).Offset(0, 1).value <> "" Then ' не последн€€ в сплошном диапазоне
                Set resultRange = Range(rng.Cells(1, 1), rng.Cells(1, 1).End(xlToRight))
            Else ' последн€€ в сплошном диапазоне
                Set resultRange = rng.Cells(1, 1)
            End If
        Case Else
            Set resultRange = rng.Cells(1, 1)
            
        End Select

    End If
    
    Set enlargeRange = resultRange
    
End Function