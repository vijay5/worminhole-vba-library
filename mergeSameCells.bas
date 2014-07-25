' группирует €чейки с одинаковыми значени€ми или цветами заливки
' selRng   - вектор-строка или вектор-столбец
' useColor - True - группируем по цвету заливки, False - группируем по значени€м €чеек
Sub mergeSameCells(selRng As Range, Optional useColor As Boolean = False)
    Dim begCl As Range
    Dim endCl As Range
    Dim prevEndCl As Range
    Dim cl as Range
    Dim cellsTotalCnt as Long
    Dim cellsCnt as Long
    Dim chk as Boolean
    
    If Not (selRng.Rows.Count = 1 Or selRng.Columns.Count = 1) Then ' нельз€ группировать большие диапазоны (можно, но сложно)
        Exit Sub
    End If
    
    Set begCl = selRng.Cells(1, 1)
    Set endCl = Nothing
    
    cellsTotalCnt = selRng.Cells.Count
    cellsCnt = 0
    
    For Each cl In selRng
        cellsCnt = cellsCnt + 1 ' считаем количество пройденных €чеек
        
        Set prevEndCl = endCl   ' предыдуща€ €чейка
        Set endCl = cl          ' текуща€ €чейка
        
        ' выбираем условие, по которому группируем
        If useColor Then ' группируем по цвету
            chk = (begCl.Interior.Color <> endCl.Interior.Color)
        Else ' группируем по значению
            chk = (begCl.Value <> endCl.Value)
        End If
        
        If chk Then ' начальна€ €чейка диапазона отличаетс€ от текущей - диапазон закончилс€ на предыдущей €чейке
            If Range(begCl, prevEndCl).Cells.Count > 1 Then
                tmp = Application.DisplayAlerts
                Application.DisplayAlerts = False
                Range(begCl, prevEndCl).Merge
                Application.DisplayAlerts = tmp
            End If
            Range(begCl, prevEndCl).VerticalAlignment = xlCenter
            Range(begCl, prevEndCl).HorizontalAlignment = xlCenter
            
            Set begCl = cl
        Else
            ' pass
        End If

        If cellsCnt = cellsTotalCnt Then ' мы находимс€ на последней €чейке - группируем
            If Range(begCl, endCl).Cells.Count > 1 Then
                tmp = Application.DisplayAlerts
                Application.DisplayAlerts = False
                Range(begCl, endCl).Merge
                Application.DisplayAlerts = tmp
            End If
            Range(begCl, endCl).VerticalAlignment = xlCenter
            Range(begCl, endCl).HorizontalAlignment = xlCenter
        Else
            ' pass
        End If
    Next cl
End Sub
