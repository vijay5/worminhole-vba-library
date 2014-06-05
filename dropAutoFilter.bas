''' сброс автофильтра
Sub dropAutoFilter()
    Dim sh As Worksheet
    Dim cnt As Integer
    Dim flt As Object
    
    Set sh = ActiveSheet
    
    If sh.AutoFilter Is Nothing Then ' автофильтра нет
    Else ' автофильтр есть
        cnt = 0
        For Each flt In sh.AutoFilter.Filters
            cnt = cnt + 1
            If flt.On Then
                sh.AutoFilter.Range.AutoFilter Field:=cnt ' сбрасываем фильтр для конкретной колонки
            End If
        Next flt
    End If
End Sub
