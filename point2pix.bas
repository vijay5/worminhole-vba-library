' преобразует условные единицы ширины/высоты €чейки в пиксели
Function point2pix(points As Single, Optional forColumn As Boolean = True) As Integer
    
    If points > 0 Then
        If forColumn Then  ' перевод дл€ столбцов
            If points >= 1 Then
                point2pix = Round(points * 7 + 5, 0) '(pixels - 5) / 7
            Else
                point2pix = Round(points * 12, 0)
            End If
        Else
            point2pix = Round(points / 0.75, 0)
        End If
    Else
        point2pix = 0
    End If
    
End Function