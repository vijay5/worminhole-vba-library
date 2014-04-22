' преобразует пиксели в условные единицы ширины/высоты €чейки
Function pix2Point(pixels As Integer, Optional forColumn As Boolean = True) As Single
    If pixels >= 1 Then
        If forColumn Then  ' перевод дл€ столбцов
            If pixels >= 12 Then
                pix2Point = Round((pixels - 5) / 7, 2)
            Else
                pix2Point = Round((pixels) / 12, 2)
            End If
        Else               ' перевод дл€ строк
            pix2Point = Round(pixels * 0.75, 2)
        End If
    Else
        pix2Point = 0 ' спр€чет колонку / строку
    End If
End Function