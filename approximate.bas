Type point ' пользовательский тип - точка
    x As Single
    y As Single
End Type

' интерполирует внутри точек, экстраполирует за пределами точек, координаты точек заданы строкой
Function approximate(x0 As Single, Optional points As String = "") As Variant
    Dim pointsArray() As point, tmpPoint As point, pointsArrayTmp As Variant
    Dim x1 As Single, x2 As Single
    Dim y1 As Single, y2 As Single
    Dim tmp As Variant
    Dim i As Integer, j As Integer
    Dim minX As Single, maxX As Single
    Dim numOfPoints As Integer
    
    If points = "" Then ' пользователь самостоятельно может задавать множество точек, не ковыряя функцию
        ' в качестве примера дана функция x^2 на отрезке x=[0, 4]
        points = "(0,0);(0.5,0.25);(1,1);(1.5,2.25);(2,4);(2.5,6.25);(3,9);(4,16)" ' массив точке в формате "(x1,y1);(x2,y2)"
    End If
    
    pointsArrayTmp = Split(points, ";")
    numOfPoints = UBound(pointsArrayTmp) - LBound(pointsArrayTmp) + 1 ' количество точек в массиве
    If numOfPoints < 2 Then ' чтобы по одной точке прогноз никто не строил
        approximate = "#Error!"
        Exit Function
    End If
    ReDim pointsArray(LBound(pointsArrayTmp) To UBound(pointsArrayTmp))
    
    minX = 3.4E+38 ' максимальное и минимальное значения (знаки не трогать, здесь всё верно)
    maxX = -3.4E+38
    
    For i = LBound(pointsArrayTmp) To UBound(pointsArrayTmp)
        tmp = Mid(pointsArrayTmp(i), 2, Len(pointsArrayTmp(i)) - 2) ' берём всё кроме скобок
        tmp = Split(tmp, ",", 2) ' бьём строку по запятым, но не более чем на 2 части
         
         ' контроль здесь не делаю
        pointsArray(i).x = CSng(Trim(tmp(0)))  ' преобразуем из текста в дробные числа
        pointsArray(i).y = CSng(Trim(tmp(1)))
        
        ' ищем минимум/максимум, чтобы сразу знать, когда экстраполировать
        ' последовательно уменьшаем минимум и увеличиваем максимум
        If pointsArray(i).x > maxX Then maxX = pointsArray(i).x
        If pointsArray(i).x < minX Then minX = pointsArray(i).x
    Next i
    ' в этой точке у нас есть весь массив точек
    
    ' сортируем пузырьком
    For i = LBound(pointsArray) To UBound(pointsArray) - 1
        For j = i + 1 To UBound(pointsArray)
            If pointsArray(i).x > pointsArray(j).x Then ' если слева бОльший элемент - делаем обмен
                tmpPoint = pointsArray(i)
                pointsArray(i) = pointsArray(j)
                pointsArray(j) = tmpPoint
            End If
        Next j
    Next i
    ' у нас есть массив точек, сортированный по X
    
    
    ' есть три участка:
    ' 1 - "до" кривой (экстраполяция по первым двум точкам)
    ' 2 - кривая (интерполяция)
    ' 3 - "после" кривой (экстраполяция по последним двум точкам)
    
    If x0 < minX Then ' до кривой
        x1 = pointsArray(0).x
        x2 = pointsArray(1).x
        y1 = pointsArray(0).y
        y2 = pointsArray(1).y
        
    ElseIf minX <= x0 And x0 <= maxX Then ' кривая
        For i = LBound(pointsArray) To UBound(pointsArray) - 1
            If pointsArray(i).x <= x0 And x0 <= pointsArray(i + 1).x Then
                x1 = pointsArray(i).x
                x2 = pointsArray(i + 1).x
                y1 = pointsArray(i).y
                y2 = pointsArray(i + 1).y
                Exit For
            End If
        Next i
        
    ElseIf maxX < x0 Then ' после кривой
        x1 = pointsArray(numOfPoints - 2).x
        x2 = pointsArray(numOfPoints - 1).x
        y1 = pointsArray(numOfPoints - 2).y
        y2 = pointsArray(numOfPoints - 1).y
    
    End If
    
    ' задача интерполяции
    ' y = b + kx
    k = (y1 - y2) / (x1 - x2) ' то же что и тангенс угла = противолежащий катет к прилежащему
    b = ((y1 * x2) - (y2 * x1)) / (x2 - x1)
    approximate = b + k * x0 ' значение функции в точке x0

End Function

