''' создаёт разреженный вектор значений. Числа с шагом step, ограничение сверху и/или по макс. числу элементов в векторе
Function stepFunction(startAt As Variant, step As Variant, Optional stopAt As Variant = "", Optional numOfPoints As Integer = -1) As Variant
    Dim curValue As Variant, curPoints As Integer
    Dim chk As Boolean
    Dim listOfValues As Variant
    
    If (stopAt = "" And numOfPoints = -1) Or step = 0 Then
        stepFunction = "!Error"
    End If
    
    curValue = startAt
    curPoints = 1
    
    ' сложное условие
        If stopAt <> "" And numOfPoints <> -1 Then         ' заданы оба параметра
        chk = -Sgn(step) * (curValue - stopAt) >= 0 And curPoints <= numOfPoints
    ElseIf Not stopAt <> "" And numOfPoints <> -1 Then     ' задано количество точек
        chk = curPoints <= numOfPoints
    ElseIf stopAt <> "" And Not numOfPoints <> -1 Then     ' задано предельное значение
        chk = -Sgn(step) * (curValue - stopAt) >= 0
    ElseIf Not stopAt <> "" And Not numOfPoints <> -1 Then ' не задан ни один параметр
        ' не должен дойти до сюда
    End If
        
    Do While chk
        If curPoints = 1 Then ReDim listOfValues(1 To curPoints) Else ReDim Preserve listOfValues(1 To curPoints)
        listOfValues(curPoints) = curValue
        
        curValue = curValue + step
        curPoints = curPoints + 1
        
        ' сложное условие
            If stopAt <> "" And numOfPoints <> -1 Then         ' заданы оба параметра
            chk = -Sgn(step) * (curValue - stopAt) >= 0 And curPoints <= numOfPoints
        ElseIf Not stopAt <> "" And numOfPoints <> -1 Then     ' задано количество точек
            chk = curPoints <= numOfPoints
        ElseIf stopAt <> "" And Not numOfPoints <> -1 Then     ' задано предельное значение
            chk = -Sgn(step) * (curValue - stopAt) >= 0
        ElseIf Not stopAt <> "" And Not numOfPoints <> -1 Then ' не задан ни один параметр
            ' не должен дойти до сюда
        End If
    Loop
    stepFunction = listOfValues
End Function