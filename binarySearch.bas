' Ищет по коллекции элемент с ближайшим кодом, возвращает номер элемента коллекции
' (перед/после) которого нужно вставить искомый элемент
' Предполагается, что в коллекции лежат одномерные массивы,
' а значение ищется внутри заданного столбца одномерного массива в колонке useColumn
' Возвращает Array(-1/0/1, elemNum), где -1 - вставить "до", 0 - просто вставить, +1 - вставить "после",
' elemNum - номер элемента
' Время на поиск почти линейно зависит от числа элементов (проверял на 50000 элементов типа Long)
' для 10000: t=0.789 мс на поиск 1 элемента
' для 25000: t=2.069 мс -/-
' для 50000: t=4.533 мс -/-
' Время на добавление элемента линейно, -/-
' для 10000: t=0.046 мс на вставку 1 элемента
' для 25000: t=0.090 мс -/-
' для 50000: t=0.176 мс -/-
Public Function BinarySearch(arr As Collection, valueToFind As Variant, useColumn As Integer) As Variant
    Dim minRow As Long, maxRow As Long, midRow As Long
    Dim globalMin As Long, globalMax As Long
    
    globalMin = 1
    globalMax = arr.Count
    minRow = globalMin
    maxRow = globalMax

    If globalMax = 0 Then ' на случай елси в коллекции пусто
        BinarySearch = Array(0, 0) ' просто вставить
        Exit Function
    End If
    
    Do
        midRow = (minRow + maxRow) \ 2 ' среднее значение
        If valueToFind < arr(midRow)(useColumn) Then
            maxRow = midRow - 1
        ElseIf valueToFind > arr(midRow)(useColumn) Then
            minRow = midRow + 1
        Else
            BinarySearch = Array(1, midRow)
            Exit Do
        End If
        
        If (minRow > maxRow) Then ' точного элемента не найдено
            ' внутри массива
            If minRow <= globalMax And maxRow >= globalMin Then
                If valueToFind > arr(maxRow)(useColumn) Then
                    BinarySearch = Array(1, maxRow)
                ElseIf valueToFind < arr(minRow)(useColumn) Then
                    BinarySearch = Array(-1, minRow)
                    Stop ' эта ветвь никогда не срабатывала
                Else
                    Stop ' эта ветвь никогда не срабатывала
                    BinarySearch = Array(1, midRow)
                End If
            Else ' за границами массива
                If maxRow < globalMin Then
                    BinarySearch = Array(-1, minRow)
                ElseIf minRow > globalMax Then
                    BinarySearch = Array(1, maxRow)
                End If
            End If
            Exit Do
        End If
    Loop
End Function