' возвращает коллекцию, где ключом является один или несколько столбцов,
' а значениями - вектор-строка (1D)
' если keyList не задан - результат аналогичен array2col
' ключ должен быть уникальным для массива (остаётся только самое последнее записанное значение)
' REQUIRES: arrayDepth, isInCollection
Function array2IndexedCol(arr2D As Variant, Optional keyList As Variant = "", Optional mergeSymbol As String = "_") As Collection
    Dim outCol As New Collection
    Dim rowNum As Long
    Dim i As Long
    Dim chk As Boolean
    Dim keyArr As Variant
    Dim key As String
    Dim valArr As Variant
    
    Set array2IndexedCol = outCol
    
    If arrayDepth(arr2D) <> 2 Then
        MsgBox ("[array2IndexedCol] На вход необходимо подать 2D-массив")
        Exit Function
    End If
    
    If IsArray(keyList) Then
        chk = True
        For i = LBound(keyList) To UBound(keyList) ' перебор указанных индексов
            chk = chk And (keyList(i) >= LBound(arr2D, 2)) And (keyList(i) <= UBound(arr2D, 2))
        Next i
        If Not chk Then
            MsgBox ("[array2IndexedCol] Один из указанных индексов выходит за пределы массива")
            Exit Function
        End If
    End If

    For rowNum = LBound(arr2D, 1) To UBound(arr2D, 1)
        ' собираем ключ
        ReDim keyArr(LBound(keyList) To UBound(keyList))
        For i = LBound(keyList) To UBound(keyList) ' перебор указанных индексов
            keyArr(i) = CStr(arr2D(rowNum, keyList(i)))
        Next i
        key = Join(keyArr, mergeSymbol)
        
        ' собираем значение (1D-массив)
        ReDim valArr(LBound(arr2D, 2) To UBound(arr2D, 2))
        For i = LBound(arr2D, 2) To UBound(arr2D, 2)
            valArr(i) = arr2D(rowNum, i)
        Next i
        
        ' удаляем значение из коллекции, если там уже есть такой ключ (ключ должен быть уникальным)
        If isInCollection(key, outCol) Then
            outCol.Remove key
        End If
        
        ' добавляем пару ключ-значение в коллекцию
        outCol.Add valArr, key
    Next rowNum
    
    Set array2IndexedCol = outCol ' возвращаем результат
End Function
