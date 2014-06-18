''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'копирование матрицы
'
'Параметры:
'    A           -   матрица-источник
'    minRow, maxRow    -   диапазон строк, в которых находится подматрица-источник
'    JS1, JS2    -   диапазон столбцов, в которых находится подматрица-источник
'    B           -   матрица-приемник
'    ID1, ID2    -   диапазон строк, в которых находится подматрица-приемник
'    minCol, maxCol    -   диапазон столбцов, в которых находится подматрица-приемник
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopyMatrix(ByRef a As Variant, _
         ByVal minRowSource As Long, _
         ByVal minColSource As Long, _
         ByVal maxRowSource As Long, _
         ByVal maxColSource As Long, _
         ByRef b As Variant, _
         ByVal minRowDest As Long, _
         ByVal minColDest As Long)

    Dim minColDest As Long
    Dim maxColDest As Long
    Dim rowsCount as Long
    Dim colsCount as Long

    Dim rowNumSource As Long
    Dim colNumSource As Long

    ' дельты (сдвиг начала матрицы-получателя относительно матрицы-источника)
    columnDelta = -minColSource + minColDest
    rowDelta =  -minRowSource + minRowDest

    rowsCount = maxRowSource - minRowSource + 1
    colsCount = maxColSource - minColSource + 1

    maxRowDest = minRowDest + rowsCount - 1
    maxColDest = minColDest + colsCount - 1

    ' проверки
    If Not (IsArray(a) And IsArray(b)) Then Exit Sub
    If LBound(a, 1) < minRowSource Or UBound(a, 1) > maxRowSource Then Exit Sub
    If LBound(a, 2) < minColSource Or UBound(a, 2) > maxColSource Then Exit Sub
    If LBound(b, 1) < minRowDest Or UBound(b, 1) > maxRowDest Then Exit Sub
    If LBound(b, 2) < minColDest Or UBound(b, 2) > maxColDest Then Exit Sub

    ' перенос
    For rowNumSource = minRowSource To maxRowSource ' цикл по строкам

        rowNumDest = rowNumSource + rowDelta        ' номер строки в матрице-получателе

        For colNumSource = minColSource To maxColSource ' цикл по столбцам
            colNumDest = colNumSource + colDelta        ' номер столбца в матрице-получателе
            b(rowNumDest, colNumDest) = a(rowNumSource, colNumSource)
        Next colNumSource

    Next rowNumSource
End Sub
