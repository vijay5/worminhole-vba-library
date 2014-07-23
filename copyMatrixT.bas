''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'копирование матрицы с транспонированием
'
'Параметры:
'    A           -   матрица-источник
'    IS1, IS2    -   диапазон строк, в которых находится подматрица-источник
'    JS1, JS2    -   диапазон столбцов, в которых находится подматрица-источник
'    B           -   матрица-приемник
'    ID1, ID2    -   диапазон строк, в которых находится подматрица-приемник
'    JD1, JD2    -   диапазон столбцов, в которых находится подматрица-приемник
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopyMatrixT(ByRef a As Variant, _
         ByVal IS1 As Long, _
         ByVal IS2 As Long, _
         ByVal JS1 As Long, _
         ByVal JS2 As Long, _
         ByRef b As Variant, _
         ByVal id1 As Long, _
         ByVal ID2 As Long, _
         ByVal JD1 As Long, _
         ByVal JD2 As Long)
    Dim ISRC As Long
    Dim JDST As Long
    Dim i_ As Long
    Dim i1_ As Long

    If IS1 > IS2 Or JS1 > JS2 Then
        Exit Sub
    End If
    If Not (IsArray(a) And IsArray(b)) Then Exit Sub
    For ISRC = IS1 To IS2 Step 1
        JDST = ISRC - IS1 + JD1
        i1_ = (JS1) - (id1)
        For i_ = id1 To ID2 Step 1
            b(i_, JDST) = a(ISRC, i_ + i1_)
        Next i_
    Next ISRC
End Sub