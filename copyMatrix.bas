''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'����������� �������
'
'���������:
'    A           -   �������-��������
'    minRow, maxRow    -   �������� �����, � ������� ��������� ����������-��������
'    JS1, JS2    -   �������� ��������, � ������� ��������� ����������-��������
'    B           -   �������-��������
'    ID1, ID2    -   �������� �����, � ������� ��������� ����������-��������
'    minCol, maxCol    -   �������� ��������, � ������� ��������� ����������-��������
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

    ' ������ (����� ������ �������-���������� ������������ �������-���������)
    columnDelta = -minColSource + minColDest
    rowDelta =  -minRowSource + minRowDest

    rowsCount = maxRowSource - minRowSource + 1
    colsCount = maxColSource - minColSource + 1

    maxRowDest = minRowDest + rowsCount - 1
    maxColDest = minColDest + colsCount - 1

    ' ��������
    If Not (IsArray(a) And IsArray(b)) Then Exit Sub
    If LBound(a, 1) < minRowSource Or UBound(a, 1) > maxRowSource Then Exit Sub
    If LBound(a, 2) < minColSource Or UBound(a, 2) > maxColSource Then Exit Sub
    If LBound(b, 1) < minRowDest Or UBound(b, 1) > maxRowDest Then Exit Sub
    If LBound(b, 2) < minColDest Or UBound(b, 2) > maxColDest Then Exit Sub

    ' �������
    For rowNumSource = minRowSource To maxRowSource ' ���� �� �������

        rowNumDest = rowNumSource + rowDelta        ' ����� ������ � �������-����������

        For colNumSource = minColSource To maxColSource ' ���� �� ��������
            colNumDest = colNumSource + colDelta        ' ����� ������� � �������-����������
            b(rowNumDest, colNumDest) = a(rowNumSource, colNumSource)
        Next colNumSource

    Next rowNumSource
End Sub
