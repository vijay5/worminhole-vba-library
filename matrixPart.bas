' ����� � ����� www.alglib.ru (Recoded from Fortran to VBA by Bochkanov Sergey in 2005)
'�������� ����� �������
'
'���������:
'    A                 -   �������-��������
'    MinRow, MaxRow    -   �������� �����, � ������� ��������� ����������-��������
'    MinCol, MaxCol    -   �������� ��������, � ������� ��������� ����������-��������
'    makeItSingle      -   ������� ������ ��������� (������ ���� Min=Max ����-�� �� ����� ���)
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function MatrixPart(ByRef a As Variant, _
         ByVal minRow As Long, _
         ByVal maxRow As Long, _
         Optional ByVal minCol As Long, _
         Optional ByVal MaxCol As Long, _
         Optional ByVal makeItSingle As Boolean = False, _
         Optional ByVal toDouble As Boolean = True) As Variant
    Dim tmpArray() As Double
    Dim tmpArray1() As Variant ' �� ������ ������ - ������ ��� ����� :)
    Dim dimensions As Byte
    Dim tmp As Variant

    ' ������ ����� �������� ����������� �������
    On Error Resume Next
        tmp = -1.5
        tmp = UBound(a, 1)
        If tmp <> -1.5 Then dimensions = dimensions + 1
        tmp = -1.5
        tmp = UBound(a, 2)
        If tmp <> -1.5 Then dimensions = dimensions + 1
    On Error GoTo 0

    If minRow > maxRow Or (dimensions = 2 And minCol > MaxCol) Then
        Exit Function
    End If
    
    
    If dimensions = 1 Then ' ���������� ������
        ReDim tmpArray(1 To maxRow - minRow + 1) ' ��������� ������
        ReDim tmpArray1(1 To maxRow - minRow + 1) ' ��������� ������
        For i = MaxInt(LBound(a), minRow) To MinInt(UBound(a), maxRow) ' �� ������� ������� �� ������
            If toDouble Then tmpArray(i - minRow + 1) = a(i)
            If Not toDouble Then tmpArray1(i - minRow + 1) = a(i)
        Next i
        
    ElseIf dimensions = 2 Then ' ��������� ������
        If makeItSingle And (maxRow = minRow Or MaxCol = minCol) Then  ' ���� ���� ���� � ��� ����� �������
            ReDim tmpArray(1 To maxRow - minRow + MaxCol - minCol + 1) ' ��������� ������
            ReDim tmpArray1(1 To maxRow - minRow + MaxCol - minCol + 1) ' ��������� ������
            For i = MaxInt(LBound(a, 1), minRow) To MinInt(UBound(a, 1), maxRow)
                For j = MaxInt(LBound(a, 2), minCol) To MinInt(UBound(a, 2), MaxCol)
                    If toDouble Then tmpArray(i - minRow + j - minCol + 1) = a(i, j)
                    If Not toDouble Then tmpArray1(i - minRow + j - minCol + 1) = a(i, j)
                Next j
            Next i
            
        Else ' ����� ��� ��� �� ������ ��������� ������
            ReDim tmpArray(1 To maxRow - minRow + 1, 1 To MaxCol - minCol + 1) ' ��������� ������
            ReDim tmpArray1(1 To maxRow - minRow + 1, 1 To MaxCol - minCol + 1) ' ��������� ������
        
            For i = MaxInt(LBound(a, 1), minRow) To MinInt(UBound(a, 1), maxRow)
                For j = MaxInt(LBound(a, 2), minCol) To MinInt(UBound(a, 2), MaxCol)
                    If toDouble Then tmpArray(i - minRow + 1, j - minCol + 1) = a(i, j)
                    If Not toDouble Then tmpArray1(i - minRow + 1, j - minCol + 1) = a(i, j)
                Next j
            Next i
        End If
    End If
    
    If toDouble Then MatrixPart = tmpArray
    If Not toDouble Then MatrixPart = tmpArray1
End Function