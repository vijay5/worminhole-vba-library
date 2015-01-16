''' ������ ������ �������� �����������, ����������� ���������� ��� ���������
Function fillArray(value As Variant, ParamArray dimensions_in())

    Dim tmp As Variant
    Dim dimensions As Variant
    Dim j As Long
    Dim i1 As Long, i2 As Long, i3 As Long, i4 As Long, i5 As Long, i6 As Long, i7 As Long, i8 As Long, i9 As Long, i10 As Long
    Dim numOfDims As Long
    Dim funcName As String
    Dim el as Variant
    
    dimensions = dimensions_in
    
    funcName = "fillArray"
    
    numOfDims = arrayLength(dimensions) ' ����� ����������� ������
    ReDim Preserve dimensions(0 To 9) ' ��������� �� 9 � ���������� ������
    For j = numOfDims To 9
        dimensions(j) = 0
    Next j
    
    fillArray = "" ' �������� �� ���������
    
    Select Case numOfDims
    Case 1:  ReDim tmp(0 To dimensions(0))
    Case 2:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1))
    Case 3:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2))
    Case 4:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3))
    Case 5:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4))
    Case 6:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5))
    Case 7:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5), 0 To dimensions(6))
    Case 8:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5), 0 To dimensions(6), 0 To dimensions(7))
    Case 9:  ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5), 0 To dimensions(6), 0 To dimensions(7), 0 To dimensions(8))
    Case 10: ReDim tmp(0 To dimensions(0), 0 To dimensions(1), 0 To dimensions(2), 0 To dimensions(3), 0 To dimensions(4), 0 To dimensions(5), 0 To dimensions(6), 0 To dimensions(7), 0 To dimensions(8), 0 To dimensions(9))
    Case Else
        addJournal "funcName", "[Warning]", "����� ��������� ������ ���� ����� ������ � ��������� �� 1 �� 10"
        Exit Function
    End Select
        
    ' ���������� ��� �������� ������� 
    For Each el in tmp
        If IsObject(value) Then ' ����������� ������
            Set el = value
        Else
            el = value
        End if
    Next el
    
    fillArray = tmp
    
End Function
