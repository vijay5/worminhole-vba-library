' ���������� ���������, ��� ������ �������� ���� ��� ��������� ��������,
' � ���������� - ������-������ (1D)
' ���� keyList �� ����� - ��������� ���������� array2col
' ���� ������ ���� ���������� ��� ������� (������� ������ ����� ��������� ���������� ��������)
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
        MsgBox ("[array2IndexedCol] �� ���� ���������� ������ 2D-������")
        Exit Function
    End If
    
    If IsArray(keyList) Then
        chk = True
        For i = LBound(keyList) To UBound(keyList) ' ������� ��������� ��������
            chk = chk And (keyList(i) >= LBound(arr2D, 2)) And (keyList(i) <= UBound(arr2D, 2))
        Next i
        If Not chk Then
            MsgBox ("[array2IndexedCol] ���� �� ��������� �������� ������� �� ������� �������")
            Exit Function
        End If
    End If

    For rowNum = LBound(arr2D, 1) To UBound(arr2D, 1)
        ' �������� ����
        ReDim keyArr(LBound(keyList) To UBound(keyList))
        For i = LBound(keyList) To UBound(keyList) ' ������� ��������� ��������
            keyArr(i) = CStr(arr2D(rowNum, keyList(i)))
        Next i
        key = Join(keyArr, mergeSymbol)
        
        ' �������� �������� (1D-������)
        ReDim valArr(LBound(arr2D, 2) To UBound(arr2D, 2))
        For i = LBound(arr2D, 2) To UBound(arr2D, 2)
            valArr(i) = arr2D(rowNum, i)
        Next i
        
        ' ������� �������� �� ���������, ���� ��� ��� ���� ����� ���� (���� ������ ���� ����������)
        If isInCollection(key, outCol) Then
            outCol.Remove key
        End If
        
        ' ��������� ���� ����-�������� � ���������
        outCol.Add valArr, key
    Next rowNum
    
    Set array2IndexedCol = outCol ' ���������� ���������
End Function
