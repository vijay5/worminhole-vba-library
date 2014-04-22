' ���� �� ��������� ������� � ��������� �����, ���������� ����� �������� ���������
' (�����/�����) �������� ����� �������� ������� �������
' ��������������, ��� � ��������� ����� ���������� �������,
' � �������� ������ ������ ��������� ������� ����������� ������� � ������� useColumn
' ���������� Array(-1/0/1, elemNum), ��� -1 - �������� "��", 0 - ������ ��������, +1 - �������� "�����",
' elemNum - ����� ��������
' ����� �� ����� ����� ������� ������� �� ����� ��������� (�������� �� 50000 ��������� ���� Long)
' ��� 10000: t=0.789 �� �� ����� 1 ��������
' ��� 25000: t=2.069 �� -/-
' ��� 50000: t=4.533 �� -/-
' ����� �� ���������� �������� �������, -/-
' ��� 10000: t=0.046 �� �� ������� 1 ��������
' ��� 25000: t=0.090 �� -/-
' ��� 50000: t=0.176 �� -/-
Public Function BinarySearch(arr As Collection, valueToFind As Variant, useColumn As Integer) As Variant
    Dim minRow As Long, maxRow As Long, midRow As Long
    Dim globalMin As Long, globalMax As Long
    
    globalMin = 1
    globalMax = arr.Count
    minRow = globalMin
    maxRow = globalMax

    If globalMax = 0 Then ' �� ������ ���� � ��������� �����
        BinarySearch = Array(0, 0) ' ������ ��������
        Exit Function
    End If
    
    Do
        midRow = (minRow + maxRow) \ 2 ' ������� ��������
        If valueToFind < arr(midRow)(useColumn) Then
            maxRow = midRow - 1
        ElseIf valueToFind > arr(midRow)(useColumn) Then
            minRow = midRow + 1
        Else
            BinarySearch = Array(1, midRow)
            Exit Do
        End If
        
        If (minRow > maxRow) Then ' ������� �������� �� �������
            ' ������ �������
            If minRow <= globalMax And maxRow >= globalMin Then
                If valueToFind > arr(maxRow)(useColumn) Then
                    BinarySearch = Array(1, maxRow)
                ElseIf valueToFind < arr(minRow)(useColumn) Then
                    BinarySearch = Array(-1, minRow)
                    Stop ' ��� ����� ������� �� �����������
                Else
                    Stop ' ��� ����� ������� �� �����������
                    BinarySearch = Array(1, midRow)
                End If
            Else ' �� ��������� �������
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