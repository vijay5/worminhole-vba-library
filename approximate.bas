Type point ' ���������������� ��� - �����
    x As Single
    y As Single
End Type

' ������������� ������ �����, �������������� �� ��������� �����, ���������� ����� ������ �������
Function approximate(x0 As Single, Optional points As String = "") As Variant
    Dim pointsArray() As point, tmpPoint As point, pointsArrayTmp As Variant
    Dim x1 As Single, x2 As Single
    Dim y1 As Single, y2 As Single
    Dim tmp As Variant
    Dim i As Integer, j As Integer
    Dim minX As Single, maxX As Single
    Dim numOfPoints As Integer
    
    If points = "" Then ' ������������ �������������� ����� �������� ��������� �����, �� ������� �������
        ' � �������� ������� ���� ������� x^2 �� ������� x=[0, 4]
        points = "(0,0);(0.5,0.25);(1,1);(1.5,2.25);(2,4);(2.5,6.25);(3,9);(4,16)" ' ������ ����� � ������� "(x1,y1);(x2,y2)"
    End If
    
    pointsArrayTmp = Split(points, ";")
    numOfPoints = UBound(pointsArrayTmp) - LBound(pointsArrayTmp) + 1 ' ���������� ����� � �������
    If numOfPoints < 2 Then ' ����� �� ����� ����� ������� ����� �� ������
        approximate = "#Error!"
        Exit Function
    End If
    ReDim pointsArray(LBound(pointsArrayTmp) To UBound(pointsArrayTmp))
    
    minX = 3.4E+38 ' ������������ � ����������� �������� (����� �� �������, ����� �� �����)
    maxX = -3.4E+38
    
    For i = LBound(pointsArrayTmp) To UBound(pointsArrayTmp)
        tmp = Mid(pointsArrayTmp(i), 2, Len(pointsArrayTmp(i)) - 2) ' ���� �� ����� ������
        tmp = Split(tmp, ",", 2) ' ���� ������ �� �������, �� �� ����� ��� �� 2 �����
         
         ' �������� ����� �� �����
        pointsArray(i).x = CSng(Trim(tmp(0)))  ' ����������� �� ������ � ������� �����
        pointsArray(i).y = CSng(Trim(tmp(1)))
        
        ' ���� �������/��������, ����� ����� �����, ����� ����������������
        ' ��������������� ��������� ������� � ����������� ��������
        If pointsArray(i).x > maxX Then maxX = pointsArray(i).x
        If pointsArray(i).x < minX Then minX = pointsArray(i).x
    Next i
    ' � ���� ����� � ��� ���� ���� ������ �����
    
    ' ��������� ���������
    For i = LBound(pointsArray) To UBound(pointsArray) - 1
        For j = i + 1 To UBound(pointsArray)
            If pointsArray(i).x > pointsArray(j).x Then ' ���� ����� ������� ������� - ������ �����
                tmpPoint = pointsArray(i)
                pointsArray(i) = pointsArray(j)
                pointsArray(j) = tmpPoint
            End If
        Next j
    Next i
    ' � ��� ���� ������ �����, ������������� �� X
    
    
    ' ���� ��� �������:
    ' 1 - "��" ������ (������������� �� ������ ���� ������)
    ' 2 - ������ (������������)
    ' 3 - "�����" ������ (������������� �� ��������� ���� ������)
    
    If x0 < minX Then ' �� ������
        x1 = pointsArray(0).x
        x2 = pointsArray(1).x
        y1 = pointsArray(0).y
        y2 = pointsArray(1).y
        
    ElseIf minX <= x0 And x0 <= maxX Then ' ������
        For i = LBound(pointsArray) To UBound(pointsArray) - 1
            If pointsArray(i).x <= x0 And x0 <= pointsArray(i + 1).x Then
                x1 = pointsArray(i).x
                x2 = pointsArray(i + 1).x
                y1 = pointsArray(i).y
                y2 = pointsArray(i + 1).y
                Exit For
            End If
        Next i
        
    ElseIf maxX < x0 Then ' ����� ������
        x1 = pointsArray(numOfPoints - 2).x
        x2 = pointsArray(numOfPoints - 1).x
        y1 = pointsArray(numOfPoints - 2).y
        y2 = pointsArray(numOfPoints - 1).y
    
    End If
    
    ' ������ ������������
    ' y = b + kx
    k = (y1 - y2) / (x1 - x2) ' �� �� ��� � ������� ���� = �������������� ����� � �����������
    b = ((y1 * x2) - (y2 * x1)) / (x2 - x1)
    approximate = b + k * x0 ' �������� ������� � ����� x0

End Function

