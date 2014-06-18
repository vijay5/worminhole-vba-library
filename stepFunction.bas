''' ������ ����������� ������ ��������. ����� � ����� step, ����������� ������ �/��� �� ����. ����� ��������� � �������
Function stepFunction(startAt As Variant, step As Variant, Optional stopAt As Variant = "", Optional numOfPoints As Integer = -1) As Variant
    Dim curValue As Variant, curPoints As Integer
    Dim chk As Boolean
    Dim listOfValues As Variant
    
    If (stopAt = "" And numOfPoints = -1) Or step = 0 Then
        stepFunction = "!Error"
    End If
    
    curValue = startAt
    curPoints = 1
    
    ' ������� �������
        If stopAt <> "" And numOfPoints <> -1 Then         ' ������ ��� ���������
        chk = -Sgn(step) * (curValue - stopAt) >= 0 And curPoints <= numOfPoints
    ElseIf Not stopAt <> "" And numOfPoints <> -1 Then     ' ������ ���������� �����
        chk = curPoints <= numOfPoints
    ElseIf stopAt <> "" And Not numOfPoints <> -1 Then     ' ������ ���������� ��������
        chk = -Sgn(step) * (curValue - stopAt) >= 0
    ElseIf Not stopAt <> "" And Not numOfPoints <> -1 Then ' �� ����� �� ���� ��������
        ' �� ������ ����� �� ����
    End If
        
    Do While chk
        If curPoints = 1 Then ReDim listOfValues(1 To curPoints) Else ReDim Preserve listOfValues(1 To curPoints)
        listOfValues(curPoints) = curValue
        
        curValue = curValue + step
        curPoints = curPoints + 1
        
        ' ������� �������
            If stopAt <> "" And numOfPoints <> -1 Then         ' ������ ��� ���������
            chk = -Sgn(step) * (curValue - stopAt) >= 0 And curPoints <= numOfPoints
        ElseIf Not stopAt <> "" And numOfPoints <> -1 Then     ' ������ ���������� �����
            chk = curPoints <= numOfPoints
        ElseIf stopAt <> "" And Not numOfPoints <> -1 Then     ' ������ ���������� ��������
            chk = -Sgn(step) * (curValue - stopAt) >= 0
        ElseIf Not stopAt <> "" And Not numOfPoints <> -1 Then ' �� ����� �� ���� ��������
            ' �� ������ ����� �� ����
        End If
    Loop
    stepFunction = listOfValues
End Function