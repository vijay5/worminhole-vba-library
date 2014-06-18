''' ����������� ����������� ��������� �������� �������� ����� ������������ �������� ������
''' direction - ����� ����������� ��������
''' ���� ������ ������ �� ������� �����, �� ������� ����� �� ��� - ���������� �������� ������
''' ���� ������ ������ ������ - ����������� �������� ������
Function enlargeRange(rng As Range, Optional direction As XlDirection = xlDown) As Range
    Dim resultRange As Range
    
    If rng.Cells(1, 1).value = "" Then ' ������� ������ ������
        ' pass - �����
        Set resultRange = rng.Cells(1, 1) ' ������� ������ � ����������
    Else
        Select Case direction
        Case xlDown
            If rng.Cells(1, 1).Row = rng.Parent.Cells.Rows.Count Then ' ������ � ��������� ������
                Set resultRange = rng.Cells(1, 1)
            ElseIf rng.Cells(1, 1).Offset(1, 0).value <> "" Then ' �� ��������� � �������� ���������
                Set resultRange = Range(rng.Cells(1, 1), rng.Cells(1, 1).End(xlDown))
            Else ' ���� ������� ������ ��������� � �������� ���������
                Set resultRange = rng.Cells(1, 1)
            End If
            
        Case xlUp
            If rng.Cells(1, 1).Row = 1 Then ' ������ � ������ ������
                Set resultRange = rng.Cells(1, 1)
            ElseIf rng.Cells(1, 1).Offset(-1, 0).value <> "" Then ' ������� ������ �� ���������
                Set resultRange = Range(rng.Cells(1, 1), rng.Cells(1, 1).End(xlUp))
            Else ' ���� ������� ������ ���������
                Set resultRange = rng.Cells(1, 1)
            End If
        
        Case xlToLeft
            If rng.Cells(1, 1).Column = 1 Then ' ������ � ������ �������
                Set resultRange = rng.Cells(1, 1)
            ElseIf rng.Cells(1, 1).Offset(0, -1).value <> "" Then ' �� ��������� � �������� ���������
                Set resultRange = Range(rng.Cells(1, 1), rng.Cells(1, 1).End(xlToLeft))
            Else ' ��������� � �������� ���������
                Set resultRange = rng.Cells(1, 1)
            End If
            
        Case xlToRight
            If rng.Cells(1, 1).Column = rng.Parent.Cells.Columns.Count Then ' ������ � ��������� �������
                Set resultRange = rng.Cells(1, 1)
            ElseIf rng.Cells(1, 1).Offset(0, 1).value <> "" Then ' �� ��������� � �������� ���������
                Set resultRange = Range(rng.Cells(1, 1), rng.Cells(1, 1).End(xlToRight))
            Else ' ��������� � �������� ���������
                Set resultRange = rng.Cells(1, 1)
            End If
        Case Else
            Set resultRange = rng.Cells(1, 1)
            
        End Select

    End If
    
    Set enlargeRange = resultRange
    
End Function