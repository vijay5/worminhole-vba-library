' ���������� ������ � ����������� ���������� ��� ������� �������
' selRng   - ������-������ ��� ������-�������
' useColor - True - ���������� �� ����� �������, False - ���������� �� ��������� �����
Sub mergeSameCells(selRng As Range, Optional useColor As Boolean = False)
    Dim begCl As Range
    Dim endCl As Range
    Dim prevEndCl As Range
    Dim cl as Range
    Dim cellsTotalCnt as Long
    Dim cellsCnt as Long
    Dim chk as Boolean
    
    If Not (selRng.Rows.Count = 1 Or selRng.Columns.Count = 1) Then ' ������ ������������ ������� ��������� (�����, �� ������)
        Exit Sub
    End If
    
    Set begCl = selRng.Cells(1, 1)
    Set endCl = Nothing
    
    cellsTotalCnt = selRng.Cells.Count
    cellsCnt = 0
    
    For Each cl In selRng
        cellsCnt = cellsCnt + 1 ' ������� ���������� ���������� �����
        
        Set prevEndCl = endCl   ' ���������� ������
        Set endCl = cl          ' ������� ������
        
        ' �������� �������, �� �������� ����������
        If useColor Then ' ���������� �� �����
            chk = (begCl.Interior.Color <> endCl.Interior.Color)
        Else ' ���������� �� ��������
            chk = (begCl.Value <> endCl.Value)
        End If
        
        If chk Then ' ��������� ������ ��������� ���������� �� ������� - �������� ���������� �� ���������� ������
            If Range(begCl, prevEndCl).Cells.Count > 1 Then
                tmp = Application.DisplayAlerts
                Application.DisplayAlerts = False
                Range(begCl, prevEndCl).Merge
                Application.DisplayAlerts = tmp
            End If
            Range(begCl, prevEndCl).VerticalAlignment = xlCenter
            Range(begCl, prevEndCl).HorizontalAlignment = xlCenter
            
            Set begCl = cl
        Else
            ' pass
        End If

        If cellsCnt = cellsTotalCnt Then ' �� ��������� �� ��������� ������ - ����������
            If Range(begCl, endCl).Cells.Count > 1 Then
                tmp = Application.DisplayAlerts
                Application.DisplayAlerts = False
                Range(begCl, endCl).Merge
                Application.DisplayAlerts = tmp
            End If
            Range(begCl, endCl).VerticalAlignment = xlCenter
            Range(begCl, endCl).HorizontalAlignment = xlCenter
        Else
            ' pass
        End If
    Next cl
End Sub
