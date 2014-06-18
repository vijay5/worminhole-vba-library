''' ���� ������� �������� � ������ �������/������
''' ���������� ������ � Range'���
'REQUIRES: appendTo
Function findSameValues(sourceRng As Range, Optional lookThroughRows As Boolean = True) As Variant
    Dim sourceRange As Range
    Dim cl As Range
    Dim list As Variant
    Dim firstValue As Range
    Dim lastValue As Range
    Dim prevValue As Range
    Dim cnt As Long
    
    If lookThroughRows Then
        Set sourceRange = sourceRng.Cells(1, 1).Resize(sourceRng.Rows.count, 1)
    Else
        Set sourceRange = sourceRng.Cells(1, 1).Resize(1, sourceRng.columns.count)
    End If
    
    list = ""
    cnt = 0
    Set firstValue = Nothing
    Set lastValue = Nothing
    Set prevValue = Nothing
    
    For Each cl In sourceRange ' ������� �����
        cnt = cnt + 1       ' ������� ����� ������
        If cnt = 1 Then     ' ��� ������ ������ � ���������
            Set firstValue = cl
            Set prevValue = cl
        End If
        If Not lastValue Is Nothing Then Set prevValue = lastValue
        Set lastValue = cl
        
        If firstValue.value <> lastValue.value Then  ' ���� ������� � ���������� �� ���������
            appendTo list, Range(firstValue, prevValue)
            
            Set firstValue = cl
            Set prevValue = cl
            Set lastValue = cl
        Else
            ' pass
        End If
        
        If cnt = sourceRange.Cells.count Then
            appendTo list, Range(firstValue, lastValue)
        End If
    Next cl
    
    findSameValues = list
End Function
