''' �� ������ �� ����� ���������� ������� �������
Function getPivotByCell(rng As Range) As PivotTable
    Dim sh As Worksheet
    Dim pivot As Variant
    
    
    Set getPivotByCell = Nothing ' �������� �� ���������
    Set sh = rng.Parent ' ����, �� ������� ������ �����
    
    For Each pivot In sh.PivotTables
        If Not Intersect(rng.Cells(1, 1), pivot.TableRange2) Is Nothing Then ' ���� ���� �����������
            Set getPivotByCell = pivot
            Exit For
        Else
            ' ��� ������
        End If
    Next pivot
End Function