''' ���������� ������/��������� ������ � ���������
Function getLastCell(rng As Range) As Range
    Set tmpRng = rng.Areas(rng.Areas.Count) ' ����� ��������� ��������
    
    Set getLastCell = rng.Cells(tmpRng.Rows.Count, tmpRng.Columns.Count)
End Function

Function getFirstCell(rng As Range) As Range
    Set getFirstCell = rng.Cells(1, 1)
End Function