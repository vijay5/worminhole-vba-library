''' �� ������ �� ����� ���������� �������
Function getListObjectByCell(rng As Range) As ListObject
    Dim sh As Worksheet
    Dim lstObject As ListObject
    
    
    Set getListObjectByCell = Nothing ' �������� �� ���������
    Set sh = rng.Parent ' ����, �� ������� ������ �����
    
    For Each lstObject In sh.ListObjects
        If Not Intersect(rng.Cells(1, 1), lstObject.Range) Is Nothing Then ' ���� ���� �����������
            Set getListObjectByCell = lstObject
            Exit For
        Else
            ' ��� ������
        End If
    Next lstObject
End Function