Function getUniqueValuesList(rng As Range) As Collection
    
    Dim coll As New Collection
    Dim cl As Variant
    
    For Each cl In rng
        If Not isInCollection(CStr(cl.value), coll) Then ' ��� � ��������� - ���������
            coll.Add cl.value, CStr(cl.value)
        Else
            ' pass
        End If
    Next cl
    
    Set getUniqueValuesList = coll ' ���������� ���������
End Function