' REQUIRES: isInCollection
Function substractCol(sourceCol As Collection, substrCol As Collection, Optional posIndex As Integer = -1) As Collection
    Dim el As Variant
    Dim destCol As New Collection
    Dim key As String
    
    Set destCol = sourceCol
    For Each el In substrCol ' ������� ��������� ����������� ���������
        If posIndex <> -1 Then
            key = CStr(el(posIndex)) ' ���� ��������� �������� 1D-������, ��� posIndex - ������� � ������� � ����� �����
        Else
            key = CStr(el) ' ���� � ��������� key=item
        End If
        
        If isInCollection(key, destCol) Then
            destCol.Remove key
        Else
            ' pass
        End If
    Next el
    
    Set substractCol = destCol
End Function
