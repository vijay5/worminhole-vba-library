''' ���������� ����� ����� � 2-��������� ��������� (������ �� ������ ������ ���������)
''' REQUIRES: getFirstCell, getLastCell, FindCell
Function getColumnIn2LevelHdr(targetRng As Range, hdrLevel1Str As String, hdrLevel2Str As String) As Range
    Dim destRng As Range
    Dim destRng1 As Range
    Dim colToUpdate As Range
    
    Set getColumnIn2LevelHdr = Nothing
    
    ' ���� ��������� � ������ ������ ���������
    Set destRng = Nothing
    Set destRng = FindCell(hdrLevel1Str, targetRng.Rows(1), xlValues)
    
    If Not destRng Is Nothing Then ' ����� ������
        ' �������� ����� �� ������� ������, ������� ��������� ��� ������ �������
        Set destRng1 = Range(getFirstCell(destRng.MergeArea).Offset(1, 0), _
                             getLastCell(destRng.MergeArea).Offset(1, 0))
        
        ' ������ ��� ������ ������ ���������
        Set colToUpdate = Nothing
        Set colToUpdate = FindCell(hdrLevel2Str, destRng1)
        If Not colToUpdate Is Nothing Then ' �����
            Set getColumnIn2LevelHdr = colToUpdate ' ����������
        Else
            ' pass
        End If
    Else
        ' pass
    End If
End Function
