''' ����� �����������
Sub dropAutoFilter(Optional sh As Worksheet = Nothing)
    Dim cnt As Integer
    Dim flt As Object
    
    If sh Is Nothing Then
        Set sh = ActiveSheet
    Else
        ' pass
    End If
    
    If sh.AutoFilter Is Nothing Then ' ����������� ���
    Else ' ���������� ����
        cnt = 0
        For Each flt In sh.AutoFilter.Filters
            cnt = cnt + 1
            If flt.On Then
                sh.AutoFilter.Range.AutoFilter Field:=cnt ' ���������� ������ ��� ���������� �������
            End If
        Next flt
    End If
End Sub

