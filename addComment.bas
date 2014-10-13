''' ��������� ����������� � ������ ������ ������� targetCl
Sub addComment(targetCl As Range, comment As String, Optional append As Boolean = False)
    Dim cl As Range
    Dim tmpStr As String
    
    For Each cl In targetCl
        If cl.comment Is Nothing Then ' ����������� ��� - ������
            cl.addComment comment
        Else ' ����������� ���� - ��������� / ��������
            If append = True Then ' ��������� ����� � ����� ������������� �����������
                tmpStr = cl.comment.Text
                cl.comment.Delete
                cl.addComment tmpStr + Chr(10) + comment
            Else ' ��������
                cl.comment.Delete
                cl.addComment comment
            End If
        End If
    Next cl

End Sub