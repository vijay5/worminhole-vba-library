''' ��������� ����������� � ������ ������ ������� targetCl
Sub addComment(targetCl As Range, comment As String, Optional append As Boolean = False)
    Dim cl As Range
    
    For Each cl In targetCl
        If cl.comment Is Nothing Then ' ����������� ��� - ������
            cl.addComment comment
        Else ' ����������� ���� - ��������� / ��������
            If append = True Then ' ��������� ����� � ����� ������������� �����������
                If Len(cl.comment.Text) = 0 Then ' ������ ����������� ����, �� ������ ��� (������ �����������)
                    cl.comment.Text = comment
                Else
                    cl.comment.Text = cl.comment.Text + Chr(10) + comment
                End If
            Else ' ��������
                cl.comment.Text = comment
            End If
        End If
    Next cl

End Sub
