''' добавл€ем комментарий к каждой €чейке области targetCl
Sub addComment(targetCl As Range, comment As String, Optional append As Boolean = False)
    Dim cl As Range
    
    For Each cl In targetCl
        If cl.comment Is Nothing Then ' комментари€ нет - создаЄм
            cl.addComment comment
        Else ' комментарий есть - добавл€ем / замен€ем
            If append = True Then ' добавл€ем текст в хвост существующего комментари€
                If Len(cl.comment.Text) = 0 Then ' объект комментари€ есть, но текста нет (пустой комментарий)
                    cl.comment.Text = comment
                Else
                    cl.comment.Text = cl.comment.Text + Chr(10) + comment
                End If
            Else ' замен€ем
                cl.comment.Text = comment
            End If
        End If
    Next cl

End Sub
