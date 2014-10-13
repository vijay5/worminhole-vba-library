''' добавл€ем комментарий к каждой €чейке области targetCl
Sub addComment(targetCl As Range, comment As String, Optional append As Boolean = False)
    Dim cl As Range
    Dim tmpStr As String
    
    For Each cl In targetCl
        If cl.comment Is Nothing Then ' комментари€ нет - создаЄм
            cl.addComment comment
        Else ' комментарий есть - добавл€ем / замен€ем
            If append = True Then ' добавл€ем текст в хвост существующего комментари€
                tmpStr = cl.comment.Text
                cl.comment.Delete
                cl.addComment tmpStr + Chr(10) + comment
            Else ' замен€ем
                cl.comment.Delete
                cl.addComment comment
            End If
        End If
    Next cl

End Sub