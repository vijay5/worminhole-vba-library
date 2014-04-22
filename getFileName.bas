Function getFileName(fullPath As String, Optional getPath As Boolean = False) As String
    Dim pos As Long
    
    pos = InStrRev(fullPath, "\")
    If getPath Then
        getFileName = IIf(pos > 0, Left(fullPath, pos), fullPath) ' с первого символа до последнего слеша
    Else
        getFileName = IIf(pos > 0, Mid(fullPath, pos + 1), fullPath) ' со следующего символа после последнего слеша до конца строки
    End If
End Function