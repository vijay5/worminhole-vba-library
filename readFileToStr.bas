Function readFileToStr(filePath As String, Optional fso As Object = Nothing) As String
    Dim fileTs As Object ' Scripting.TextStream
    
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    If fso.FileExists(filePath) Then
        Set fileTs = fso.OpenTextFile(filePath, 1, False, 0) ' ForReading, TristateFalse
        readFileToStr = fileTs.ReadAll ' читаем файл
        fileTs.Close ' закрываем файл
    Else
        readFileToStr = ""
    End If
End Function