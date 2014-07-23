Sub writeStrToFile(strToWrite As String, filePath As String, Optional append As Boolean = True, Optional createIfMissing As Boolean = True, Optional fso As Object = Nothing)
    Dim fileTs As Object ' Scripting.TextStream
    Dim ioModeLocal As Byte
    
    If append Then
        ioModeLocal = 8 ' IOMode.ForAppending
    Else
        ioModeLocal = 2 ' IOMode.ForWriting
    End If
    
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    Set fileTs = fso.OpenTextFile(filePath, ioModeLocal, createIfMissing, False)
    
    If Not fileTs Is Nothing Then
        fileTs.Write strToWrite
        fileTs.Close
    Else ' не удалось открыть файл
        ' pass
    End If
End Sub