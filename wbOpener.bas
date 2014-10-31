' возвращает ссылку на книгу по адресу
'REQUIRES: isInCollection, getFileName
Function wbOpener(wbPath As String) As Workbook
    Dim fso As New Scripting.FileSystemObject
    Dim fileName As String
    
    If fso.FileExists(wbPath) Then
        fileName = getFileName(wbPath, False)
        If isInCollection(fileName, Workbooks) Then ' уже открыт
            Set wbOpener = Workbooks.Item(fileName)
        Else
            Set wbOpener = Workbooks.Open(wbPath)
        End If
    Else
        Set wbOpener = Nothing
    End If
End Function