' возвращает ссылку на книгу по адресу
'REQUIRES: isInCollection, getFileName
Function wbOpener(wbPath As String) As Workbook
    Dim fso As New Scripting.FileSystemObject
    Dim wbCur As Workbook
    Dim fileName As String
    
    Set wbCur = ActiveWorkbook
    
    If fso.FileExists(wbPath) Then
        fileName = getFileName(wbPath, False)
        If isInCollection(fileName, Workbooks) Then ' уже открыт
            Set wbOpener = Workbooks.Item(fileName)
        Else
            Application.ScreenUpdating = False
            Set wbOpener = Workbooks.Open(wbPath)
            wbCur.Activate ' чтобы не терять фокус
            Application.ScreenUpdating = True
        End If
    Else
        Set wbOpener = Nothing
    End If
End Function