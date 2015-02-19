' возвращает ссылку на книгу по адресу
'REQUIRES: isInCollection, getFileName
Function wbOpener(wbPath As String, Optional ByRef alreadyOpened As Boolean, Optional openReadOnly = False) As Workbook
    Dim fso As New Scripting.FileSystemObject
    Dim wbCur As Workbook
    Dim fileName As String
    
    Set wbCur = ActiveWorkbook
    
    If fso.FileExists(wbPath) Then
        fileName = getFileName(wbPath, False)
        If isInCollection(fileName, Workbooks) Then ' уже открыт
            Set wbOpener = Workbooks.Item(fileName)
            alreadyOpened = True

        Else
            Application.ScreenUpdating = False
            Set wbOpener = Workbooks.Open(wbPath, ReadOnly:=openReadOnly)
            wbCur.Activate ' чтобы не терять фокус
            Application.ScreenUpdating = True
            alreadyOpened = False
            
        End If
    Else
        Set wbOpener = Nothing
    End If
End Function