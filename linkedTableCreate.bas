''' создаёт таблицу, подключенную к серверу (без запроса)
Sub linkedTableCreate(destRange As Range, tableName As String, connStr As String)
    Dim l As ListObject
    Dim sh As Worksheet
    
    Set sh = destRange.Parent
    
    Set l = sh.ListObjects.Add(XlListObjectSourceType.xlSrcExternal, connStr, True, XlYesNoGuess.xlYes, destRange.Cells(1, 1))
    l.Name = tableName
End Sub

