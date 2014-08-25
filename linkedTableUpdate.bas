''' задёт запрос для таблицы
Sub linkedTableUpdate(sh As Worksheet, tableName As String, queryStr As String)
    Dim l As ListObject
                     
    Set l = sh.ListObjects(tableName)
                     
    l.QueryTable.CommandType = XlCmdType.xlCmdSql
    l.QueryTable.CommandText = queryStr
    
    l.QueryTable.Refresh False
    
    'this now works! ' по сути второй раз просим сервер предоставить данные
    'l.QueryTable.Refresh False
End Sub