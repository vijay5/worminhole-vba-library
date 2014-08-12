''' выполняет хранимую процедуру (пока без входящих параметров)
Sub db_exec(conn As ADODB.Connection, queryStr As String)
    conn.Open
        conn.Execute queryStr, , 4 ' 4 =ADODB.adCmdStoredProc запуск процедуры
    conn.Close
End Sub


