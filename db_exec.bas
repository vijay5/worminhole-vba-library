''' выполняет хранимую процедуру (пока без входящих параметров)
Sub db_exec(conn As ADODB.Connection, queryStr As String, Optional procedureType As Byte = 4)
    conn.Open
        conn.Execute queryStr, , procedureType ' 4 =ADODB.adCmdStoredProc запуск процедуры
    conn.Close
End Sub
