''' ��������� �������� ��������� (���� ��� �������� ����������)
Sub db_exec(conn As ADODB.Connection, queryStr As String)
    conn.Open
        conn.Execute queryStr, , 4 ' 4 =ADODB.adCmdStoredProc ������ ���������
    conn.Close
End Sub


