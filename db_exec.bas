''' ��������� �������� ��������� (���� ��� �������� ����������)
Sub db_exec(conn As ADODB.Connection, queryStr As String, Optional procedureType As Byte = 4)
    conn.Open
        conn.Execute queryStr, , procedureType ' 4 =ADODB.adCmdStoredProc ������ ���������
    conn.Close
End Sub
