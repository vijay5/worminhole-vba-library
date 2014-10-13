''' ���������� 2D-������, ���� ���� ������ � �������
''' ���� ������ ��� - ���������� ������ ������
' REQUIRES: transposeArr, col2Array, isInCollection
Function db_DBToExcel(conn As ADODB.Connection, srcTblName As String, Optional srcFieldNames As Variant = "*", Optional fieldsCol As Collection) As Variant
    Dim srcFieldNamesArr As Variant
    Dim tableRs As New ADODB.Recordset
    Dim tmp As Variant
    Dim fldName As String
    Dim fldNameAlpha As String
    Dim fldNum As Long
    Dim queryStr As String
    Dim noFldsList As Collection
    Dim existingFldsList As Collection
    Dim fldsListStr As String
    Dim srcFieldNamesTmp As Collection
    Dim i As Long
    
    ' �� ��������� - ��� ������
    db_DBToExcel = ""


    ' �������� ������������� �������� �����
    If Not IsArray(srcFieldNames) Then
        fldsListStr = "*"
    Else
        ' ����� 1 - �������� ������� ���� ����� � ��
        ' ������������� ���������� � ��
        queryStr = "SELECT * FROM " + srcTblName + ";"
        tableRs.ActiveConnection = conn.ConnectionString
        tableRs.LockType = adLockOptimistic
        tableRs.source = queryStr
        tableRs.Open ' ��������� ��������

        
        ' ����������� ������ ����� � ������
        If TypeName(srcFieldNames) = "Collection" Then
        Set srcFieldNamesTmp = srcFieldNames
            srcFieldNamesArr = col2Array(srcFieldNamesTmp) ' 0-based array
        Else
            srcFieldNamesArr = srcFieldNames
        End If
    
        ' �������� ������� ���� ����� � �������
        Set noFldsList = New Collection
        Set existingFldsList = New Collection
        
        For fldNum = LBound(srcFieldNamesArr, 1) To UBound(srcFieldNamesArr, 1) ' ������� �����
            tmp = "abrakadabra"
            On Error Resume Next
                fldName = srcFieldNamesArr(fldNum)
                ' ��� �������� ���������, ���������� ������ ��� Access...
                fldNameAlpha = Replace(Replace(fldName, "[", ""), "]", "")
                tmp = tableRs.Fields.item(fldNameAlpha).Name
            On Error GoTo 0
            If tmp = "abrakadabra" Then ' ��� ������ ����
                If Not isInCollection(fldNameAlpha, noFldsList) Then
                    noFldsList.Add fldNameAlpha
                Else
                    ' pass
                End If
            Else ' ���� ����� ����
                If Not isInCollection(fldName, existingFldsList) Then
                    existingFldsList.Add fldName
                Else
                    ' pass
                End If
            End If
        Next fldNum
    
        If noFldsList.Count > 0 Then
            MsgBox "[db_excelToDB] � ������� ����������� ����: " + Join(col2Array(noFldsList), ", ")
            Exit Function
        Else
            ' pass
        End If
    
        fldsListStr = Join(col2Array(existingFldsList), ", ")
        tableRs.Close ' ��������� ��������
    End If
    
    
    ' � ���� ����� �� ���������, ��� ��� ���� ����������
    queryStr = "SELECT " + fldsListStr + " FROM " + srcTblName + ";"
    tableRs.ActiveConnection = conn.ConnectionString
    tableRs.LockType = adLockOptimistic
    tableRs.source = queryStr
    tableRs.Open ' ��������� ��������
    
        ' ���������� ������ ����� � ���������
        Set fieldsCol = Nothing
        Set fieldsCol = New Collection
        For i = 1 To tableRs.Fields.Count
            fieldsCol.Add tableRs.Fields.item(i - 1).Name
        Next i
        
        ' ���������� ���������� ������
        If Not tableRs.EOF Then
            db_DBToExcel = transposeArr(tableRs.GetRows) ' �������������
        Else
            db_DBToExcel = ""
        End If
    tableRs.Close ' ��������� ��������
    
End Function
