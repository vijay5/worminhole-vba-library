''' ���������� ������ �� ����� ���� � ������
Sub db_DBToDB(srcConn As ADODB.Connection, srcTblName As String, destConn As ADODB.Connection, destTblName As String, Optional srcFldsList As Variant = "", Optional destFldsList As Variant = "", Optional srcQryCondition As String = "")

    Dim srcQueryStr As String
    Dim srcTableRs As New ADODB.Recordset
    Dim destQueryStr As String
    Dim destTableRs As New ADODB.Recordset
    
    Dim srcFieldNamesTmp As Collection
    Dim destFieldNamesTmp As Collection
    Dim srcFieldNamesArr As Variant
    Dim destFieldNamesArr As Variant
    
    Dim chk1 As Boolean
    Dim chk2 As Boolean
    Dim chk3 As Boolean
    Dim chk4 As Boolean
    Dim chk5 As Boolean
    Dim chk6 As Boolean
    Dim chk7 As Boolean
    Dim chk8 As Boolean
    
    Dim srcFieldNamesList As Collection
    Dim destFieldNamesList As Collection
    Dim value As Variant
    
    Dim el As Variant
    Dim i As Long
    Dim srcFldName As String
    Dim destFldName As String
    
    ' ///// ��������
    ' ����������� ������ ����� � ������
    If TypeName(srcFldsList) = "Collection" Then
        Set srcFieldNamesTmp = srcFldsList
        srcFieldNamesArr = col2Array(srcFieldNamesTmp) ' 0-based array
    Else
        srcFieldNamesArr = srcFldsList
    End If
    
    ' ����������� ������ ����� � ������
    If TypeName(destFldsList) = "Collection" Then
        Set destFieldNamesTmp = destFldsList
        destFieldNamesArr = col2Array(destFieldNamesTmp) ' 0-based array
    Else
        destFieldNamesArr = destFldsList
    End If
    
    chk1 = InStr(TypeName(srcFieldNamesArr), "()") > 0
    chk2 = InStr(TypeName(destFieldNamesArr), "()") > 0
    chk3 = (TypeName(srcFieldNamesArr) = "String")
    chk4 = (TypeName(destFieldNamesArr) = "String")
    
    ' ����� �������� ������ ����� ���� ��� �������, ���� ��� �� ������
    chk5 = (chk1 And chk2) Or (chk3 And chk4)
    
    If Not chk5 Then
        MsgBox "[db_DbToDb] ������ ���� ������ ��� ��� ������ �����, ��� �� ������"
        Exit Sub
    End If
    
    ' ��� �������� ��� �������� - ���������� ������������
    If chk1 And chk2 Then
        chk6 = (LBound(srcFieldNamesArr) = LBound(destFieldNamesArr))
        chk7 = (UBound(srcFieldNamesArr) = UBound(destFieldNamesArr))
    Else
        chk6 = (Len(srcFieldNamesArr) = 0)
        chk7 = (Len(destFieldNamesArr) = 0)
    End If
    
    If chk1 And chk2 And Not (chk6 And chk7) Then
        MsgBox "[db_DbToDb] ������ ����� � ��������� � ���������� ������ ��������� ���������� ����� ���������"
        Exit Sub
    End If
    
    If Not (chk1 And chk2) And Not (chk6 And chk7) Then
        MsgBox "[db_DbToDb] ������ � ����� �� ������� �����"
        Exit Sub
    End If
    
    ' \ �������� ���������

    ' ���������� � ����������
    srcQueryStr = "SELECT * FROM " + srcTblName + IIf(srcQryCondition <> "", " WHERE " + srcQryCondition, "")
    srcTableRs.ActiveConnection = srcConn.ConnectionString
    srcTableRs.LockType = adLockOptimistic
    srcTableRs.source = srcQueryStr
    
    ' ���������� � ���������
    destQueryStr = "SELECT * FROM " + destTblName
    destTableRs.ActiveConnection = destConn.ConnectionString
    destTableRs.LockType = adLockOptimistic
    destTableRs.source = destQueryStr
    
    
    
    srcTableRs.Open ' ��������� ��������-��������
    destTableRs.Open ' ��������� ��������-�������
    
    If Not (chk1 And chk2) Then ' ������ ����� �� ����� - ���������, ��� ���������� ����� � ��������� � ��������� �����
        chk8 = (srcTableRs.Fields.Count = destTableRs.Fields.Count)
    Else
        chk8 = True
    End If
    
    If Not srcTableRs.EOF Then ' ������ �� ������ ������
        srcTableRs.MoveFirst
    End If
    
    
    ' ������� ��������������� ������ ����� (�� �������� ��� �� ����� ������)
    ' ����� ��� ���������, ��� � �������� ����� ����� ���������
    If chk1 And chk2 Then ' ����� ������
        Set srcFieldNamesList = array2col(srcFieldNamesArr)
        Set destFieldNamesList = array2col(destFieldNamesArr)
    Else ' �������� �� ������� ����� �� ����� �������
        Set srcFieldNamesList = New Collection
        Set destFieldNamesList = New Collection
        
        For Each el In srcTableRs.Fields
            srcFieldNamesList.Add CStr(el.Name), CStr(el.Name)
        Next el
        
        For Each el In destTableRs.Fields
            destFieldNamesList.Add CStr(el.Name), CStr(el.Name)
        Next el
    End If
    
    Do While (Not srcTableRs.EOF) And chk8
        destTableRs.AddNew
        
        For i = 1 To srcFieldNamesList.Count ' ���������� ������ �����
            srcFldName = srcFieldNamesList.item(i)
            destFldName = destFieldNamesList.item(i)
            
            ' �������� ��������, ��� ���� ������ � ��������� � ���������� ���������
            Select Case destTableRs.Fields(destFldName).Type
            Case DataTypeEnum.adBSTR ' MEMO
                value = CStr(srcTableRs.Fields(srcFldName).value)
            Case Else
                value = srcTableRs.Fields(srcFldName).value
            End Select
            
            destTableRs.Fields(destFldName).value = value
        Next i
        
        destTableRs.Update
        srcTableRs.MoveNext
    Loop
    
    destTableRs.Close ' ��������� ��������-�������
    srcTableRs.Close ' ��������� ��������-��������
    
    

End Sub
