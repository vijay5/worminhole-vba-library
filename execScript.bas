' REQUIRES: addJournal
Function execScript(scriptStr As String, Optional objects As Collection = Nothing, Optional sc As Object = Nothing) As Variant
    Dim obj As Variant
    Dim objName As String
    Dim objObject As Object
    Dim tmpVal As Variant
    Dim sc1 As New MSScriptControl.ScriptControl
    
    ' �������������� ������
    If sc Is Nothing Then
        Set sc = CreateObject("MSScriptControl.ScriptControl") ' ��� ���������� ��������� ���� "1+2+Sin(1)"
        sc.Language = "VBScript"
    End If
    tmpVal = 0
    
    ' ������� ������ �� ������� (���� ��� ������)
    On Error GoTo equError1
        For Each obj In objects
            objName = obj(0) ' ��� ����������
            If IsObject(obj(1)) Then ' ���� ����� ������ - ������� ������
                Set objObject = obj(1)
                Call sc.AddObject(objName, objObject)
            Else ' ����� - ����������� �������� ����������
                Call sc.ExecuteStatement(CStr(obj(0)) & " = " & CStr(obj(1)))
            End If
equError1:
            If sc.Error.Number <> 0 Then ' ������ ��������� :)
                addJournal "[execScript]", "[Warning]", "�� ������� �������� ������ / ��������� �������� ��� ����������: " & CStr(obj(0))
            End If
            
        Next obj
    On Error GoTo 0
    
    On Error GoTo equError2
        sc.Error.Clear ' ������ ������
        tmpVal = sc.eval(scriptStr) ' ��������� ���
        
equError2:
        If sc.Error.Number <> 0 Then ' ������ ��������� :)
            addJournal "[execScript]", "[Warning]", "�� ������� ��������� �������� ���������: " & scriptStr
            tmpVal = 0 ' ��-���������, ��� ������ ����� ������� ���
        End If
    On Error GoTo 0
    
    execScript = tmpVal
End Function