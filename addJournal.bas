''' ������� ��� ������� ������� (� ����� ��� ����������� �����)
''' ����������:
'''
''' ��� ������������� ������ ����������:
'''    addJournal "[��� �������]", "[��� ������]", "����� ���������"
'''    ��� �������:
'''      ��� ������ ������� ����� ��� � ���������� funcName, ����� ������ ������� "["+funcName+"]"
'''      ��� ������ ����� ������ className, �������������� � addJournal ���������  "["+className+"."+funcName+"]"
'''    ��� ������:
'''      "[Error]" - ����������� ������, ������ ��������� ����������, ����� ����� ������� �� �����
'''      "[Event]" - ����� �������, ��� ������� � ���, �� ����� �� ���������
'''      "[Warning]" � ��� ��������� �������� - ����� ������, ��������� �� ����� ����� MsgBox, ���������� ������������
Function addJournal(ParamArray items_in() As Variant) As Variant
    Dim curTimeStr As String
    Dim journal As Variant
    Dim fso As Object, file As Object
    Dim jrnPath As Variant
    Dim wf As Variant
    Dim items As Variant
    Dim shortMessage As String, fullMessage As String
    
    ' ��������
    items = items_in ' ����� �� �� ����� �������� � �������� ��� � ��������
    
    Set wf = Application.WorksheetFunction
    
    shortMessage = ""
    If arrayLength(items) = 3 Then ' ���� �������� ��� ������� ��� ���������� ������
        If InStr(items(0), "[") = 0 And InStr(items(0), "]") = 0 Then items(0) = "[" + items(0) + "]"
    End If
    
    fullMessage = Join(items, Chr(9)) ' �� ��������� ����� ����, ����� � ������ ������ ��������
    If (UBound(items) - LBound(items) + 1) >= 3 Then
        shortMessage = items(2) ' 3� �������
    End If
    
    ' ��������� ������
    curTimeStr = wf.text(Now(), "yyyy-mm-dd hh:mm:ss")
    If ActiveSheet Is Nothing Then
        journal = curTimeStr & Chr(9) & "-" & Chr(9) & fullMessage
    Else
        journal = curTimeStr & Chr(9) & ActiveSheet.Name & Chr(9) & fullMessage
    End If
    
    ' ����� � ����� ����� ��
    Set fso = CreateObject("Scripting.FileSystemObject")
    jrnPath = ThisWorkbook.path + "\" + Mid(ThisWorkbook.Name, 1, InStrRev(ThisWorkbook.Name, ".xl") - 1) + ".journal"
    Set file = fso.OpenTextFile(jrnPath, 8, 1, 0)
    If Not file Is Nothing Then ' ���� �������� (����� �� ��������������)
        file.WriteLine journal
        file.Close
    Else ' ���� �� �������� ��� ������
        Call MsgBox("�� ������� ������� ���� ������� �� ������: " + jrnPath)
        addJournal = False
    End If
    
    If InStr(journal, "[Error]") > 0 Then ' ���� ����� �������� ������ - ��������� � ��������
        addJournal "", "[Event]", "----------== Session Terminated ==----------"
        Shell "notepad.exe " + jrnPath, vbNormalFocus
    End If
    
    If InStr(journal, "[State]") > 0 Then ' ���� ����� �������� ������ - ������� � StatusBar'�
        Application.StatusBar = CStr(shortMessage)
    End If
    
    If shortMessage <> "" And InStr(fullMessage, "[Event]") = 0 And InStr(fullMessage, "[State]") = 0 Then ' �������� ���������
        addJournal = MsgBox(shortMessage)
    End If
    
    Set file = Nothing
    Set fso = Nothing
End Function