''' ��������� ������
''' SendFromLotus sendToLogistics.GetValues = Array(), copyToLogistics.GetValues = Array(), subject + " �� ������������", letterText, attachList = array(pathToFile1, pathToFile2), False, True, True
''' REQUIRES: addToText, addJournal, arrayLength, arrayDepth
Sub SendFromLotus(strSendTo As Variant, strCopyTo As Variant, strSubject As String, strBodyText As Variant, Optional attachList As Variant, Optional sendMail As Boolean = True, Optional makeDraft As Boolean = True, Optional markRead As Boolean = True)
    '''��������� �������� ��������� ����� Lotus
    Dim session, objDatabase, objDocument, objBody As Object
    Dim userName As String, serverName As String
    Dim MailDbName As String
    Dim dbRef As Object
    Dim el As Variant
    Dim attach As Variant
    Dim funcName As String
    Dim bodyText As String
    Dim bodyTextArray As Variant
    
    funcName = "SendFromLotus"
    
    On Error Resume Next
        Set session = CreateObject("Notes.NotesSession")
        If Err.Number = -2147024894 Then
            MsgBox ("�� ������� ������� Lotus. �� ���� �� �������, ���� �� ����������." + Chr(10) + Chr(10) + "������� ����� �� ������������!")
            Exit Sub
        End If
    On Error GoTo 0
    
    userName = session.userName
    MailDbName = session.GetEnvironmentString("MailFile", True)
    serverName = session.GetEnvironmentString("MailServer", True)
    userName = session.userName ' ��� ������������
    If serverName = "" Or MailDbName = "" Or userName = "" Then ' ��������� �� �������� ������
        addJournal funcName, "[Warning]", "�� ������� ������� ���� ������. ������� � �����. �������� ����� �� ��������."
        Exit Sub
    Else
        ' pass
    End If
    Set objDatabase = session.GetDatabase(serverName, MailDbName)
    If objDatabase.IsOpen = True Then
        ' pass
    Else
        objDatabase.OPENMAIL
    End If
        
    
    Set objDocument = objDatabase.CreateDocument
    Call objDocument.ReplaceItemValue("Principal", "")         ' ?
    Call objDocument.ReplaceItemValue("From", userName)        ' ��� ����������� (����� ���� �����!!)
    Call objDocument.ReplaceItemValue("Form", "Memo")          ' ��� ���������
    Call objDocument.ReplaceItemValue("SendTo", strSendTo)     ' ����������
    Call objDocument.ReplaceItemValue("CopyTo", strCopyTo)     ' ����������
    'Call objDocument.ReplaceItemValue("Recipients", strSendTo) ' ����������
    If Not makeDraft Then
        Call objDocument.ReplaceItemValue("PostedDate", Now())     ' ���� ��������
    End If
    Call objDocument.ReplaceItemValue("Subject", strSubject)   ' ����
    Call objDocument.ReplaceItemValue("DeliveryPririty", "N")  ' �������� ��������
    Call objDocument.ReplaceItemValue("ReturnReceipt", "0")    ' ���������� � ��������
    Call objDocument.ReplaceItemValue("Importance", "0")       ' �������� ������
    
    ' ������ ���� ������
    Set objBody = objDocument.CreateRichTextItem("Body")
    bodyText = ""
    If IsArray(strBodyText) Then
        For Each el In strBodyText
            bodyText = addToText(bodyText, CStr(el), Chr(10))
        Next el
    Else
        bodyText = addToText(bodyText, CStr(strBodyText), Chr(10))
    End If
    
    If InStr(bodyText, "$files$") > 0 Then ' ����� �����, ���� ����� �������� ����� � ������ ������
        bodyTextArray = Split(bodyText, "$files$", 2)  ' ������
        Call objBody.AppendText(bodyTextArray(0) + Chr(10)) ' ����������� ������ ����� ������
    
        '����������� ��������
        If IsArray(attachList) Then
            For Each attach In attachList
                Call objBody.EmbedObject(1454, "", attach, "Attachment")
            Next attach
        Else
            If attachList <> "" Then
                Call objBody.EmbedObject(1454, "", attachList, "Attachment")
            Else
                ' ������ �� ������
            End If
        End If

        Call objBody.AppendText(Chr(10) + bodyTextArray(1)) ' ����������� ���������� ����� ������
    
    Else
        Call objBody.AppendText(bodyText) ' ����������� ����� ������
    
        '����������� ��������
        If IsArray(attachList) Then
            For Each attach In attachList
                Call objBody.EmbedObject(1454, "", attach, "Attachment")
            Next attach
        Else
            If attachList <> "" Then
                Call objBody.EmbedObject(1454, "", attachList, "Attachment")
            Else
                ' ������ �� ������
            End If
        End If
        
    End If
'    Call objBody.AppendText(el + Chr(10))
    
    
    
    Call objDocument.Save(False, False, markRead)
    If Not makeDraft Then ' �������� �� ����������
        If sendMail Then objDocument.Send False
    Else
        ' pass
    End If
    
    Set session = Nothing
    Set objDatabase = Nothing
    Set objDocument = Nothing
    Set objBody = Nothing

End Sub