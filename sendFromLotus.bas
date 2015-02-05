''' формируем письмо
''' SendFromLotus sendToLogistics.GetValues = Array(), copyToLogistics.GetValues = Array(), subject + " на согласование", letterText, attachList = array(pathToFile1, pathToFile2), False, True, True
''' REQUIRES: addToText, addJournal, arrayLength, arrayDepth
Sub SendFromLotus(strSendTo As Variant, strCopyTo As Variant, strSubject As String, strBodyText As Variant, Optional attachList As Variant, Optional sendMail As Boolean = True, Optional makeDraft As Boolean = True, Optional markRead As Boolean = True)
    '''процедура отправки сообщени€ через Lotus
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
            MsgBox ("Ќе удалось открыть Lotus. ќн либо не запущен, либо не установлен." + Chr(10) + Chr(10) + "Ўаблоны писем не сформированы!")
            Exit Sub
        End If
    On Error GoTo 0
    
    userName = session.userName
    MailDbName = session.GetEnvironmentString("MailFile", True)
    serverName = session.GetEnvironmentString("MailServer", True)
    userName = session.userName ' им€ пользовател€
    If serverName = "" Or MailDbName = "" Or userName = "" Then ' провер€ем на открытие Ћотуса
        addJournal funcName, "[Warning]", "Ќе удалось открыть базу данных. ¬ойдите в Ћотус. ќтправка писем не возможна."
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
    Call objDocument.ReplaceItemValue("From", userName)        ' им€ отправител€ (может быть любым!!)
    Call objDocument.ReplaceItemValue("Form", "Memo")          ' тип документа
    Call objDocument.ReplaceItemValue("SendTo", strSendTo)     ' получатели
    Call objDocument.ReplaceItemValue("CopyTo", strCopyTo)     ' получатели
    'Call objDocument.ReplaceItemValue("Recipients", strSendTo) ' получатели
    If Not makeDraft Then
        Call objDocument.ReplaceItemValue("PostedDate", Now())     ' дата отправки
    End If
    Call objDocument.ReplaceItemValue("Subject", strSubject)   ' тема
    Call objDocument.ReplaceItemValue("DeliveryPririty", "N")  ' скорость доставки
    Call objDocument.ReplaceItemValue("ReturnReceipt", "0")    ' оповещение о доставке
    Call objDocument.ReplaceItemValue("Importance", "0")       ' важность письма
    
    ' создаЄм тело письма
    Set objBody = objDocument.CreateRichTextItem("Body")
    bodyText = ""
    If IsArray(strBodyText) Then
        For Each el In strBodyText
            bodyText = addToText(bodyText, CStr(el), Chr(10))
        Next el
    Else
        bodyText = addToText(bodyText, CStr(strBodyText), Chr(10))
    End If
    
    If InStr(bodyText, "$files$") > 0 Then ' нашли место, куда нужно вставить файлы в тексте письма
        bodyTextArray = Split(bodyText, "$files$", 2)  ' массив
        Call objBody.AppendText(bodyTextArray(0) + Chr(10)) ' прикрепл€ем первую часть текста
    
        'ѕрикрепл€ем вложение
        If IsArray(attachList) Then
            For Each attach In attachList
                Call objBody.EmbedObject(1454, "", attach, "Attachment")
            Next attach
        Else
            If attachList <> "" Then
                Call objBody.EmbedObject(1454, "", attachList, "Attachment")
            Else
                ' ничего не делаем
            End If
        End If

        Call objBody.AppendText(Chr(10) + bodyTextArray(1)) ' прикрепл€ем оставшуюс€ часть текста
    
    Else
        Call objBody.AppendText(bodyText) ' прикрепл€ем текст письма
    
        'ѕрикрепл€ем вложение
        If IsArray(attachList) Then
            For Each attach In attachList
                Call objBody.EmbedObject(1454, "", attach, "Attachment")
            Next attach
        Else
            If attachList <> "" Then
                Call objBody.EmbedObject(1454, "", attachList, "Attachment")
            Else
                ' ничего не делаем
            End If
        End If
        
    End If
'    Call objBody.AppendText(el + Chr(10))
    
    
    
    Call objDocument.Save(False, False, markRead)
    If Not makeDraft Then ' черновик не отправл€ем
        If sendMail Then objDocument.Send False
    Else
        ' pass
    End If
    
    Set session = Nothing
    Set objDatabase = Nothing
    Set objDocument = Nothing
    Set objBody = Nothing

End Sub