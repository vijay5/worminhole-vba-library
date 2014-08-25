' для мониторинга производительности (не допилен)
Sub addTimeJournal(item As Variant, eventType As Byte, Optional groupName As String = "")
    Dim curTimeStr As String
    Dim journal As Variant
    Dim fso As Object, file As Object
    Dim jrnPath As Variant
    Dim curTimePrecise, millisecs As Single
    
    ' формируем строку
    curTimePrecise = Now()
    millisecs = Timer
    If millisecs = 0 Then
        curTimeStr = Format(curTimePrecise, "yyyy-mm-dd hh:mm:ss") + ".0"
    Else
        curTimeStr = Format(curTimePrecise, "yyyy-mm-dd hh:mm:ss") + Mid(CStr(millisecs), InStr(CStr(millisecs), "."))
    End If
    
    Select Case eventType
    Case 1 ' начало события
        journal = curTimeStr + Chr(9) + curTimeStr + Chr(9) + CStr(millisecs) + Chr(9) + "[Begin]" + Chr(9) + item
    Case 2 ' конец события
        journal = curTimeStr + Chr(9) + curTimeStr + Chr(9) + CStr(millisecs) + Chr(9) + "[End]" + Chr(9) + item
    Case Else ' просто отметка
        journal = curTimeStr + Chr(9) + curTimeStr + Chr(9) + CStr(millisecs) + Chr(9) + Chr(9) + item
    End Select

    If groupName <> "" Then journal = journal + Chr(9) + groupName
        
    ' пишем в конец файла
    Set fso = CreateObject("Scripting.FileSystemObject")
    jrnPath = ThisWorkbook.Path + "\" + Mid(ThisWorkbook.Name, 1, InStrRev(ThisWorkbook.Name, ".xl") - 1) + ".perfLog"
    
    Set file = fso.OpenTextFile(jrnPath, 8, 1, 0)
    If Not file Is Nothing Then ' если открылся (никем не модифицируется)
        file.WriteLine journal
        file.Close
    End If
    
    Set file = Nothing
    Set fso = Nothing
End Sub