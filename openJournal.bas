Sub openJournal()
    Dim jrnPath As Variant
    Dim fso As Object
    

    Set fso = CreateObject("Scripting.FileSystemObject")
    jrnPath = ThisWorkbook.Path + "\" + Mid(ThisWorkbook.Name, 1, InStrRev(ThisWorkbook.Name, ".xl") - 1) + ".journal"
    If fso.FileExists(jrnPath) Then
        Shell "notepad.exe " + jrnPath, vbNormalFocus
    Else
        Call MsgBox("�� ������� ������� ���� ������� �� ������: " + jrnPath)
    End If
End Sub
