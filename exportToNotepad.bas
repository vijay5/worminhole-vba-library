' экспортирует одномерный массив в блокнот
Sub exportToNotepad(tmpArray As Variant)
    Dim notepadID As Variant, hwnd As Long, hwndEdit As Long
    Dim buffer As DataObject ' типа буфер обмена
    Set buffer = New DataObject
    Dim tmp As Variant
    
    notepadID = Shell("notepad.exe", vbNormalFocus)
    hwnd = 0
    hwnd = FindWindowEx(0, 0, vbNullString, "Untitled - Notepad")
    If hwnd = 0 Then
        hwnd = FindWindowEx(0, 0, vbNullString, "Безымянный — Блокнот") ' в имени окна стоит длинный дефис!!!
    End If
    If hwnd <> 0 Then hwndEdit = FindWindowEx(hwnd, 0, "Edit", vbNullString)
    buffer.Clear ' через буфер обмена
    If arrayDepth(tmpArray) = 2 Then
        If arrayLength(tmpArray, 2) = 1 Then
            tmp = MatrixPart(tmpArray, LBound(tmpArray, 1), UBound(tmpArray, 1), LBound(tmpArray, 2), UBound(tmpArray, 2), True, False)
            tmp = Join(tmp, Chr(13) + Chr(10))
        Else
            Call MsgBox("Некорректное использование функции exportToNotepad - двумерный массив на входе!")
            Exit Sub
        End If
    Else
        tmp = Join(tmpArray, Chr(13) + Chr(10))
    End If
    tmp = reReplace(tmp, ";\s\>", Chr(13) & Chr(10) & ">") ' разбиваем на несколько строк
    tmp = reReplace(tmp, "(\r\n){2,}", Chr(13) + Chr(10))  ' убираем пропуски строк
    Sleep 100
    buffer.SetText (tmp)
    Sleep 100
    buffer.PutInClipboard
    If hwnd <> 0 Then
        tmp = SendMessage(hwnd, &H111, &H302, 0)  'WM_COMMAND = &H111, WM_PASTE = &H302 (Ctrl+V через SendKeys глючит)
    End If
End Sub