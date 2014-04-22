' создаЄт случайную текстовую строку
Function MakeRandomName(Optional num As Integer = 15, Optional ByVal complexity As Integer = 3) As String
    Dim st As String
    Dim strArray As Variant
    Dim i As Integer
    
    complexity = WorksheetFunction.Max(WorksheetFunction.Min(complexity, 4), 1)
    
    Randomize Timer
    st = ""
    strArray = ""
    For i = 1 To complexity
        Select Case i
        Case 1: strArray = strArray + "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        Case 2: strArray = strArray + "0123456789"
        Case 3: strArray = strArray + "абвгдеЄжзийклмнопрстуфцчшщъыьэю€јЅ¬√ƒ≈®∆«»… ЋћЌќѕ–—“”‘’÷„ЎўЏџ№Ёёя"
        Case 4: strArray = strArray + "`~!@#$%^&*()-_=+\|/,.<>[]{};:'""?"
        End Select
    Next i
    
    For i = 1 To num ' дл€ листов это максимальна€ длина
        st = st + Mid(strArray, CInt(Rnd() * Len(strArray) + 1), 1)
    Next i
    ' первый символ не может быть цифрой (дл€ имЄн таблиц, например)
    If InStr("0123456789", Left(st, 1)) > 0 Then
        st = "_" + Mid(st, 2, num - 1)
    End If
    MakeRandomName = st
End Function