Function MakeRandomName(Optional num As Integer = 15) As String
    ' Aggregate - для генерации имен листов
    Dim st As String
    Dim strArray As Variant
    Dim i As Long
    
    Randomize Timer
    st = ""
    strArray = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    For i = 1 To WorksheetFunction.Min(num, 20) ' для листов это максимальная длина
        st = st + Mid(strArray, CInt(Rnd() * Len(strArray) + 1), 1)
    Next i
    MakeRandomName = st
End Function