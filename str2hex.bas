''' преобразует строку в набор hex-кодов
'REQUIRES: addToText
Function str2hex(inString As Variant, Optional divisor As String = ",") As Variant
    Dim i As Long
    Dim symbol As String
    
    str2hex = ""
    
    For i = 1 To Len(inString)
        symbol = Mid(inString, i, 1)
        symbol = Hex(Asc(symbol))
        Select Case Len(symbol)
        Case 1
            symbol = "0" + symbol
        Case 3
            symbol = "0" + symbol
        Case Else
            'pass
        End Select
        symbol = "&H" + symbol
        
        addToText str2hex, symbol, divisor
    Next i
End Function
