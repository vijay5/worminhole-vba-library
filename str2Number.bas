''' Преобразует строку в число с проверкой на пустое значение и возможность преобразования вообще
Function str2Number(value As Variant, Optional dataType As String = "Double", Optional defValue As Variant) As Variant
    Dim defVal As Variant
    
    
    str2Number = "" ' значение по умолчанию
    
    If IsEmpty(decimalSeparator) Or IsMissing(decimalSeparator) Then getSeparators ' читаем разделители, есил они не заданы
    
    ' задаём значение по-умолчанию
    If IsMissing(defValue) Then
        defVal = "0"
    Else
        defVal = defValue
    End If
    
    ' преобразуем "пусто" в значение по-умолчанию
    If IsMissing(value) Then value = defVal
    If IsError(value) Then value = defVal
    If value = "" Then value = defVal
    
    On Error Resume Next
        Select Case LCase(dataType)
        Case "double"
            value = Replace(value, antiDecimalSeparator, decimalSeparator)
            str2Number = CDbl(value)
        Case "long"
            value = Replace(value, antiDecimalSeparator, decimalSeparator)
            str2Number = CLng(value)
        Case "single"
            value = Replace(value, antiDecimalSeparator, decimalSeparator)
            str2Number = CSng(value)
        Case "integer"
            value = Replace(value, antiDecimalSeparator, decimalSeparator)
            str2Number = CInt(value)
        Case "date"
            value = Replace(value, timeSeparator, ":")
            value = Replace(value, dateSeparator, "/")
            str2Number = CDate(value)
        End Select
        If Err.Number <> 0 Then str2Number = defVal ' затычка
    On Error GoTo 0
    
End Function
