''' а-ля MAP для hexString: преобразует hex строку в строку символов, применяя к каждой подстроке chr(cdbl(hex(x)))
'REQUIRES: addJournal, addToText
Function hex2str(hexString_in As Variant, Optional divisor As String = ",") As Variant
    Dim funcName As String
    Dim isUnicode As Boolean
    Dim symbol As Variant
    Dim resultString As Variant
    Dim hexString As Variant
    Dim hexArray As Variant
    Dim hexAtom As Variant
    Dim i As Long
    
    funcName = "hex2str"
    hexString = Replace(hexString_in, "&H", "")
    hexString = Replace(hexString, "&h", "")
    hexArray = Split(hexString, divisor)
    hex2str = "" ' по умолчанию на выходе пустая строка
    
    For i = LBound(hexArray) To UBound(hexArray)
        hexAtom = hexArray(i)
        isUnicode = False
        
        Select Case Len(hexAtom) ' смотрим на длину подстроки
        Case 0
            hexAtom = "00"
        Case 1
            hexAtom = "0" + hexAtom
        Case 2
            ' всё круто
        Case 3
            isUnicode = True
            hexAtom = "0" + hexAtom ' 4-х значный код - ChrW
        Case 4
            isUnicode = True
        Case Else
            addJournal funcName, "[Warning]", "Неверное количество символов в атоме. Проверьте разделитель. Код символа может быть задан 1-4 символами. Строка: " + CStr(hexString)
            Exit Function
        End Select
        
        ' приклеиваем префикс &H
        hexAtom = "&H" + hexAtom
        
        ' проверяем подстроку на вшивость
        On Error Resume Next
            symbol = Array(0)
            If isUnicode Then
                symbol = ChrW(CLng(hexAtom)) ' символ Unicode
            Else
                symbol = Chr(CLng(hexAtom))  ' символ ASCII
            End If
        On Error GoTo 0
        If IsArray(symbol) Then
            addJournal funcName, "[Warning]", "Не удалось преобразовать подстроку """ + CStr(hexAtom) + """ в hex-символ"
            Exit Function
        End If
        
        ' собираем строку из атомов (подстрок)
        addToText resultString, symbol, ""
    Next i
    hex2str = resultString
End Function
