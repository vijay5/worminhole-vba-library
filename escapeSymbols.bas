''' экранирует непечатные символы в строке
Function escapeSymbols(inString As String, Optional escape As Boolean = True, Optional symbList As Variant = "") As String
    Dim outString As String
    Dim defReplaceList As Variant
    Dim el As Variant
    
    defReplaceList = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)
    
    outString = inString
    If escape Then
        For Each el In defReplaceList
            outString = Replace(outString, Chr(el), "\x" & CStr(Hex(el)))
        Next el
        
        If TypeName(symbList) = "Collection" Or InStr(TypeName(symbList), "()") > 0 Then
            For Each el In symbList
                outString = Replace(outString, Chr(el), "\x" & CStr(Hex(el)))
            Next el
        End If
    Else
        For Each el In defReplaceList
            outString = Replace(outString, "\x" & CStr(Hex(el)), Chr(el))
        Next el
        
        If TypeName(symbList) = "Collection" Or InStr(TypeName(symbList), "()") > 0 Then
            For Each el In symbList
                outString = Replace(outString, "\x" & CStr(Hex(el)), Chr(el))
            Next el
        End If
    End If
    escapeSymbols = outString
End Function