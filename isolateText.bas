''' изолирует служебные символы, которые нельзя передавать куда-либо в строке
'REQUIRES: appendTo
Function isolateText(inString As String, Optional hex2escape As Boolean = True) As String
    Dim str As String
    Dim coll As Variant
    Dim i As Long
    
    appendTo coll, Array("\\", Chr(92)) ' 0x5C
    appendTo coll, Array("\r", Chr(13)) ' 0x0D
    appendTo coll, Array("\n", Chr(10)) ' 0x0A
    appendTo coll, Array("\t", Chr(9))  ' 0x09
    appendTo coll, Array("\f", Chr(12)) ' 0x0C
    appendTo coll, Array("\a", Chr(7))  ' 0x07
    appendTo coll, Array("\e", Chr(27)) ' 0x1B
    
    str = inString
    
    For i = LBound(coll) To UBound(coll)
        If hex2escape Then
            str = Replace(str, coll(i)(1), coll(i)(0))
        Else
            str = Replace(str, coll(i)(0), coll(i)(1))
        End If
    Next i
    
    isolateText = str
End Function
