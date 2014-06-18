Function reReplace(ByVal text1 As Variant, ByVal patt As Variant, ByVal replaceWith As Variant) As Variant ', Optional ByVal useItemNumber As Variant = 0) As Variant ', Optional index As Integer = 0, Optional replArray As Variant) As Variant
    ' поиск регулярных выражений и замена их на что-то (как правило на "пусто")
    Dim re, Matches
    Dim leftPart, rightPart
    Dim i As Integer
    
    Set re = CreateObject("vbscript.regexp") ' позднее связывание - подключаем RegExp
    
    If text1 > "" Then ' если текст задан - иначе ругается
        With re
              .MultiLine = False
              .Global = True
              .IgnoreCase = False
              .Pattern = patt
        End With
        Set Matches = re.Execute(text1)
        For i = Matches.count To 1 Step -1 ' цикл по Match'ам
            leftPart = ""
            rightPart = ""
            leftPart = Left(text1, Matches(i - 1).FirstIndex)
            rightPart = Mid(text1, Matches(i - 1).FirstIndex + 1 + Matches(i - 1).Length)
            text1 = leftPart + CStr(replaceWith) + rightPart
        Next i
    End If
    
    reReplace = text1
    Set re = Nothing
End Function